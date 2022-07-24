// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ResourceManagementClient } from "@azure/arm-resources";
import {
  M365TokenProvider,
  AzureAccountProvider,
  AzureSolutionSettings,
  EnvConfigFileNameTemplate,
  EnvNamePlaceholder,
  err,
  FxError,
  Json,
  ok,
  Platform,
  QTreeNode,
  Result,
  SystemError,
  TokenProvider,
  UserError,
  v2,
  v3,
  Void,
  SubscriptionInfo,
} from "@microsoft/teamsfx-api";
import { isUndefined, snakeCase } from "lodash";
import { Container } from "typedi";
import { v4 as uuidv4 } from "uuid";
import { hasAzureResource } from "../../../../common";
import { PluginDisplayName } from "../../../../common/constants";
import { getDefaultString, getLocalizedString } from "../../../../common/localizeUtils";
import { LocalStateAuthKeys, LocalStateBotKeys } from "../../../../common/localStateConstants";
import {
  CustomizeResourceGroupType,
  TelemetryEvent,
  TelemetryProperty,
} from "../../../../common/telemetry";
import { AppStudioScopes, getHashedEnv, getResourceGroupInPortal } from "../../../../common/tools";
import { convertToAlphanumericOnly } from "../../../../common/utils";
import { ComponentNames } from "../../../../component/constants";
import { AppStudioPluginV3 } from "../../../resource/appstudio/v3";
import arm, { updateResourceBaseName } from "../arm";
import { ResourceGroupInfo } from "../commonQuestions";
import {
  FillInAzureConfigsResult,
  SolutionError,
  SolutionSource,
  ProvisionSubscriptionCheckResult,
} from "../constants";
import { configLocalEnvironment, setupLocalEnvironment } from "../debug/provisionLocal";
import { resourceGroupHelper } from "../utils/ResourceGroupHelper";
import { executeConcurrently } from "../v2/executor";
import { BuiltInFeaturePluginNames } from "./constants";
import { solutionGlobalVars } from "./solutionGlobalVars";

export async function getQuestionsForProvision(
  ctx: v2.Context,
  inputs: v2.InputsWithProjectPath,
  envInfo: v2.DeepReadonly<v3.EnvInfoV3>,
  tokenProvider: TokenProvider
): Promise<Result<QTreeNode | undefined, FxError>> {
  const solutionSetting = ctx.projectSetting.solutionSettings as AzureSolutionSettings;
  const root = new QTreeNode({ type: "group" });
  for (const pluginName of solutionSetting.activeResourcePlugins) {
    const plugin = Container.get<v3.PluginV3>(pluginName);
    if (plugin.getQuestionsForProvision) {
      const res = await plugin.getQuestionsForProvision(ctx, inputs, envInfo, tokenProvider);
      if (res.isErr()) {
        return res;
      }
      if (res.value) {
        const node = res.value;
        if (node && node.data) {
          root.addChild(node);
        }
      }
    }
  }
  return ok(root);
}
export async function provisionResources(
  ctx: v2.Context,
  inputs: v2.InputsWithProjectPath,
  envInfo: v3.EnvInfoV3,
  tokenProvider: TokenProvider,
  telemetryProps?: Json
): Promise<Result<v3.EnvInfoV3, FxError>> {
  const solutionSetting = ctx.projectSetting.solutionSettings as AzureSolutionSettings | undefined;
  // 1. check M365 tenant
  if (!envInfo.state[BuiltInFeaturePluginNames.appStudio])
    envInfo.state[BuiltInFeaturePluginNames.appStudio] = {};
  const teamsAppResource = envInfo.state[
    BuiltInFeaturePluginNames.appStudio
  ] as v3.TeamsAppResource;
  if (!envInfo.state.solution) envInfo.state.solution = {};
  const solutionConfig = envInfo.state.solution as v3.AzureSolutionConfig;
  solutionConfig.provisionSucceeded = false;
  const tenantIdInConfig = teamsAppResource.tenantId;
  const tenantIdInTokenRes = await getM365TenantId(tokenProvider.m365TokenProvider);
  if (tenantIdInTokenRes.isErr()) {
    return err(tenantIdInTokenRes.error);
  }
  const tenantIdInToken = tenantIdInTokenRes.value;
  if (tenantIdInConfig && tenantIdInToken && tenantIdInToken !== tenantIdInConfig) {
    return err(
      new UserError(
        "Solution",
        SolutionError.TeamsAppTenantIdNotRight,
        `The signed in M365 account does not match the M365 tenant in config file for '${envInfo.envName}' environment. Please sign out and sign in with the correct M365 account.`
      )
    );
  }
  if (!tenantIdInConfig) {
    teamsAppResource.tenantId = tenantIdInToken;
    solutionConfig.teamsAppTenantId = tenantIdInToken;
  }

  // 2. register teams app
  const appStudioV3 = Container.get<AppStudioPluginV3>(BuiltInFeaturePluginNames.appStudio);
  const registerTeamsAppRes = await appStudioV3.registerTeamsApp(
    ctx,
    inputs,
    envInfo,
    tokenProvider
  );
  if (registerTeamsAppRes.isErr()) return err(registerTeamsAppRes.error);
  const teamsAppId = registerTeamsAppRes.value;
  teamsAppResource.teamsAppId = teamsAppId;
  solutionGlobalVars.TeamsAppId = teamsAppId;

  if (solutionSetting) {
    // 3. check Azure configs
    if (hasAzureResource(ctx.projectSetting) && envInfo.envName !== "local") {
      // ask common question and fill in solution config
      const solutionConfigRes = await fillInAzureConfigs(
        ctx,
        inputs,
        envInfo as v3.EnvInfoV3,
        tokenProvider
      );
      if (solutionConfigRes.isErr()) {
        return err(solutionConfigRes.error);
      }

      if (!solutionConfigRes.value.hasSwitchedSubscription) {
        // ask for provision consent
        const consentResult = await askForProvisionConsent(
          ctx,
          tokenProvider.azureAccountProvider,
          envInfo as v3.EnvInfoV3
        );
        if (consentResult.isErr()) {
          return err(consentResult.error);
        }
      }

      // create resource group if needed
      if (solutionConfig.needCreateResourceGroup) {
        const createRgRes = await resourceGroupHelper.createNewResourceGroup(
          solutionConfig.resourceGroupName,
          tokenProvider.azureAccountProvider,
          solutionConfig.subscriptionId,
          solutionConfig.location
        );
        if (createRgRes.isErr()) {
          return err(createRgRes.error);
        }
      }

      if (solutionConfigRes.value.hasSwitchedSubscription) {
        await updateResourceBaseName(
          inputs.projectPath,
          ctx.projectSetting.appName,
          envInfo.envName
        );
      }
    }

    // 4. collect plugins and provisionResources
    const plugins = solutionSetting.activeResourcePlugins.map((p) => Container.get<v3.PluginV3>(p));
    const provisionThunks = plugins
      .filter((plugin: v3.PluginV3) => !isUndefined(plugin.provisionResource))
      .map((plugin: v3.PluginV3) => {
        return {
          pluginName: `${plugin.name}`,
          taskName: "provisionResource",
          thunk: () => {
            if (!envInfo.state[plugin.name]) {
              envInfo.state[plugin.name] = {};
            }
            // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
            return plugin.provisionResource!(ctx, inputs, envInfo, tokenProvider);
          },
        };
      });
    ctx.logProvider.info(
      getLocalizedString("core.provision.StartNotice", PluginDisplayName.Solution)
    );
    const provisionResult = await executeConcurrently(provisionThunks, ctx.logProvider);
    if (provisionResult.kind !== "success") {
      return err(provisionResult.error);
    }

    ctx.logProvider.info(
      getLocalizedString("core.provision.ProvisionFinishNotice", PluginDisplayName.Solution)
    );

    if (envInfo.envName === "local") {
      //5.1 setup local env
      const localEnvSetupResult = await setupLocalEnvironment(ctx, inputs, envInfo);
      if (localEnvSetupResult.isErr()) {
        return err(localEnvSetupResult.error);
      }
    } else {
      //5.2 deploy arm templates for remote
      ctx.logProvider.info(
        getLocalizedString("core.deployArmTemplates.StartNotice", PluginDisplayName.Solution)
      );
      const armRes = await arm.deployArmTemplates(
        ctx,
        inputs,
        envInfo,
        tokenProvider.azureAccountProvider
      );
      if (armRes.isErr()) {
        return err(armRes.error);
      }
      ctx.logProvider.info(
        getLocalizedString("core.deployArmTemplates.SuccessNotice", PluginDisplayName.Solution)
      );
    }

    // 6. collect plugins and call configureResource
    const configureResourceThunks = plugins
      .filter((plugin: v3.PluginV3) => !isUndefined(plugin.configureResource))
      .map((plugin: v3.PluginV3) => {
        if (!envInfo.state[plugin.name]) {
          envInfo.state[plugin.name] = {};
        }
        return {
          pluginName: `${plugin.name}`,
          taskName: "configureResource",
          thunk: () =>
            // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
            plugin.configureResource!(ctx, inputs, envInfo, tokenProvider),
        };
      });
    const configureResourceResult = await executeConcurrently(
      configureResourceThunks,
      ctx.logProvider
    );
    ctx.logProvider.info(
      getLocalizedString("core.provision.configurationFinishNotice", PluginDisplayName.Solution)
    );
    const envStates = envInfo.state as v3.TeamsFxAzureResourceStates;
    if (configureResourceResult.kind !== "success") {
      const msg = getLocalizedString("core.provision.failNotice", ctx.projectSetting.appName);
      ctx.logProvider.error(msg);
      envStates.solution.provisionSucceeded = false;
      return err(configureResourceResult.error);
    }

    if (envInfo.envName === "local") {
      // 7.1 config local env
      const localConfigResult = await configLocalEnvironment(ctx, inputs, envInfo);
      if (localConfigResult.isErr()) {
        return err(localConfigResult.error);
      }
    } else {
      // 7.2 show message for remote azure provision
      const url = getResourceGroupInPortal(
        envStates.solution.subscriptionId,
        envStates.solution.tenantId,
        envStates.solution.resourceGroupName
      );
      const msg = getLocalizedString("core.provision.successAzure");
      if (url) {
        const title = "View Provisioned Resources";
        ctx.userInteraction.showMessage("info", msg, false, title).then((result: any) => {
          const userSelected = result.isOk() ? result.value : undefined;
          if (userSelected === title) {
            ctx.userInteraction.openUrl(url);
          }
        });
      } else {
        ctx.userInteraction.showMessage("info", msg, false);
      }
    }
  }
  //update Teams App
  const updateTeamsAppRes = await appStudioV3.updateTeamsApp(ctx, inputs, envInfo, tokenProvider);
  if (updateTeamsAppRes.isErr()) {
    return err(updateTeamsAppRes.error);
  }
  if (envInfo.envName !== "local") {
    const msg = getLocalizedString("core.provision.successNotice", ctx.projectSetting.appName);
    ctx.userInteraction.showMessage("info", msg, false);
    ctx.logProvider.info(msg);
  }
  solutionConfig.provisionSucceeded = true;
  return ok(envInfo);
}

/**
 * make sure subscription is correct before provision
 *
 */
export async function checkProvisionAzureSubscription(
  ctx: v2.Context,
  envInfo: v3.EnvInfoV3,
  azureAccountProvider: AzureAccountProvider
): Promise<Result<ProvisionSubscriptionCheckResult, FxError>> {
  const subscriptionIdInConfig: string | undefined = envInfo.config.azure?.subscriptionId;
  const subscriptionNameInConfig: string | undefined =
    envInfo.config.azure?.subscriptionName || subscriptionIdInConfig;
  const subscriptionIdInState: string | undefined = envInfo.state.solution.subscriptionId;
  const subscriptionNameInState: string | undefined =
    envInfo.state.solution.subscriptionName || subscriptionIdInState;

  const subscriptionInAccount = await azureAccountProvider.getSelectedSubscription(true);

  if (!subscriptionIdInState && !subscriptionIdInConfig) {
    if (!subscriptionInAccount) {
      return err(
        new UserError(
          SolutionSource,
          SolutionError.SubscriptionNotFound,
          "Failed to select subscription"
        )
      );
    } else {
      updateEnvInfoSubscription(envInfo, subscriptionInAccount);
      ctx.logProvider.info(`[${PluginDisplayName.Solution}] checkAzureSubscription pass!`);
      return ok({ hasSwitchedSubscription: false });
    }
  }

  // make sure the user is logged in
  await azureAccountProvider.getAccountCredentialAsync(true);
  // verify valid subscription (permission)
  const subscriptions = await azureAccountProvider.listSubscriptions();

  if (subscriptionIdInConfig) {
    const targetConfigSubInfo = subscriptions.find(
      (item) => item.subscriptionId === subscriptionIdInConfig
    );

    if (!targetConfigSubInfo) {
      return err(
        new UserError(
          SolutionSource,
          SolutionError.SubscriptionNotFound,
          `The subscription '${subscriptionIdInConfig}'(${subscriptionNameInConfig}) for '${
            envInfo.envName
          }' environment is not found in the current account, please use the right Azure account or check the '${EnvConfigFileNameTemplate.replace(
            EnvNamePlaceholder,
            envInfo.envName
          )}' file.`
        )
      );
    } else {
      return compareWithStateSubscription(
        ctx,
        envInfo,
        targetConfigSubInfo,
        subscriptionIdInState,
        subscriptionNameInState,
        azureAccountProvider
      );
    }
  } else {
    const targetStateSubInfo = subscriptions.find(
      (item) => item.subscriptionId === subscriptionIdInState
    );

    if (!subscriptionInAccount) {
      if (targetStateSubInfo) {
        updateEnvInfoSubscription(envInfo, targetStateSubInfo);
        ctx.logProvider.info(`[${PluginDisplayName.Solution}] checkAzureSubscription pass!`);
        return ok({ hasSwitchedSubscription: false });
      } else {
        return err(
          new UserError(
            SolutionSource,
            SolutionError.SubscriptionNotFound,
            `The subscription '${subscriptionIdInState}'(${subscriptionNameInState}) for '${envInfo.envName}' environment is not found in the current account, please use the right Azure account.`
          )
        );
      }
    } else {
      return compareWithStateSubscription(
        ctx,
        envInfo,
        subscriptionInAccount,
        subscriptionIdInState,
        subscriptionNameInState,
        azureAccountProvider
      );
    }
  }
}

function updateEnvInfoSubscription(envInfo: v3.EnvInfoV3, subscriptionInfo: SubscriptionInfo) {
  envInfo.state.solution.subscriptionId = subscriptionInfo.subscriptionId;
  envInfo.state.solution.subscriptionName = subscriptionInfo.subscriptionName;
  envInfo.state.solution.tenantId = subscriptionInfo.tenantId;
}

async function compareWithStateSubscription(
  ctx: v2.Context,
  envInfo: v3.EnvInfoV3,
  targetSubscriptionInfo: SubscriptionInfo,
  subscriptionInStateId: string | undefined,
  subscriptionInStateName: string | undefined,
  azureAccountProvider: AzureAccountProvider
): Promise<Result<ProvisionSubscriptionCheckResult, FxError>> {
  const shouldAskForSubscriptionConfirmation =
    !!subscriptionInStateId && targetSubscriptionInfo.subscriptionId !== subscriptionInStateId;
  if (shouldAskForSubscriptionConfirmation) {
    const confirmResult = await askForSubscriptionConfirm(
      ctx,
      subscriptionInStateName!,
      targetSubscriptionInfo.subscriptionName || targetSubscriptionInfo.subscriptionId,
      azureAccountProvider,
      envInfo
    );
    if (confirmResult.isErr()) {
      return err(confirmResult.error);
    } else {
      updateEnvInfoSubscription(envInfo, targetSubscriptionInfo);
      clearEnvInfoStateResource(envInfo);

      ctx.logProvider.info(`[${PluginDisplayName.Solution}] checkAzureSubscription pass!`);
      return ok({ hasSwitchedSubscription: true });
    }
  } else {
    updateEnvInfoSubscription(envInfo, targetSubscriptionInfo);
    ctx.logProvider.info(`[${PluginDisplayName.Solution}] checkAzureSubscription pass!`);
    return ok({ hasSwitchedSubscription: false });
  }
}

// clear resources related info in envInfo so that we could provision successfully using new sub.
function clearEnvInfoStateResource(envInfo: v3.EnvInfoV3): void {
  envInfo.state.solution.resourceGroupName = "";
  envInfo.state.solution.resourceNameSuffix = "";

  const keysToClear = [
    BuiltInFeaturePluginNames.bot,
    BuiltInFeaturePluginNames.frontend,
    BuiltInFeaturePluginNames.apim,
    BuiltInFeaturePluginNames.function,
    BuiltInFeaturePluginNames.identity,
    BuiltInFeaturePluginNames.keyVault,
    BuiltInFeaturePluginNames.sql,
    BuiltInFeaturePluginNames.simpleAuth,
    ComponentNames.TeamsBot,
    ComponentNames.TeamsTab,
    ComponentNames.APIM,
    ComponentNames.TeamsApi,
    ComponentNames.Identity,
    ComponentNames.KeyVault,
    ComponentNames.AzureSQL,
  ];
  const keys = Object.keys(envInfo.state);
  for (let index = 0; index < keys.length; index++) {
    if (keysToClear.includes(keys[index])) {
      envInfo.state[keys[index]] = {};
    }
  }
}

async function askForSubscriptionConfirm(
  ctx: v2.Context,
  subscriptionInState: string,
  subscriptionInAccount: string,
  azureAccountProvider: AzureAccountProvider,
  envInfo: v3.EnvInfoV3
): Promise<Result<Void, FxError>> {
  const azureToken = await azureAccountProvider.getAccountCredentialAsync();

  const username = (azureToken as any).username || "";
  const msgNew = getLocalizedString(
    "core.provision.confirmSubscription",
    subscriptionInState,
    envInfo.envName,
    username,
    subscriptionInAccount
  );
  const confirmRes = await ctx.userInteraction.showMessage("warn", msgNew, true, "Provision");
  const confirm = confirmRes?.isOk() ? confirmRes.value : undefined;

  if (confirm !== "Provision") {
    return err(new UserError(SolutionSource, "CancelProvision", "CancelProvision"));
  }
  return ok(Void);
}

/**
 * Asks common questions and puts the answers in the global namespace of SolutionConfig
 *
 */
export async function fillInAzureConfigs(
  ctx: v2.Context,
  inputs: v2.InputsWithProjectPath,
  envInfo: v3.EnvInfoV3,
  tokenProvider: TokenProvider
): Promise<Result<FillInAzureConfigsResult, FxError>> {
  //1. check subscriptionId
  const subscriptionResult = await checkProvisionAzureSubscription(
    ctx,
    envInfo,
    tokenProvider.azureAccountProvider
  );
  if (subscriptionResult.isErr()) {
    return err(subscriptionResult.error);
  }

  // Note setSubscription here will change the token returned by getAccountCredentialAsync according to the subscription selected.
  // So getting azureToken needs to precede setSubscription.
  const azureToken = await tokenProvider.azureAccountProvider.getAccountCredentialAsync();
  if (azureToken === undefined) {
    return err(
      new UserError(
        SolutionSource,
        SolutionError.NotLoginToAzure,
        "Login to Azure using the Azure Account extension"
      )
    );
  }

  //2. check resource group
  ctx.telemetryReporter?.sendTelemetryEvent(
    TelemetryEvent.CheckResourceGroupStart,
    inputs.env ? { [TelemetryProperty.Env]: getHashedEnv(inputs.env) } : {}
  );

  const rmClient = new ResourceManagementClient(azureToken, envInfo.state.solution.subscriptionId);

  // Resource group info precedence are:
  //   0. ctx.answers, for VS targetResourceGroupName and targetResourceLocationName to create a new rg
  //   1. ctx.answers, for CLI --resource-group argument, only support existing resource group
  //   2. env config (config.{envName}.json), for user customization, only support existing resource group
  //   3. states (state.{envName}.json), for re-provision
  //   4. asking user with a popup
  const resourceGroupNameFromEnvConfig = envInfo.config.azure?.resourceGroupName;
  const resourceGroupNameFromState = envInfo.state.solution.resourceGroupName;
  const resourceGroupLocationFromState = envInfo.state.solution.location;
  const appName = convertToAlphanumericOnly(ctx.projectSetting.appName);
  const defaultResourceGroupName = `${snakeCase(appName)}${"-" + envInfo.envName}-rg`;
  let resourceGroupInfo: ResourceGroupInfo;
  const telemetryProperties: { [key: string]: string } = {};
  if (inputs.env) {
    telemetryProperties[TelemetryProperty.Env] = getHashedEnv(inputs.env);
  }

  if (inputs.targetResourceGroupName) {
    const getRes = await resourceGroupHelper.getResourceGroupInfo(
      inputs.targetResourceGroupName,
      rmClient
    );
    if (getRes.isErr()) {
      // support vs to create a new resource group
      if (inputs.platform === Platform.VS && inputs.targetResourceLocationName) {
        resourceGroupInfo = {
          createNewResourceGroup: true,
          name: inputs.targetResourceGroupName,
          location: inputs.targetResourceLocationName,
        };
      } else return err(getRes.error);
    } else {
      if (!getRes.value) {
        // Currently we do not support creating resource group from command line arguments
        return err(
          new UserError(
            SolutionSource,
            SolutionError.ResourceGroupNotFound,
            `Resource group '${inputs.targetResourceGroupName}' does not exist, please specify an existing resource group.`
          )
        );
      }
      telemetryProperties[TelemetryProperty.CustomizeResourceGroupType] =
        CustomizeResourceGroupType.CommandLine;
      resourceGroupInfo = getRes.value;
    }
  } else if (resourceGroupNameFromEnvConfig) {
    const resourceGroupName = resourceGroupNameFromEnvConfig;
    const getRes = await resourceGroupHelper.getResourceGroupInfo(resourceGroupName, rmClient);
    if (getRes.isErr()) return err(getRes.error);
    if (!getRes.value) {
      // Currently we do not support creating resource group by input config, so just throw an error.
      const envFile = EnvConfigFileNameTemplate.replace(EnvNamePlaceholder, inputs.envName);
      return err(
        new UserError(
          SolutionSource,
          SolutionError.ResourceGroupNotFound,
          `Resource group '${resourceGroupName}' does not exist, please check your '${envFile}' file.`
        )
      );
    }
    telemetryProperties[TelemetryProperty.CustomizeResourceGroupType] =
      CustomizeResourceGroupType.EnvConfig;
    resourceGroupInfo = getRes.value;
  } else if (resourceGroupNameFromState && resourceGroupLocationFromState) {
    const checkRes = await resourceGroupHelper.checkResourceGroupExistence(
      resourceGroupNameFromState,
      rmClient
    );
    if (checkRes.isErr()) {
      return err(checkRes.error);
    }
    const exist = checkRes.value;
    resourceGroupInfo = {
      createNewResourceGroup: !exist,
      name: resourceGroupNameFromState,
      location: resourceGroupLocationFromState,
    };
    telemetryProperties[TelemetryProperty.CustomizeResourceGroupType] =
      CustomizeResourceGroupType.EnvState;
  } else {
    const resourceGroupInfoResult = await resourceGroupHelper.askResourceGroupInfo(
      ctx,
      inputs,
      tokenProvider.azureAccountProvider,
      rmClient,
      defaultResourceGroupName
    );
    if (resourceGroupInfoResult.isErr()) {
      return err(resourceGroupInfoResult.error);
    }

    resourceGroupInfo = resourceGroupInfoResult.value;
    if (resourceGroupInfo.createNewResourceGroup) {
      if (resourceGroupInfo.name === defaultResourceGroupName) {
        telemetryProperties[TelemetryProperty.CustomizeResourceGroupType] =
          CustomizeResourceGroupType.InteractiveCreateDefault;
      } else {
        telemetryProperties[TelemetryProperty.CustomizeResourceGroupType] =
          CustomizeResourceGroupType.InteractiveCreateCustomized;
      }
    } else {
      telemetryProperties[TelemetryProperty.CustomizeResourceGroupType] =
        CustomizeResourceGroupType.InteractiveUseExisting;
    }
  }

  ctx.telemetryReporter?.sendTelemetryEvent(TelemetryEvent.CheckResourceGroup, telemetryProperties);

  envInfo.state.solution.needCreateResourceGroup = resourceGroupInfo.createNewResourceGroup;
  envInfo.state.solution.resourceGroupName = resourceGroupInfo.name;
  envInfo.state.solution.location = resourceGroupInfo.location;
  ctx.logProvider?.info(`[${PluginDisplayName.Solution}] check resource group pass!`);
  ctx.logProvider?.info(`[${PluginDisplayName.Solution}] check teamsAppTenantId pass!`);

  //resourceNameSuffix
  const resourceNameSuffix =
    (envInfo.config.azure?.resourceNameSuffix as string) ||
    envInfo.state.solution.resourceNameSuffix ||
    uuidv4().substr(0, 6);
  envInfo.state.solution.resourceNameSuffix = resourceNameSuffix;
  ctx.logProvider?.info(`[${PluginDisplayName.Solution}] check resourceNameSuffix pass!`);
  return ok({ hasSwitchedSubscription: subscriptionResult.value.hasSwitchedSubscription });
}

export async function askForDeployConsent(
  ctx: v2.Context,
  azureAccountProvider: AzureAccountProvider,
  envInfo: v3.EnvInfoV3
): Promise<Result<Void, FxError>> {
  const azureToken = await azureAccountProvider.getAccountCredentialAsync();

  // Only Azure project requires this confirm dialog
  const username = (azureToken as any).username || "";
  const subscriptionId = envInfo.state.solution?.subscriptionId || "";
  const subscriptionName = envInfo.state.solution?.subscriptionName || "";
  const msg = getLocalizedString(
    "core.deploy.confirmEnvNotice",
    envInfo.envName,
    username,
    subscriptionName ? subscriptionName : subscriptionId
  );
  const deployOption = "Deploy";
  const result = await ctx.userInteraction.showMessage("warn", msg, true, deployOption);
  const choice = result?.isOk() ? result.value : undefined;

  if (choice === deployOption) {
    return ok(Void);
  }
  return err(new UserError(SolutionSource, "UserCancel", "UserCancel"));
}

export async function askForProvisionConsent(
  ctx: v2.Context,
  azureAccountProvider: AzureAccountProvider,
  envInfo: v3.EnvInfoV3
): Promise<Result<Void, FxError>> {
  const azureToken = await azureAccountProvider.getAccountCredentialAsync();

  // Only Azure project requires this confirm dialog
  const username = (azureToken as any).username || "";
  const subscriptionId = envInfo.state.solution?.subscriptionId || "";
  const subscriptionName = envInfo.state.solution?.subscriptionName || "";
  const msgNew = getLocalizedString(
    "core.provision.confirmEnvNotice",
    envInfo.envName,
    username,
    subscriptionName ? subscriptionName : subscriptionId
  );
  const confirmRes = await ctx.userInteraction.showMessage("warn", msgNew, true, "Provision");
  const confirm = confirmRes?.isOk() ? confirmRes.value : undefined;

  if (confirm !== "Provision") {
    if (confirm === "Pricing calculator") {
      ctx.userInteraction.openUrl("https://azure.microsoft.com/en-us/pricing/calculator/");
    }
    return err(new UserError(SolutionSource, "CancelProvision", "CancelProvision"));
  }
  return ok(Void);
}

export async function getM365TenantId(
  m365TokenProvider: M365TokenProvider
): Promise<Result<string, FxError>> {
  // Just to trigger M365 login before the concurrent execution of localDebug.
  // Because concurrent execution of localDebug may getAccessToken() concurrently, which
  // causes 2 M365 logins before the token caching in common lib takes effect.
  const appStudioTokenRes = await m365TokenProvider.getAccessToken({ scopes: AppStudioScopes });
  if (appStudioTokenRes.isErr()) {
    return err(appStudioTokenRes.error);
  }
  const appStudioTokenJsonRes = await m365TokenProvider.getJsonObject({ scopes: AppStudioScopes });
  const appStudioTokenJson = appStudioTokenJsonRes.isOk() ? appStudioTokenJsonRes.value : undefined;
  if (appStudioTokenJson === undefined) {
    return err(
      new SystemError(
        SolutionSource,
        SolutionError.NoAppStudioToken,
        getDefaultString("error.NoAppStudioToken"),
        getLocalizedString("error.NoAppStudioToken")
      )
    );
  }
  const tenantIdInToken = (appStudioTokenJson as any).tid;
  if (!tenantIdInToken || !(typeof tenantIdInToken === "string")) {
    return err(
      new SystemError(
        SolutionSource,
        SolutionError.NoTeamsAppTenantId,
        getDefaultString("error.NoTeamsAppTenantId"),
        getLocalizedString("error.NoTeamsAppTenantId")
      )
    );
  }
  return ok(tenantIdInToken);
}
