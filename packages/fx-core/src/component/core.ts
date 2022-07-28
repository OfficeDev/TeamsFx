// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Action,
  CloudResource,
  ConfigFolderName,
  ContextV3,
  err,
  FunctionAction,
  FxError,
  InputsWithProjectPath,
  MaybePromise,
  ok,
  Platform,
  ProjectSettingsV3,
  ResourceContextV3,
  Result,
  UserError,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import path from "path";
import "reflect-metadata";
import { Container, Service } from "typedi";
import {
  CoreQuestionNames,
  ProjectNamePattern,
  QuestionRootFolder,
  ScratchOptionNo,
} from "../core/question";
import { isVSProject, newProjectSettings } from "./../common/projectSettingsHelper";
import "./bicep";
import "./code/apiCode";
import "./code/botCode";
import "./code/spfxTabCode";
import "./code/tabCode";
import "./connection/apimConfig";
import "./connection/azureFunctionConfig";
import "./connection/azureWebAppConfig";
import { configLocalEnvironment, setupLocalEnvironment } from "./debug";
import { createNewEnv } from "./envManager";
import "./feature/api";
import "./feature/apiConnector";
import "./feature/apim";
import "./feature/bot";
import "./feature/cicd";
import "./feature/keyVault";
import "./feature/spfx";
import "./feature/sql";
import "./feature/sso";
import "./feature/tab";
import "./resource/aadApp/aadApp";
import "./resource/apim";
import { AppManifest } from "./resource/appManifest/appManifest";
import "./resource/azureAppService/azureFunction";
import "./resource/azureAppService/azureWebApp";
import "./resource/azureSql";
import "./resource/azureStorage";
import "./resource/botService";
import "./resource/keyVault";
import "./resource/spfx";

import { AADApp } from "@microsoft/teamsfx-api/build/v3";
import * as jsonschema from "jsonschema";
import { cloneDeep } from "lodash";
import { PluginDisplayName } from "../common/constants";
import { globalStateUpdate } from "../common/globalState";
import { getDefaultString, getLocalizedString } from "../common/localizeUtils";
import { hasAAD, hasAzureResourceV3, hasBot } from "../common/projectSettingsHelperV3";
import { getResourceGroupInPortal } from "../common/tools";
import { downloadSample } from "../core/downloadSample";
import { InvalidInputError } from "../core/error";
import { globalVars } from "../core/globalVars";
import { getQuestionsForCreateProjectV2 } from "../core/middleware/questionModel";
import arm, { updateResourceBaseName } from "../plugins/solution/fx-solution/arm";
import {
  ApiConnectionOptionItem,
  AzureResourceApim,
  AzureResourceFunctionNewUI,
  AzureResourceKeyVaultNewUI,
  AzureResourceSQLNewUI,
  AzureSolutionQuestionNames,
  BotFeatureIds,
  CicdOptionItem,
  M365SearchAppOptionItem,
  M365SsoLaunchPageOptionItem,
  SingleSignOnOptionItem,
  TabFeatureIds,
  TabSPFxItem,
} from "../plugins/solution/fx-solution/question";
import { resourceGroupHelper } from "../plugins/solution/fx-solution/utils/ResourceGroupHelper";
import { executeConcurrently } from "../plugins/solution/fx-solution/v2/executor";
import {
  checkWhetherLocalDebugM365TenantMatches,
  getBotTroubleShootMessage,
} from "../plugins/solution/fx-solution/v2/utils";
import { checkDeployAzureSubscription } from "../plugins/solution/fx-solution/v3/deploy";
import {
  askForDeployConsent,
  askForProvisionConsent,
  fillInAzureConfigs,
  getM365TenantId,
} from "../plugins/solution/fx-solution/v3/provision";
import { AzureResources, ComponentNames } from "./constants";
import { pluginName2ComponentName } from "./migrate";
import {
  getQuestionsForAddFeatureV3,
  getQuestionsForDeployV3,
  getQuestionsForProvisionV3,
} from "./questionV3";
import { runActionByName } from "./workflow";
@Service("fx")
export class TeamsfxCore {
  name = "fx";

  /**
   * create project
   */
  create(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): MaybePromise<Result<Action | undefined, FxError>> {
    const createAction: FunctionAction = {
      name: "fx.create",
      type: "function",
      question: async (context, inputs) => {
        return await getQuestionsForCreateProjectV2(inputs);
      },
      execute: async (context, inputs) => {
        const folder = inputs[QuestionRootFolder.name] as string;
        if (!folder) {
          return err(InvalidInputError("folder is undefined"));
        }
        inputs.folder = folder;
        const scratch = inputs[CoreQuestionNames.CreateFromScratch] as string;
        let projectPath: string;
        const automaticNpmInstall = "automaticNpmInstall";
        if (scratch === ScratchOptionNo.id) {
          // create from sample
          const downloadRes = await downloadSample(inputs);
          if (downloadRes.isErr()) {
            return err(downloadRes.error);
          }
          projectPath = downloadRes.value;
        } else {
          // create from new
          const appName = inputs[CoreQuestionNames.AppName] as string;
          if (undefined === appName) return err(InvalidInputError(`App Name is empty`, inputs));
          const validateResult = jsonschema.validate(appName, {
            pattern: ProjectNamePattern,
          });
          if (validateResult.errors && validateResult.errors.length > 0) {
            return err(InvalidInputError(`${validateResult.errors[0].message}`, inputs));
          }
          projectPath = path.join(folder, appName);
          inputs.projectPath = projectPath;
          const initRes = await this.init(context, inputs);
          if (initRes.isErr()) return err(initRes.error);
          const features = inputs.capabilities;
          delete inputs.folder;
          if (
            features === M365SsoLaunchPageOptionItem.id ||
            features === M365SearchAppOptionItem.id
          ) {
            context.projectSetting.isM365 = true;
            inputs.isM365 = true;
          }
          if (BotFeatureIds.includes(features)) {
            inputs[AzureSolutionQuestionNames.Features] = features;
            const res = await runActionByName("teams-bot.add", context, inputs);
            if (res.isErr()) return err(res.error);
          }
          if (TabFeatureIds.includes(features)) {
            inputs[AzureSolutionQuestionNames.Features] = features;
            const res = await runActionByName("teams-tab.add", context, inputs);
            if (res.isErr()) return err(res.error);
          }
          if (features === TabSPFxItem.id) {
            inputs[AzureSolutionQuestionNames.Features] = features;
            const res = await runActionByName("spfx-tab.add", context, inputs);
            if (res.isErr()) return err(res.error);
          }
        }
        if (inputs.platform === Platform.VSCode) {
          await globalStateUpdate(automaticNpmInstall, true);
        }
        context.projectPath = projectPath;
        return ok(["create a new project with capability:" + inputs.capabilities]);
      },
    };

    return ok(createAction);
  }
  /**
   * add feature
   */
  addFeature(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): MaybePromise<Result<Action | undefined, FxError>> {
    const action: FunctionAction = {
      type: "function",
      name: "fx.addFeature",
      question: async (context: ContextV3, inputs: InputsWithProjectPath) => {
        return await getQuestionsForAddFeatureV3(context, inputs);
      },
      execute: async (context, inputs) => {
        const features = inputs[AzureSolutionQuestionNames.Features];
        let actionName;
        if (BotFeatureIds.includes(features)) {
          actionName = "teams-bot.add";
        } else if (TabFeatureIds.includes(features)) {
          actionName = "teams-tab.add";
        } else if (features === AzureResourceSQLNewUI.id) {
          actionName = "sql.add";
        } else if (features === AzureResourceFunctionNewUI.id) {
          actionName = "teams-api.add";
        } else if (features === AzureResourceApim.id) {
          actionName = "apim-feature.add";
        } else if (features === AzureResourceKeyVaultNewUI.id) {
          actionName = "key-vault-feature.add";
        } else if (features === CicdOptionItem.id) {
          actionName = "cicd.add";
        } else if (features === ApiConnectionOptionItem.id) {
          actionName = "api-connector.add";
        } else if (features === SingleSignOnOptionItem.id) {
          actionName = "sso.add";
        }
        if (actionName) {
          const res = await runActionByName(actionName, context, inputs);
          if (res.isErr()) return err(res.error);
        }
        return ok([]);
      },
    };
    return ok(action);
  }
  async init(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<undefined, FxError>> {
    const projectSettings = newProjectSettings() as ProjectSettingsV3;
    projectSettings.appName = inputs["app-name"];
    projectSettings.components = [];
    context.projectSetting = projectSettings;
    await fs.ensureDir(inputs.projectPath);
    await fs.ensureDir(path.join(inputs.projectPath, `.${ConfigFolderName}`));
    await fs.ensureDir(path.join(inputs.projectPath, `.${ConfigFolderName}`, "configs"));
    {
      const appManifest = Container.get<AppManifest>(ComponentNames.AppManifest);
      const res = await appManifest.init(context, inputs);
      if (res.isErr()) return res;
    }
    {
      const res = await createNewEnv(context, inputs);
      if (res.isErr()) return res;
    }
    return ok(undefined);
  }
  provision(
    ctx: ContextV3,
    inputs: InputsWithProjectPath
  ): MaybePromise<Result<Action | undefined, FxError>> {
    const action: FunctionAction = {
      name: "fx.provision",
      type: "function",
      question: async (context: ContextV3, inputs: InputsWithProjectPath) => {
        return await getQuestionsForProvisionV3(context, context.envInfo!, inputs);
      },
      execute: async (context, inputs) => {
        const ctx = context as ResourceContextV3;
        ctx.envInfo.state.solution = ctx.envInfo.state.solution || {};
        ctx.envInfo.state.solution.provisionSucceeded = false;

        // 1. pre provision
        {
          const res = await preProvision(ctx, inputs);
          if (res.isErr()) return err(res.error);
        }
        // 2. create a teams app
        const appManifest = Container.get<AppManifest>(ComponentNames.AppManifest);
        {
          const res = await appManifest.provision(ctx, inputs);
          if (res.isErr()) return err(res.error);
        }

        // 3. call resources provision api
        const componentsToProvision = ctx.projectSetting.components.filter((r) => r.provision);
        {
          const thunks = [];
          for (const componentConfig of componentsToProvision) {
            const componentInstance = Container.get<CloudResource>(componentConfig.name);
            if (componentInstance.provision) {
              thunks.push({
                pluginName: `${componentConfig.name}`,
                taskName: "provision",
                thunk: () => {
                  ctx.envInfo.state[componentConfig.name] =
                    ctx.envInfo.state[componentConfig.name] || {};
                  return componentInstance.provision!(ctx, inputs);
                },
              });
            }
          }
          const provisionResult = await executeConcurrently(thunks, ctx.logProvider);
          if (provisionResult.kind !== "success") {
            return err(provisionResult.error);
          }
          ctx.logProvider.info(
            getLocalizedString("core.provision.ProvisionFinishNotice", PluginDisplayName.Solution)
          );
        }

        // 4
        if (ctx.envInfo.envName === "local") {
          //4.1 setup local env
          const localEnvSetupResult = await setupLocalEnvironment(ctx, inputs, ctx.envInfo);
          if (localEnvSetupResult.isErr()) {
            return err(localEnvSetupResult.error);
          }
        } else {
          //4.2 deploy arm templates for remote
          ctx.logProvider.info(
            getLocalizedString("core.deployArmTemplates.StartNotice", PluginDisplayName.Solution)
          );
          const armRes = await arm.deployArmTemplates(
            ctx,
            inputs,
            ctx.envInfo,
            ctx.tokenProvider.azureAccountProvider
          );
          if (armRes.isErr()) {
            return err(armRes.error);
          }
          ctx.logProvider.info(
            getLocalizedString("core.deployArmTemplates.SuccessNotice", PluginDisplayName.Solution)
          );
        }

        // 5.0 "aad-app.setApplicationInContext"
        const aadApp = Container.get<AADApp>(ComponentNames.AadApp);
        if (hasAAD(ctx.projectSetting)) {
          const res = await aadApp.setApplicationInContext(ctx, inputs);
          if (res.isErr()) return err(res.error);
        }
        // 5. call resources configure api
        {
          const thunks = [];
          for (const componentConfig of componentsToProvision) {
            const componentInstance = Container.get<CloudResource>(componentConfig.name);
            if (componentInstance.configure) {
              thunks.push({
                pluginName: `${componentConfig.name}`,
                taskName: "configure",
                thunk: () => {
                  ctx.envInfo.state[componentConfig.name] =
                    ctx.envInfo.state[componentConfig.name] || {};
                  return componentInstance.configure!(ctx, inputs);
                },
              });
            }
          }
          const configResult = await executeConcurrently(thunks, ctx.logProvider);
          if (configResult.kind !== "success") {
            return err(configResult.error);
          }
          ctx.logProvider.info(
            getLocalizedString(
              "core.provision.configurationFinishNotice",
              PluginDisplayName.Solution
            )
          );
        }

        // 6.
        if (ctx.envInfo.envName === "local") {
          // 6.1 config local env
          const localConfigResult = await configLocalEnvironment(ctx, inputs, ctx.envInfo);
          if (localConfigResult.isErr()) {
            return err(localConfigResult.error);
          }
        } else {
          // 6.2 show message for remote azure provision
          const url = getResourceGroupInPortal(
            ctx.envInfo.state.solution.subscriptionId,
            ctx.envInfo.state.solution.tenantId,
            ctx.envInfo.state.solution.resourceGroupName
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

        // 7. update teams app
        {
          const res = await appManifest.configure(ctx, inputs);
          if (res.isErr()) return err(res.error);
        }

        // 8. show and set state
        if (ctx.envInfo.envName !== "local") {
          const msg = getLocalizedString(
            "core.provision.successNotice",
            ctx.projectSetting.appName
          );
          ctx.userInteraction.showMessage("info", msg, false);
          ctx.logProvider.info(msg);
        }
        ctx.envInfo.state.solution.provisionSucceeded = true;
        return ok([]);
      },
    };
    return ok(action);
  }

  build(context: ContextV3, inputs: InputsWithProjectPath): Result<Action | undefined, FxError> {
    const projectSettings = context.projectSetting as ProjectSettingsV3;
    const actions: Action[] = projectSettings.components
      .filter((resource) => resource.build)
      .map((resource) => {
        const component = resource.code || resource.name;
        return {
          name: `call:${component}.build`,
          type: "call",
          targetAction: `${component}.build`,
          required: true,
        };
      });
    const group: Action = {
      type: "group",
      name: "fx.build",
      mode: "parallel",
      actions: actions,
    };
    return ok(group);
  }

  deploy(context: ContextV3, inputs: InputsWithProjectPath): Result<Action | undefined, FxError> {
    const action: FunctionAction = {
      name: "fx.deploy",
      type: "function",
      question: async (context: ContextV3, inputs: InputsWithProjectPath) => {
        return await getQuestionsForDeployV3(context, context.envInfo!, inputs);
      },
      execute: async (context, inputs) => {
        const ctx = context as ResourceContextV3;
        const projectSettings = ctx.projectSetting as ProjectSettingsV3;
        const inputPlugins = inputs[AzureSolutionQuestionNames.PluginSelectionDeploy] || [];
        const inputComponentNames = inputPlugins.map(pluginName2ComponentName) as string[];
        const thunks = [];
        let hasAzureResource = false;
        // 1. collect resources to deploy
        const isVS = isVSProject(projectSettings);
        for (const component of projectSettings.components) {
          if (
            component.deploy &&
            component.hosting !== undefined &&
            (isVS || inputComponentNames.includes(component.name))
          ) {
            const componentInstance = Container.get<CloudResource>(component.hosting);
            // const actionName = `${component.hosting}.deploy`;
            // const action = await getAction(actionName, ctx, inputs, true);
            thunks.push({
              pluginName: `${component.name}`,
              taskName: "deploy",
              thunk: () => {
                const clonedInputs = cloneDeep(inputs);
                clonedInputs.folder = component.folder;
                clonedInputs.artifactFolder = component.artifactFolder;
                clonedInputs.componentId = component.name;
                return componentInstance.deploy!(ctx, clonedInputs);
              },
            });
            if (AzureResources.includes(component.hosting)) {
              hasAzureResource = true;
            }
          }
        }
        if (inputComponentNames.includes(ComponentNames.AppManifest)) {
          const appManifest = Container.get<AppManifest>(ComponentNames.AppManifest);
          thunks.push({
            pluginName: ComponentNames.AppManifest,
            taskName: "deploy",
            thunk: () => {
              return appManifest.configure(ctx, inputs);
            },
          });
        }
        if (thunks.length === 0) {
          return err(
            new UserError(
              "fx",
              "NoResourcePluginSelected",
              getDefaultString("core.NoPluginSelected"),
              getLocalizedString("core.NoPluginSelected")
            )
          );
        }

        ctx.logProvider.info(
          getLocalizedString(
            "core.deploy.selectedPluginsToDeployNotice",
            PluginDisplayName.Solution,
            JSON.stringify(thunks.map((p) => p.pluginName))
          )
        );

        // 2. check azure account
        if (hasAzureResource) {
          const subscriptionResult = await checkDeployAzureSubscription(
            ctx,
            ctx.envInfo,
            ctx.tokenProvider.azureAccountProvider
          );
          if (subscriptionResult.isErr()) {
            return err(subscriptionResult.error);
          }
          const consent = await askForDeployConsent(
            ctx,
            ctx.tokenProvider.azureAccountProvider,
            ctx.envInfo
          );
          if (consent.isErr()) {
            return err(consent.error);
          }
        }

        // 3. build
        {
          const res = await runActionByName("fx.build", context, inputs);
          if (res.isErr()) return err(res.error);
        }

        // 4. start deploy
        ctx.logProvider.info(
          getLocalizedString("core.deploy.startNotice", PluginDisplayName.Solution)
        );
        const result = await executeConcurrently(thunks, ctx.logProvider);

        if (result.kind === "success") {
          if (hasAzureResource) {
            const botTroubleShootMsg = getBotTroubleShootMessage(hasBot(context.projectSetting));
            const msg =
              getLocalizedString("core.deploy.successNotice", context.projectSetting.appName) +
              botTroubleShootMsg.textForLogging;
            context.logProvider.info(msg);
            if (botTroubleShootMsg.textForLogging) {
              // Show a `Learn more` action button for bot trouble shooting.
              context.userInteraction
                .showMessage(
                  "info",
                  `${getLocalizedString(
                    "core.deploy.successNotice",
                    context.projectSetting.appName
                  )} ${botTroubleShootMsg.textForMsgBox}`,
                  false,
                  botTroubleShootMsg.textForActionButton
                )
                .then((result) => {
                  const userSelected = result.isOk() ? result.value : undefined;
                  if (userSelected === botTroubleShootMsg.textForActionButton) {
                    context.userInteraction.openUrl(botTroubleShootMsg.troubleShootLink);
                  }
                });
            } else {
              context.userInteraction.showMessage("info", msg, false);
            }
          }
          return ok([]);
        } else {
          const msg = getLocalizedString("core.deploy.failNotice", ctx.projectSetting.appName);
          ctx.logProvider.info(msg);
          return err(result.error);
        }
      },
    };
    return ok(action);
  }
}

async function preProvision(
  context: ContextV3,
  inputs: InputsWithProjectPath
): Promise<Result<undefined, FxError>> {
  const ctx = context as ResourceContextV3;
  const envInfo = ctx.envInfo;
  // 1. check M365 tenant
  envInfo.state[ComponentNames.AppManifest] = envInfo.state[ComponentNames.AppManifest] || {};
  envInfo.state.solution = envInfo.state.solution || {};
  const appManifest = envInfo.state[ComponentNames.AppManifest];
  const solutionConfig = envInfo.state.solution;
  solutionConfig.provisionSucceeded = false;
  const tenantIdInConfig = appManifest.tenantId;

  const isLocalDebug = envInfo.envName === "local";
  const tenantIdInTokenRes = await getM365TenantId(ctx.tokenProvider.m365TokenProvider);
  if (tenantIdInTokenRes.isErr()) {
    return err(tenantIdInTokenRes.error);
  }
  const tenantIdInToken = tenantIdInTokenRes.value;

  if (!isLocalDebug) {
    if (tenantIdInConfig && tenantIdInToken && tenantIdInToken !== tenantIdInConfig) {
      return err(
        new UserError(
          "Solution",
          "TeamsAppTenantIdNotRight",
          getLocalizedString("error.M365AccountNotMatch", envInfo.envName)
        )
      );
    }
    if (!tenantIdInConfig) {
      appManifest.tenantId = tenantIdInToken;
      solutionConfig.teamsAppTenantId = tenantIdInToken;
      globalVars.m365TenantId = tenantIdInToken;
    }
  } else {
    const res = await checkWhetherLocalDebugM365TenantMatches(
      envInfo,
      tenantIdInConfig,
      ctx.tokenProvider.m365TokenProvider,
      inputs.projectPath
    );
    if (res.isErr()) {
      return err(res.error);
    }
    envInfo.state[ComponentNames.AppManifest] = envInfo.state[ComponentNames.AppManifest] || {};
    envInfo.state[ComponentNames.AppManifest].tenantId = tenantIdInToken;
    envInfo.state.solution.teamsAppTenantId = tenantIdInToken;
    globalVars.m365TenantId = tenantIdInToken;
  }

  // 3. check Azure configs
  if (hasAzureResourceV3(ctx.projectSetting) && envInfo.envName !== "local") {
    // ask common question and fill in solution config
    const solutionConfigRes = await fillInAzureConfigs(ctx, inputs, envInfo, ctx.tokenProvider);
    if (solutionConfigRes.isErr()) {
      return err(solutionConfigRes.error);
    }

    if (!solutionConfigRes.value.hasSwitchedSubscription) {
      // ask for provision consent
      const consentResult = await askForProvisionConsent(
        ctx,
        ctx.tokenProvider.azureAccountProvider,
        envInfo
      );
      if (consentResult.isErr()) {
        return err(consentResult.error);
      }
    }

    // create resource group if needed
    if (solutionConfig.needCreateResourceGroup) {
      const createRgRes = await resourceGroupHelper.createNewResourceGroup(
        solutionConfig.resourceGroupName,
        ctx.tokenProvider.azureAccountProvider,
        solutionConfig.subscriptionId,
        solutionConfig.location
      );
      if (createRgRes.isErr()) {
        return err(createRgRes.error);
      }
    }

    if (solutionConfigRes.value.hasSwitchedSubscription) {
      updateResourceBaseName(inputs.projectPath, ctx.projectSetting.appName, envInfo.envName);
    }
  }
  return ok(undefined);
}
