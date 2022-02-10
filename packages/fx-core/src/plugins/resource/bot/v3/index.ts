// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  AzureAccountProvider,
  AzureSolutionSettings,
  EnvConfig,
  err,
  FxError,
  ok,
  Result,
  TokenProvider,
  v2,
  v3,
  Void,
} from "@microsoft/teamsfx-api";
import * as path from "path";
import { Service } from "typedi";
import { ArmTemplateResult } from "../../../../common/armInterface";
import { Bicep, ConstantString } from "../../../../common/constants";
import {
  generateBicepFromFile,
  getResourceGroupNameFromResourceId,
  getSiteNameFromResourceId,
  getSubscriptionIdFromResourceId,
} from "../../../../common/tools";
import { CommonErrorHandlerMW } from "../../../../core/middleware/CommonErrorHandlerMW";
import { getTemplatesFolder } from "../../../../folder";
import {
  AzureSolutionQuestionNames,
  BotOptionItem,
  MessageExtensionItem,
  TabOptionItem,
} from "../../../solution/fx-solution/question";
import { BuiltInFeaturePluginNames } from "../../../solution/fx-solution/v3/constants";
import {
  AzureConstants,
  BotBicep,
  DeployConfigs,
  FolderNames,
  MaxLengths,
  PathInfo,
  ProgressBarConstants,
  TemplateProjectsConstants,
} from "../constants";
import { LanguageStrategy } from "../languageStrategy";
import { ProgressBarFactory } from "../progressBars";
import { Messages } from "../resources/messages";
import fs from "fs-extra";
import { CommonStrings, ConfigNames, PluginLocalDebug } from "../resources/strings";
import { TokenCredentialsBase } from "@azure/ms-rest-nodeauth";
import * as factory from "../clientFactory";
import { ResourceNameFactory } from "../utils/resourceNameFactory";
import { AADRegistration } from "../aadRegistration";
import { IBotRegistration } from "../appStudio/interfaces/IBotRegistration";
import { AppStudio } from "../appStudio/appStudio";
import { DeployMgr } from "../deployMgr";
import * as utils from "../utils/common";
import * as appService from "@azure/arm-appservice";
import { AzureOperations } from "../azureOps";
import { getZipDeployEndpoint } from "../utils/zipDeploy";
import {
  ScaffoldAction,
  ScaffoldActionName,
  ScaffoldContext,
  scaffoldFromTemplates,
} from "../../../../common/template-utils/templatesActions";
import { ProgrammingLanguage } from "../enums/programmingLanguage";
import {
  CheckThrowSomethingMissing,
  PackDirectoryExistenceError,
  PreconditionError,
  TemplateZipFallbackError,
  UnzipError,
} from "./error";
import { ensureSolutionSettings } from "../../../solution/fx-solution/utils/solutionSettingsHelper";

@Service(BuiltInFeaturePluginNames.bot)
export class NodeJSBotPluginV3 implements v3.FeaturePlugin {
  name = BuiltInFeaturePluginNames.bot;
  displayName = "NodeJS Bot";

  getProgrammingLanguage(ctx: v2.Context): ProgrammingLanguage {
    const rawProgrammingLanguage = ctx.projectSetting.programmingLanguage;
    if (
      rawProgrammingLanguage &&
      utils.existsInEnumValues(rawProgrammingLanguage, ProgrammingLanguage)
    ) {
      return rawProgrammingLanguage as ProgrammingLanguage;
    }
    return ProgrammingLanguage.JavaScript;
  }
  getLangKey(ctx: v2.Context): string {
    const rawProgrammingLanguage = ctx.projectSetting.programmingLanguage;
    if (
      rawProgrammingLanguage &&
      utils.existsInEnumValues(rawProgrammingLanguage, ProgrammingLanguage)
    ) {
      const programmingLanguage = rawProgrammingLanguage as ProgrammingLanguage;
      return utils.convertToLangKey(programmingLanguage);
    }
    return "js";
  }

  @hooks([CommonErrorHandlerMW({ telemetry: { component: BuiltInFeaturePluginNames.bot } })])
  async scaffold(
    ctx: v3.ContextWithManifestProvider,
    inputs: v2.InputsWithProjectPath
  ): Promise<Result<Void, FxError>> {
    ctx.logProvider.info(Messages.ScaffoldingBot);

    const handler = await ProgressBarFactory.newProgressBar(
      ProgressBarConstants.SCAFFOLD_TITLE,
      ProgressBarConstants.SCAFFOLD_STEPS_NUM,
      ctx
    );
    await handler?.start(ProgressBarConstants.SCAFFOLD_STEP_START);
    const group_name = TemplateProjectsConstants.GROUP_NAME_BOT_MSGEXT;
    const lang = this.getLangKey(ctx);
    const workingDir = path.join(inputs.projectPath, CommonStrings.BOT_WORKING_DIR_NAME);

    await handler?.next(ProgressBarConstants.SCAFFOLD_STEP_FETCH_ZIP);
    await scaffoldFromTemplates({
      group: group_name,
      lang: lang,
      scenario: TemplateProjectsConstants.DEFAULT_SCENARIO_NAME,
      templatesFolderName: TemplateProjectsConstants.TEMPLATE_FOLDER_NAME,
      dst: workingDir,
      onActionEnd: async (action: ScaffoldAction, context: ScaffoldContext) => {
        if (action.name === ScaffoldActionName.FetchTemplatesUrlWithTag) {
          ctx.logProvider.info(Messages.SuccessfullyRetrievedTemplateZip(context.zipUrl ?? ""));
        }
      },
      onActionError: async (action: ScaffoldAction, context: ScaffoldContext, error: Error) => {
        ctx.logProvider.info(error.toString());
        switch (action.name) {
          case ScaffoldActionName.FetchTemplatesUrlWithTag:
          case ScaffoldActionName.FetchTemplatesZipFromUrl:
            ctx.logProvider.info(Messages.FallingBackToUseLocalTemplateZip);
            break;
          case ScaffoldActionName.FetchTemplateZipFromLocal:
            throw new TemplateZipFallbackError();
          case ScaffoldActionName.Unzip:
            throw new UnzipError(context.dst);
          default:
            throw new Error(error.message);
        }
      },
    });
    ctx.logProvider.info(Messages.SuccessfullyScaffoldedBot);
    handler?.end(true);
    return ok(Void);
  }
  @hooks([CommonErrorHandlerMW({ telemetry: { component: BuiltInFeaturePluginNames.bot } })])
  async generateResourceTemplate(
    ctx: v3.ContextWithManifestProvider,
    inputs: v2.InputsWithProjectPath
  ): Promise<Result<v2.ResourceTemplate[], FxError>> {
    ctx.logProvider.info(Messages.GeneratingArmTemplatesBot);
    const solutionSettings = ctx.projectSetting.solutionSettings as
      | AzureSolutionSettings
      | undefined;
    const pluginCtx = { plugins: solutionSettings ? solutionSettings.activeResourcePlugins : [] };
    const bicepTemplateDir = path.join(getTemplatesFolder(), PathInfo.BicepTemplateRelativeDir);
    const provisionOrchestration = await generateBicepFromFile(
      path.join(bicepTemplateDir, Bicep.ProvisionFileName),
      pluginCtx
    );
    const provisionModules = await generateBicepFromFile(
      path.join(bicepTemplateDir, PathInfo.ProvisionModuleTemplateFileName),
      pluginCtx
    );
    const configOrchestration = await generateBicepFromFile(
      path.join(bicepTemplateDir, Bicep.ConfigFileName),
      pluginCtx
    );
    const configModule = await generateBicepFromFile(
      path.join(bicepTemplateDir, PathInfo.ConfigurationModuleTemplateFileName),
      pluginCtx
    );
    const result: ArmTemplateResult = {
      Provision: {
        Orchestration: provisionOrchestration,
        Modules: { bot: provisionModules },
      },
      Configuration: {
        Orchestration: configOrchestration,
        Modules: { bot: configModule },
      },
      Reference: {
        resourceId: BotBicep.resourceId,
        hostName: BotBicep.hostName,
        webAppEndpoint: BotBicep.webAppEndpoint,
      },
      Parameters: JSON.parse(
        await fs.readFile(
          path.join(bicepTemplateDir, Bicep.ParameterFileName),
          ConstantString.UTF8Encoding
        )
      ),
    };
    ctx.logProvider.info(Messages.SuccessfullyGenerateArmTemplatesBot);
    return ok([{ kind: "bicep", template: result }]);
  }
  @hooks([CommonErrorHandlerMW({ telemetry: { component: BuiltInFeaturePluginNames.bot } })])
  async addFeature(
    ctx: v3.ContextWithManifestProvider,
    inputs: v2.InputsWithProjectPath
  ): Promise<Result<v2.ResourceTemplate[], FxError>> {
    ensureSolutionSettings(ctx.projectSetting);
    const solutionSettings = ctx.projectSetting.solutionSettings as AzureSolutionSettings;
    const capabilities = solutionSettings.capabilities;
    const newCapabilitySet = new Set<string>();
    capabilities.forEach((c) => newCapabilitySet.add(c));
    let templates: v2.ResourceTemplate[] = [];
    if (!(capabilities.includes(TabOptionItem.id) || capabilities.includes(BotOptionItem.id))) {
      // bot is added for first time, scaffold and generate resource template
      const scaffoldRes = await this.scaffold(ctx, inputs);
      if (scaffoldRes.isErr()) return err(scaffoldRes.error);
      const armRes = await this.generateResourceTemplate(ctx, inputs);
      if (armRes.isErr()) return err(armRes.error);
      templates = armRes.value;
    }
    const capabilitiesToAddManifest: v3.ManifestCapability[] = [];
    const capabilitiesAnswer = inputs[AzureSolutionQuestionNames.Capabilities] as string[];
    if (capabilitiesAnswer.includes(BotOptionItem.id)) {
      capabilitiesToAddManifest.push({ name: "Bot" });
      newCapabilitySet.add(BotOptionItem.id);
    }
    if (capabilitiesAnswer.includes(MessageExtensionItem.id)) {
      capabilitiesToAddManifest.push({ name: "MessageExtension" });
      newCapabilitySet.add(MessageExtensionItem.id);
    }
    const update = await ctx.appManifestProvider.addCapabilities(
      ctx,
      inputs,
      capabilitiesToAddManifest
    );
    if (update.isErr()) return err(update.error);

    solutionSettings.capabilities = Array.from(newCapabilitySet);

    const activeResourcePlugins = solutionSettings.activeResourcePlugins;
    if (!activeResourcePlugins.includes(this.name)) activeResourcePlugins.push(this.name);
    return ok(templates);
  }
  @hooks([CommonErrorHandlerMW({ telemetry: { component: BuiltInFeaturePluginNames.bot } })])
  async afterOtherFeaturesAdded(
    ctx: v3.ContextWithManifestProvider,
    inputs: v3.OtherFeaturesAddedInputs
  ): Promise<Result<v2.ResourceTemplate[], FxError>> {
    ctx.logProvider.info(Messages.UpdatingArmTemplatesBot);
    const solutionSettings = ctx.projectSetting.solutionSettings as
      | AzureSolutionSettings
      | undefined;
    const pluginCtx = { plugins: solutionSettings ? solutionSettings.activeResourcePlugins : [] };
    const bicepTemplateDir = path.join(getTemplatesFolder(), PathInfo.BicepTemplateRelativeDir);
    const configModule = await generateBicepFromFile(
      path.join(bicepTemplateDir, PathInfo.ConfigurationModuleTemplateFileName),
      pluginCtx
    );
    const result: ArmTemplateResult = {
      Reference: {
        resourceId: BotBicep.resourceId,
        hostName: BotBicep.hostName,
        webAppEndpoint: BotBicep.webAppEndpoint,
      },
      Configuration: {
        Modules: { bot: configModule },
      },
    };
    ctx.logProvider.info(Messages.SuccessfullyUpdateArmTemplatesBot);
    return ok([{ kind: "bicep", template: result }]);
  }
  private async getAzureAccountCredenial(
    tokenProvider: AzureAccountProvider
  ): Promise<TokenCredentialsBase> {
    const serviceClientCredentials = await tokenProvider.getAccountCredentialAsync();
    if (!serviceClientCredentials) {
      throw new PreconditionError(Messages.FailToGetAzureCreds, [Messages.TryLoginAzure]);
    }
    return serviceClientCredentials;
  }

  private async createOrGetBotAppRegistration(
    ctx: v2.Context,
    envInfo: v3.EnvInfoV3,
    tokenProvider: TokenProvider
  ): Promise<Result<Void, FxError>> {
    const token = await tokenProvider.graphTokenProvider.getAccessToken();
    CheckThrowSomethingMissing(ConfigNames.GRAPH_TOKEN, token);
    CheckThrowSomethingMissing(CommonStrings.SHORT_APP_NAME, ctx.projectSetting.appName);
    let botConfig = envInfo.state[this.name];
    if (!botConfig) botConfig = {};
    botConfig = botConfig as v3.AzureBot;
    const botAADCreated = botConfig?.botId !== undefined && botConfig?.botPassword !== undefined;
    if (!botAADCreated) {
      const solutionConfig = envInfo?.state.solution as v3.AzureSolutionConfig;
      const resourceNameSuffix = solutionConfig.resourceNameSuffix
        ? solutionConfig.resourceNameSuffix
        : utils.genUUID();
      const aadDisplayName = ResourceNameFactory.createCommonName(
        resourceNameSuffix,
        ctx.projectSetting.appName,
        MaxLengths.AAD_DISPLAY_NAME
      );
      const botAuthCreds = await AADRegistration.registerAADAppAndGetSecretByGraph(
        token!,
        aadDisplayName,
        botConfig.objectId,
        botConfig.botId
      );
      botConfig.botId = botAuthCreds.clientId;
      botConfig.botPassword = botAuthCreds.clientSecret;
      botConfig.objectId = botAuthCreds.objectId;
      ctx.logProvider.info(Messages.SuccessfullyCreatedBotAadApp);
    }

    if (envInfo.envName === "local") {
      // 2. Register bot by app studio.
      const botReg: IBotRegistration = {
        botId: botConfig.botId,
        name: ctx.projectSetting.appName + PluginLocalDebug.LOCAL_DEBUG_SUFFIX,
        description: "",
        iconUrl: "",
        messagingEndpoint: "",
        callingEndpoint: "",
      };
      ctx.logProvider.info(Messages.ProvisioningBotRegistration);
      const appStudioToken = await tokenProvider.appStudioToken.getAccessToken();
      CheckThrowSomethingMissing(ConfigNames.APPSTUDIO_TOKEN, appStudioToken);
      await AppStudio.createBotRegistration(appStudioToken!, botReg);
      ctx.logProvider.info(Messages.SuccessfullyProvisionedBotRegistration);
    }
    return ok(Void);
  }

  @hooks([CommonErrorHandlerMW({ telemetry: { component: BuiltInFeaturePluginNames.bot } })])
  async provisionResource(
    ctx: v2.Context,
    inputs: v2.InputsWithProjectPath,
    envInfo: v3.EnvInfoV3,
    tokenProvider: TokenProvider
  ): Promise<Result<Void, FxError>> {
    if (envInfo.envName === "local") {
      const handler = await ProgressBarFactory.newProgressBar(
        ProgressBarConstants.LOCAL_DEBUG_TITLE,
        ProgressBarConstants.LOCAL_DEBUG_STEPS_NUM,
        ctx
      );

      await handler?.start(ProgressBarConstants.LOCAL_DEBUG_STEP_START);

      await handler?.next(ProgressBarConstants.LOCAL_DEBUG_STEP_BOT_REG);
      await this.createOrGetBotAppRegistration(ctx, envInfo, tokenProvider);
    } else {
      ctx.logProvider.info(Messages.ProvisioningBot);
      // Create and register progress bar for cleanup.
      const handler = await ProgressBarFactory.newProgressBar(
        ProgressBarConstants.PROVISION_TITLE,
        ProgressBarConstants.PROVISION_STEPS_NUM,
        ctx
      );
      await handler?.start(ProgressBarConstants.PROVISION_STEP_START);

      // 0. Check Resource Provider
      const azureCredential = await this.getAzureAccountCredenial(
        tokenProvider.azureAccountProvider
      );
      const solutionConfig = envInfo.state.solution as v3.AzureSolutionConfig;
      const rpClient = factory.createResourceProviderClient(
        azureCredential,
        solutionConfig.subscriptionId!
      );
      await factory.ensureResourceProvider(rpClient, AzureConstants.requiredResourceProviders);

      // 1. Do bot registration.
      await handler?.next(ProgressBarConstants.PROVISION_STEP_BOT_REG);
      await this.createOrGetBotAppRegistration(ctx, envInfo, tokenProvider);
    }
    return ok(Void);
  }

  private async updateMessageEndpointOnAppStudio(
    appName: string,
    tokenProvider: TokenProvider,
    botId: string,
    endpoint: string
  ) {
    const appStudioToken = await tokenProvider.appStudioToken.getAccessToken();
    CheckThrowSomethingMissing(ConfigNames.APPSTUDIO_TOKEN, appStudioToken);
    CheckThrowSomethingMissing(ConfigNames.LOCAL_BOT_ID, botId);

    const botReg: IBotRegistration = {
      botId: botId,
      name: appName + PluginLocalDebug.LOCAL_DEBUG_SUFFIX,
      description: "",
      iconUrl: "",
      messagingEndpoint: endpoint,
      callingEndpoint: "",
    };

    await AppStudio.updateMessageEndpoint(appStudioToken!, botReg.botId!, botReg);
  }

  @hooks([CommonErrorHandlerMW({ telemetry: { component: BuiltInFeaturePluginNames.bot } })])
  async configureResource(
    ctx: v2.Context,
    inputs: v2.InputsWithProjectPath,
    envInfo: v3.EnvInfoV3,
    tokenProvider: TokenProvider
  ): Promise<Result<Void, FxError>> {
    if (envInfo.envName === "local") {
      const botConfig = envInfo.state[this.name] as v3.AzureBot;
      CheckThrowSomethingMissing(ConfigNames.LOCAL_ENDPOINT, botConfig.siteEndpoint);
      await this.updateMessageEndpointOnAppStudio(
        ctx.projectSetting.appName,
        tokenProvider,
        botConfig.botId,
        `${botConfig.siteEndpoint}${CommonStrings.MESSAGE_ENDPOINT_SUFFIX}`
      );
    }
    return ok(Void);
  }

  @hooks([CommonErrorHandlerMW({ telemetry: { component: BuiltInFeaturePluginNames.bot } })])
  async deploy(
    ctx: v2.Context,
    inputs: v2.InputsWithProjectPath,
    envInfo: v2.DeepReadonly<v3.EnvInfoV3>,
    tokenProvider: AzureAccountProvider
  ): Promise<Result<Void, FxError>> {
    ctx.logProvider.info(Messages.PreDeployingBot);

    // Preconditions checking.
    const workingDir = path.join(inputs.projectPath, CommonStrings.BOT_WORKING_DIR_NAME);
    if (!workingDir) {
      throw new PreconditionError(Messages.WorkingDirIsMissing, []);
    }
    const packDirExisted = await fs.pathExists(workingDir);
    if (!packDirExisted) {
      throw new PackDirectoryExistenceError();
    }

    const botConfig = envInfo.state[this.name] as v3.AzureBot;
    const programmingLanguage = this.getProgrammingLanguage(ctx);
    CheckThrowSomethingMissing(ConfigNames.SITE_ENDPOINT, botConfig.siteEndpoint);
    CheckThrowSomethingMissing(ConfigNames.PROGRAMMING_LANGUAGE, programmingLanguage);
    CheckThrowSomethingMissing(ConfigNames.BOT_SERVICE_RESOURCE_ID, botConfig.botWebAppResourceId);

    const subscriptionId = getSubscriptionIdFromResourceId(botConfig.botWebAppResourceId);
    const resourceGroup = getResourceGroupNameFromResourceId(botConfig.botWebAppResourceId);
    const siteName = getSiteNameFromResourceId(botConfig.botWebAppResourceId);

    CheckThrowSomethingMissing(ConfigNames.SUBSCRIPTION_ID, subscriptionId);
    CheckThrowSomethingMissing(ConfigNames.RESOURCE_GROUP, resourceGroup);

    ctx.logProvider.info(Messages.DeployingBot);

    const deployTimeCandidate = Date.now();
    const deployMgr = new DeployMgr(workingDir, envInfo.envName);
    await deployMgr.init();

    if (!(await deployMgr.needsToRedeploy())) {
      ctx.logProvider.debug(Messages.SkipDeployNoUpdates);
      return ok(Void);
    }

    const handler = await ProgressBarFactory.newProgressBar(
      ProgressBarConstants.DEPLOY_TITLE,
      ProgressBarConstants.DEPLOY_STEPS_NUM,
      ctx
    );

    await handler?.start(ProgressBarConstants.DEPLOY_STEP_START);

    await handler?.next(ProgressBarConstants.DEPLOY_STEP_NPM_INSTALL);
    const unPackFlag = (envInfo.config as EnvConfig).bot?.unPackFlag as string;
    await LanguageStrategy.localBuild(
      programmingLanguage,
      workingDir,
      unPackFlag === "false" ? false : true
    );

    await handler?.next(ProgressBarConstants.DEPLOY_STEP_ZIP_FOLDER);
    const zipBuffer = utils.zipAFolder(workingDir, DeployConfigs.UN_PACK_DIRS, [
      `${FolderNames.NODE_MODULES}/${FolderNames.KEYTAR}`,
    ]);

    // 2.2 Retrieve publishing credentials.
    const webSiteMgmtClient = new appService.WebSiteManagementClient(
      await this.getAzureAccountCredenial(tokenProvider),
      subscriptionId!
    );
    const listResponse = await AzureOperations.ListPublishingCredentials(
      webSiteMgmtClient,
      resourceGroup!,
      siteName!
    );

    const publishingUserName = listResponse.publishingUserName
      ? listResponse.publishingUserName
      : "";
    const publishingPassword = listResponse.publishingPassword
      ? listResponse.publishingPassword
      : "";
    const encryptedCreds: string = utils.toBase64(`${publishingUserName}:${publishingPassword}`);

    const config = {
      headers: {
        Authorization: `Basic ${encryptedCreds}`,
      },
      maxContentLength: Infinity,
      maxBodyLength: Infinity,
    };

    const zipDeployEndpoint: string = getZipDeployEndpoint(botConfig.siteName);
    await handler?.next(ProgressBarConstants.DEPLOY_STEP_ZIP_DEPLOY);
    await AzureOperations.ZipDeployPackage(zipDeployEndpoint, zipBuffer, config);

    await deployMgr.updateLastDeployTime(deployTimeCandidate);

    ctx.logProvider.info(Messages.SuccessfullyDeployedBot);

    return ok(Void);
  }
}
