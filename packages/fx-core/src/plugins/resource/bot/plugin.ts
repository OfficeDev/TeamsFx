// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { PluginContext } from "@microsoft/teamsfx-api";

import { AADRegistration } from "./aadRegistration";
import * as factory from "./clientFactory";
import * as utils from "./utils/common";
import { LanguageStrategy } from "./languageStrategy";
import { Messages } from "./resources/messages";
import { FxResult, FxBotPluginResultFactory as ResultFactory } from "./result";
import {
  ProgressBarConstants,
  DeployConfigs,
  FolderNames,
  WebAppConstants,
  TemplateProjectsConstants,
  MaxLengths,
  PathInfo,
  BotBicep,
  Alias,
  ConfigKeys,
} from "./constants";
import { getZipDeployEndpoint } from "./utils/zipDeploy";

import * as appService from "@azure/arm-appservice";
import * as fs from "fs-extra";
import {
  CommonStrings,
  PluginBot,
  ConfigNames,
  PluginLocalDebug,
  HostTypes,
} from "./resources/strings";
import {
  CheckAndThrowIfMissing,
  CheckThrowSomethingMissing,
  PackDirExistenceError,
  PreconditionError,
  SomethingMissingError,
} from "./errors";
import { TeamsBotConfig } from "./configs/teamsBotConfig";
import { ProgressBarFactory } from "./progressBars";
import { ResourceNameFactory } from "./utils/resourceNameFactory";
import { AppStudio } from "./appStudio/appStudio";
import { IBotRegistration } from "./appStudio/interfaces/IBotRegistration";
import { Logger } from "./logger";
import { DeployMgr } from "./deployMgr";
import { BotAuthCredential } from "./botAuthCredential";
import { AzureOperations } from "./azureOps";
import { TokenCredentialsBase } from "@azure/ms-rest-nodeauth";
import path from "path";
import { getTemplatesFolder } from "../../../folder";
import { ArmTemplateResult } from "../../../common/armInterface";
import { Bicep, ConstantString, ResourcePlugins } from "../../../common/constants";
import {
  getResourceGroupNameFromResourceId,
  getSiteNameFromResourceId,
  getSubscriptionIdFromResourceId,
} from "../../../common";
import { getActivatedV2ResourcePlugins } from "../../solution/fx-solution/ResourcePluginContainer";
import { NamedArmResourcePluginAdaptor } from "../../solution/fx-solution/v2/adaptor";
import {
  generateBicepFromFile,
  isBotNotificationEnabled,
  isConfigUnifyEnabled,
} from "../../../common/tools";
import { PluginImpl } from "./interface";
import { BOT_ID } from "../appstudio/constants";

export class TeamsBotImpl implements PluginImpl {
  // Made config public, because expect the upper layer to fill inputs.
  public config: TeamsBotConfig = new TeamsBotConfig();
  protected ctx?: PluginContext;

  private async getAzureAccountCredential(): Promise<TokenCredentialsBase> {
    const serviceClientCredentials =
      await this.ctx?.azureAccountProvider?.getAccountCredentialAsync();
    if (!serviceClientCredentials) {
      throw new PreconditionError(Messages.FailToGetAzureCreds, [Messages.TryLoginAzure]);
    }
    return serviceClientCredentials;
  }

  public async scaffold(context: PluginContext): Promise<FxResult> {
    this.ctx = context;
    await this.config.restoreConfigFromContext(context);
    Logger.info(Messages.ScaffoldingBot);

    const handler = await ProgressBarFactory.newProgressBar(
      ProgressBarConstants.SCAFFOLD_TITLE,
      ProgressBarConstants.SCAFFOLD_STEPS_NUM,
      this.ctx
    );
    await handler?.start(ProgressBarConstants.SCAFFOLD_STEP_START);

    if (isBotNotificationEnabled()) {
      // TODO: set host type from input
      // Since the UI design is not finalized yet,
      // for testing purpose we currently use an environment variable to select hostType.
      // Change the logic after question model is implemented.
      if (process.env.TEAMSFX_BOT_HOST_TYPE === "function") {
        this.config.scaffold.hostType = HostTypes.AZURE_FUNCTIONS;
      } else {
        this.config.scaffold.hostType = HostTypes.APP_SERVICE;
      }
    }

    // 1. Copy the corresponding template project into target directory.
    // Get group name.
    const group_name = TemplateProjectsConstants.GROUP_NAME_BOT;
    if (!this.config.actRoles || this.config.actRoles.length === 0) {
      throw new SomethingMissingError("act roles");
    }

    // const hasBot = this.config.actRoles.includes(PluginActRoles.Bot);
    // const hasMsgExt = this.config.actRoles.includes(PluginActRoles.MessageExtension);
    // if (hasBot && hasMsgExt) {
    // group_name = TemplateProjectsConstants.GROUP_NAME_BOT_MSGEXT;
    // } else if (hasBot) {
    //   group_name = TemplateProjectsConstants.GROUP_NAME_BOT;
    // } else {
    //   group_name = TemplateProjectsConstants.GROUP_NAME_MSGEXT;
    // }

    await handler?.next(ProgressBarConstants.SCAFFOLD_STEP_FETCH_ZIP);
    await LanguageStrategy.getTemplateProject(group_name, this.config);

    this.config.saveConfigIntoContext(context);
    Logger.info(Messages.SuccessfullyScaffoldedBot);

    return ResultFactory.Success();
  }

  public async preProvision(context: PluginContext): Promise<FxResult> {
    this.ctx = context;
    await this.config.restoreConfigFromContext(context);
    Logger.info(Messages.PreProvisioningBot);

    // Preconditions checking.
    CheckThrowSomethingMissing(
      ConfigNames.PROGRAMMING_LANGUAGE,
      this.config.scaffold.programmingLanguage
    );

    this.config.saveConfigIntoContext(context);

    return ResultFactory.Success();
  }

  public async provision(context: PluginContext): Promise<FxResult> {
    this.ctx = context;
    await this.config.restoreConfigFromContext(context);
    Logger.info(Messages.ProvisioningBot);

    // Create and register progress bar for cleanup.
    const handler = await ProgressBarFactory.newProgressBar(
      ProgressBarConstants.PROVISION_TITLE,
      ProgressBarConstants.PROVISION_STEPS_NUM,
      this.ctx
    );
    await handler?.start(ProgressBarConstants.PROVISION_STEP_START);

    // 1. Do bot registration.
    await handler?.next(ProgressBarConstants.PROVISION_STEP_BOT_REG);
    await this.registerBotApp();
    return ResultFactory.Success();
  }

  public async updateArmTemplates(ctx: PluginContext): Promise<FxResult> {
    Logger.info(Messages.UpdatingArmTemplatesBot);
    const plugins = getActivatedV2ResourcePlugins(ctx.projectSettings!).map(
      (p) => new NamedArmResourcePluginAdaptor(p)
    );
    const pluginCtx = { plugins: plugins.map((obj) => obj.name) };
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

    Logger.info(Messages.SuccessfullyUpdateArmTemplatesBot);
    return ResultFactory.Success(result);
  }

  public async generateArmTemplates(ctx: PluginContext): Promise<FxResult> {
    Logger.info(Messages.GeneratingArmTemplatesBot);
    const plugins = getActivatedV2ResourcePlugins(ctx.projectSettings!).map(
      (p) => new NamedArmResourcePluginAdaptor(p)
    );
    const pluginCtx = { plugins: plugins.map((obj) => obj.name) };
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

    Logger.info(Messages.SuccessfullyGenerateArmTemplatesBot);
    return ResultFactory.Success(result);
  }

  private async provisionWebApp() {
    CheckThrowSomethingMissing(CommonStrings.SHORT_APP_NAME, this.ctx?.projectSettings?.appName);

    const serviceClientCredentials = await this.getAzureAccountCredential();

    // Suppose we get creds and subs from context.
    const webSiteMgmtClient = factory.createWebSiteMgmtClient(
      serviceClientCredentials,
      this.config.provision.subscriptionId!
    );

    // 1. Provsion app service plan.
    const appServicePlanName =
      this.config.provision.appServicePlan ??
      ResourceNameFactory.createCommonName(
        this.config.resourceNameSuffix,
        this.ctx?.projectSettings?.appName,
        MaxLengths.APP_SERVICE_PLAN_NAME
      );
    Logger.info(Messages.ProvisioningAzureAppServicePlan);
    await AzureOperations.CreateOrUpdateAppServicePlan(
      webSiteMgmtClient,
      this.config.provision.resourceGroup!,
      appServicePlanName,
      utils.generateAppServicePlanConfig(
        this.config.provision.location!,
        this.config.provision.skuName!
      )
    );
    Logger.info(Messages.SuccessfullyProvisionedAzureAppServicePlan);

    // 2. Provision web app.
    const siteEnvelope: appService.WebSiteManagementModels.Site = LanguageStrategy.getSiteEnvelope(
      this.config.scaffold.programmingLanguage!,
      appServicePlanName,
      this.config.provision.location!
    );

    Logger.info(Messages.ProvisioningAzureWebApp);
    const webappResponse = await AzureOperations.CreateOrUpdateAzureWebApp(
      webSiteMgmtClient,
      this.config.provision.resourceGroup!,
      this.config.provision.siteName!,
      siteEnvelope
    );
    Logger.info(Messages.SuccessfullyProvisionedAzureWebApp);

    if (!this.config.provision.siteEndpoint) {
      this.config.provision.siteEndpoint = `${CommonStrings.HTTPS_PREFIX}${webappResponse.defaultHostName}`;
    }

    if (!this.config.provision.appServicePlan) {
      this.config.provision.appServicePlan = appServicePlanName;
    }

    // Update config for manifest.json
    this.ctx!.config.set(
      PluginBot.VALID_DOMAIN,
      `${this.config.provision.siteName}.${WebAppConstants.WEB_APP_SITE_DOMAIN}`
    );
  }

  public async postProvision(context: PluginContext): Promise<FxResult> {
    return ResultFactory.Success();
  }

  public async preDeploy(context: PluginContext): Promise<FxResult> {
    this.ctx = context;
    await this.config.restoreConfigFromContext(context);
    Logger.info(Messages.PreDeployingBot);

    // Preconditions checking.
    const packDirExisted = await fs.pathExists(this.config.scaffold.workingDir!);
    if (!packDirExisted) {
      throw new PackDirExistenceError();
    }

    CheckThrowSomethingMissing(ConfigNames.SITE_ENDPOINT, this.config.provision.siteEndpoint);
    CheckThrowSomethingMissing(
      ConfigNames.PROGRAMMING_LANGUAGE,
      this.config.scaffold.programmingLanguage
    );
    CheckThrowSomethingMissing(
      ConfigNames.BOT_SERVICE_RESOURCE_ID,
      this.config.provision.botWebAppResourceId
    );
    CheckThrowSomethingMissing(ConfigNames.SUBSCRIPTION_ID, this.config.provision.subscriptionId);
    CheckThrowSomethingMissing(ConfigNames.RESOURCE_GROUP, this.config.provision.resourceGroup);

    this.config.saveConfigIntoContext(context);

    return ResultFactory.Success();
  }

  public async deploy(context: PluginContext): Promise<FxResult> {
    this.ctx = context;
    await this.config.restoreConfigFromContext(context);

    this.config.provision.subscriptionId = getSubscriptionIdFromResourceId(
      this.config.provision.botWebAppResourceId!
    );
    this.config.provision.resourceGroup = getResourceGroupNameFromResourceId(
      this.config.provision.botWebAppResourceId!
    );
    this.config.provision.siteName = getSiteNameFromResourceId(
      this.config.provision.botWebAppResourceId!
    );

    Logger.info(Messages.DeployingBot);

    const workingDir = this.config.scaffold.workingDir;
    if (!workingDir) {
      throw new PreconditionError(Messages.WorkingDirIsMissing, []);
    }

    const deployTimeCandidate = Date.now();
    const deployMgr = new DeployMgr(workingDir, this.ctx.envInfo.envName);
    await deployMgr.init();

    if (!(await deployMgr.needsToRedeploy())) {
      Logger.debug(Messages.SkipDeployNoUpdates);
      return ResultFactory.Success();
    }

    const handler = await ProgressBarFactory.newProgressBar(
      ProgressBarConstants.DEPLOY_TITLE,
      ProgressBarConstants.DEPLOY_STEPS_NUM,
      this.ctx
    );

    await handler?.start(ProgressBarConstants.DEPLOY_STEP_START);

    await handler?.next(ProgressBarConstants.DEPLOY_STEP_NPM_INSTALL);
    await LanguageStrategy.localBuild(
      this.config.scaffold.programmingLanguage!,
      workingDir,
      this.config.deploy.unPackFlag === "true" ? true : false
    );

    await handler?.next(ProgressBarConstants.DEPLOY_STEP_ZIP_FOLDER);
    const zipBuffer = utils.zipAFolder(workingDir, DeployConfigs.UN_PACK_DIRS, [
      `${FolderNames.NODE_MODULES}/${FolderNames.KEYTAR}`,
    ]);

    // 2.2 Retrieve publishing credentials.
    const webSiteMgmtClient = new appService.WebSiteManagementClient(
      await this.getAzureAccountCredential(),
      this.config.provision.subscriptionId!
    );
    const listResponse = await AzureOperations.ListPublishingCredentials(
      webSiteMgmtClient,
      this.config.provision.resourceGroup!,
      this.config.provision.siteName!
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

    const zipDeployEndpoint: string = getZipDeployEndpoint(this.config.provision.siteName!);
    await handler?.next(ProgressBarConstants.DEPLOY_STEP_ZIP_DEPLOY);
    await AzureOperations.ZipDeployPackage(zipDeployEndpoint, zipBuffer, config);

    await deployMgr.updateLastDeployTime(deployTimeCandidate);

    this.config.saveConfigIntoContext(context);
    Logger.info(Messages.SuccessfullyDeployedBot);

    return ResultFactory.Success();
  }

  public async localDebug(context: PluginContext): Promise<FxResult> {
    this.ctx = context;
    await this.config.restoreConfigFromContext(context);

    const handler = await ProgressBarFactory.newProgressBar(
      ProgressBarConstants.LOCAL_DEBUG_TITLE,
      ProgressBarConstants.LOCAL_DEBUG_STEPS_NUM,
      this.ctx
    );

    await handler?.start(ProgressBarConstants.LOCAL_DEBUG_STEP_START);

    await handler?.next(ProgressBarConstants.LOCAL_DEBUG_STEP_BOT_REG);
    await this.createNewBotRegistrationOnAppStudio();

    this.config.saveConfigIntoContext(context);

    return ResultFactory.Success();
  }

  public async postLocalDebug(context: PluginContext): Promise<FxResult> {
    this.ctx = context;
    await this.config.restoreConfigFromContext(context);

    if (isConfigUnifyEnabled()) {
      CheckThrowSomethingMissing(
        ConfigNames.LOCAL_ENDPOINT,
        context.envInfo.state.get(ResourcePlugins.Bot).get(ConfigKeys.SITE_ENDPOINT)
      );
      await this.updateMessageEndpointOnAppStudio(
        `${context.envInfo.state.get(ResourcePlugins.Bot).get(ConfigKeys.SITE_ENDPOINT)}${
          CommonStrings.MESSAGE_ENDPOINT_SUFFIX
        }`
      );
    } else {
      CheckThrowSomethingMissing(ConfigNames.LOCAL_ENDPOINT, this.config.localDebug.localEndpoint);
      await this.updateMessageEndpointOnAppStudio(
        `${this.config.localDebug.localEndpoint}${CommonStrings.MESSAGE_ENDPOINT_SUFFIX}`
      );
    }

    this.config.saveConfigIntoContext(context);

    return ResultFactory.Success();
  }

  private async updateMessageEndpointOnAppStudio(endpoint: string) {
    const appStudioToken = await this.ctx?.appStudioToken?.getAccessToken();
    CheckThrowSomethingMissing(ConfigNames.APPSTUDIO_TOKEN, appStudioToken);
    CheckThrowSomethingMissing(
      ConfigNames.LOCAL_BOT_ID,
      isConfigUnifyEnabled()
        ? this.ctx?.envInfo.state.get(ResourcePlugins.Bot).get(BOT_ID)
        : this.config.localDebug.localBotId
    );

    const botReg: IBotRegistration = {
      botId: isConfigUnifyEnabled()
        ? this.ctx?.envInfo.state.get(ResourcePlugins.Bot).get(BOT_ID)
        : this.config.localDebug.localBotId,
      name: this.ctx!.projectSettings?.appName + PluginLocalDebug.LOCAL_DEBUG_SUFFIX,
      description: "",
      iconUrl: "",
      messagingEndpoint: endpoint,
      callingEndpoint: "",
    };

    await AppStudio.updateMessageEndpoint(appStudioToken!, botReg.botId!, botReg);
  }

  private async updateMessageEndpointOnAzure(endpoint: string) {
    const serviceClientCredentials = await this.getAzureAccountCredential();

    const botClient = factory.createAzureBotServiceClient(
      serviceClientCredentials,
      this.config.provision.subscriptionId!
    );

    if (!this.config.provision.botChannelRegName) {
      throw new SomethingMissingError(CommonStrings.BOT_CHANNEL_REGISTRATION);
    }
    const botChannelRegistrationName = this.config.provision.botChannelRegName;
    Logger.info(Messages.UpdatingBotMessageEndpoint);
    await AzureOperations.UpdateBotChannelRegistration(
      botClient,
      this.config.provision.resourceGroup!,
      botChannelRegistrationName,
      this.config.scaffold.botId!,
      endpoint,
      this.ctx?.projectSettings?.appName
    );
    Logger.info(Messages.SuccessfullyUpdatedBotMessageEndpoint);
  }

  private async createNewBotRegistrationOnAppStudio() {
    const token = CheckAndThrowIfMissing(
      ConfigNames.GRAPH_TOKEN,
      await this.ctx?.graphTokenProvider?.getAccessToken()
    );
    const name = CheckAndThrowIfMissing(
      CommonStrings.SHORT_APP_NAME,
      this.ctx?.projectSettings?.appName
    );

    // 1. Create a new AAD App Registration with client secret.
    const aadDisplayName = ResourceNameFactory.createCommonName(
      this.config.resourceNameSuffix,
      name,
      MaxLengths.AAD_DISPLAY_NAME
    );

    let botAuthCreds: BotAuthCredential = new BotAuthCredential();
    if (
      this.config.localDebug.botAADCreated()
      // if user input AAD, the object id is not required
      // && (await AppStudio.isAADAppExisting(appStudioToken!, this.config.localDebug.localObjectId!))
    ) {
      botAuthCreds.clientId = this.config.localDebug.localBotId;
      botAuthCreds.clientSecret = this.config.localDebug.localBotPassword;
      botAuthCreds.objectId = this.config.localDebug.localObjectId;
      Logger.debug(Messages.SuccessfullyGetExistingBotAadAppCredential);
    } else if (
      isConfigUnifyEnabled() &&
      this.ctx?.envInfo.state.get(ResourcePlugins.Bot)?.get(BOT_ID)
    ) {
      botAuthCreds.clientId = this.ctx?.envInfo.state
        .get(ResourcePlugins.Bot)
        .get(PluginBot.BOT_ID);
      botAuthCreds.clientSecret = this.ctx?.envInfo.state
        .get(ResourcePlugins.Bot)
        .get(PluginBot.BOT_PASSWORD);
      botAuthCreds.clientSecret = this.ctx?.envInfo.state
        .get(ResourcePlugins.Bot)
        .get(PluginBot.OBJECT_ID);
      Logger.debug(Messages.SuccessfullyGetExistingBotAadAppCredential);
    } else {
      Logger.info(Messages.ProvisioningBotRegistration);
      botAuthCreds = await AADRegistration.registerAADAppAndGetSecretByGraph(
        token,
        aadDisplayName,
        this.config.localDebug.localObjectId,
        this.config.localDebug.localBotId
      );
      Logger.info(Messages.SuccessfullyProvisionedBotRegistration);
    }

    // 2. Register bot by app studio.
    const botReg: IBotRegistration = {
      botId: botAuthCreds.clientId,
      name: this.ctx!.projectSettings?.appName + PluginLocalDebug.LOCAL_DEBUG_SUFFIX,
      description: "",
      iconUrl: "",
      messagingEndpoint: "",
      callingEndpoint: "",
    };

    Logger.info(Messages.ProvisioningBotRegistration);
    const appStudioToken = await this.ctx?.appStudioToken?.getAccessToken();
    CheckThrowSomethingMissing(ConfigNames.APPSTUDIO_TOKEN, appStudioToken);
    await AppStudio.createBotRegistration(appStudioToken!, botReg);
    Logger.info(Messages.SuccessfullyProvisionedBotRegistration);

    if (isConfigUnifyEnabled()) {
      if (!this.config.scaffold.botId) {
        this.config.scaffold.botId = botAuthCreds.clientId;
      }
      if (!this.config.scaffold.botPassword) {
        this.config.scaffold.botPassword = botAuthCreds.clientSecret;
      }
      if (!this.config.scaffold.objectId) {
        this.config.scaffold.objectId = botAuthCreds.objectId;
      }
    } else {
      if (!this.config.localDebug.localBotId) {
        this.config.localDebug.localBotId = botAuthCreds.clientId;
      }

      if (!this.config.localDebug.localBotPassword) {
        this.config.localDebug.localBotPassword = botAuthCreds.clientSecret;
      }

      if (!this.config.localDebug.localObjectId) {
        this.config.localDebug.localObjectId = botAuthCreds.objectId;
      }
    }
  }

  private async registerBotApp() {
    const token = CheckAndThrowIfMissing(
      ConfigNames.GRAPH_TOKEN,
      await this.ctx?.graphTokenProvider?.getAccessToken()
    );
    const name = CheckAndThrowIfMissing(
      CommonStrings.SHORT_APP_NAME,
      this.ctx?.projectSettings?.appName
    );
    const ctx = CheckAndThrowIfMissing(ConfigNames.GRAPH_TOKEN, this.ctx);

    let botAuthCredential = new BotAuthCredential();

    if (!this.config.scaffold.botAADCreated()) {
      const aadDisplayName = ResourceNameFactory.createCommonName(
        this.config.resourceNameSuffix,
        name,
        MaxLengths.AAD_DISPLAY_NAME
      );
      botAuthCredential = await AADRegistration.registerAADAppAndGetSecretByGraph(
        token,
        aadDisplayName,
        this.config.scaffold.objectId,
        this.config.scaffold.botId
      );

      this.config.scaffold.botId = botAuthCredential.clientId;
      this.config.scaffold.botPassword = botAuthCredential.clientSecret;
      this.config.scaffold.objectId = botAuthCredential.objectId;

      this.config.saveConfigIntoContext(ctx); // Checkpoint for aad app provision.
      Logger.info(Messages.SuccessfullyCreatedBotAadApp);
    }
  }

  private async provisionBotServiceOnAzure(botAuthCreds: BotAuthCredential) {
    const serviceClientCredentials = await this.getAzureAccountCredential();

    // Provision a bot channel registration resource on azure.
    const botClient = factory.createAzureBotServiceClient(
      serviceClientCredentials,
      this.config.provision.subscriptionId!
    );

    const botChannelRegistrationName = this.config.provision.botChannelRegName
      ? this.config.provision.botChannelRegName
      : ResourceNameFactory.createCommonName(
          this.config.resourceNameSuffix,
          this.ctx?.projectSettings?.appName,
          MaxLengths.BOT_CHANNEL_REG_NAME
        );

    Logger.info(Messages.ProvisioningAzureBotChannelRegistration);
    await AzureOperations.CreateBotChannelRegistration(
      botClient,
      this.config.provision.resourceGroup!,
      botChannelRegistrationName,
      botAuthCreds.clientId!,
      this.ctx?.projectSettings?.appName
    );
    Logger.info(Messages.SuccessfullyProvisionedAzureBotChannelRegistration);

    // Add Teams Client as a channel to the resource above.
    Logger.info(Messages.ProvisioningMsTeamsChannel);
    await AzureOperations.LinkTeamsChannel(
      botClient,
      this.config.provision.resourceGroup!,
      botChannelRegistrationName
    );
    Logger.info(Messages.SuccessfullyProvisionedMsTeamsChannel);

    if (!this.config.provision.botChannelRegName) {
      this.config.provision.botChannelRegName = botChannelRegistrationName;
    }
  }
}
