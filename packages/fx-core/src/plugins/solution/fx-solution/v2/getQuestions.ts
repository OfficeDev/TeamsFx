import {
  AzureSolutionSettings,
  CLIPlatforms,
  DynamicPlatforms,
  err,
  Func,
  FxError,
  Inputs,
  InvalidInputError,
  MultiSelectQuestion,
  ok,
  OptionItem,
  Platform,
  QTreeNode,
  Result,
  SingleSelectQuestion,
  Stage,
  TokenProvider,
  UserError,
  v2,
} from "@microsoft/teamsfx-api";
import Container from "typedi";
import { HelpLinks, ResourcePlugins } from "../../../../common/constants";
import { Constants as AppStudioConstants } from "../../../resource/appstudio/constants";
import { PluginNames, SolutionError, SolutionSource } from "../constants";
import {
  ApiConnectionOptionItem,
  AskSubscriptionQuestion,
  AzureResourceApim,
  AzureResourceFunction,
  AzureResourceKeyVault,
  AzureResourceSQL,
  AzureResourcesQuestion,
  AzureSolutionQuestionNames,
  BotNewUIOptionItem,
  BotOptionItem,
  BotSsoItem,
  CicdOptionItem,
  CommandAndResponseOptionItem,
  createAddAzureResourceQuestion,
  createAddCloudResourceOptions,
  DeployPluginSelectQuestion,
  getUserEmailQuestion,
  M365SearchAppOptionItem,
  M365SsoLaunchPageOptionItem,
  MessageExtensionItem,
  MessageExtensionNewUIItem,
  NotificationOptionItem,
  SingleSignOnOptionItem,
  TabNewUIOptionItem,
  TabNonSsoItem,
  TabOptionItem,
  TabSPFxItem,
  TabSsoItem,
} from "../question";
import {
  getAllV2ResourcePluginMap,
  getAllV2ResourcePlugins,
  ResourcePluginsV2,
} from "../ResourcePluginContainer";
import { checkWetherProvisionSucceeded, getSelectedPlugins, isAzureProject } from "./utils";
import { isV3 } from "../../../../core/globalVars";
import { TeamsAppSolutionNameV2 } from "./constants";
import { BuiltInFeaturePluginNames } from "../v3/constants";
import { AppStudioPluginV3 } from "../../../resource/appstudio/v3";
import { canAddCapability, canAddResource } from "./executeUserTask";
import { NoCapabilityFoundError } from "../../../../core/error";
import { isExistingTabApp, isVSProject } from "../../../../common/projectSettingsHelper";
import {
  canAddApiConnection,
  canAddSso,
  canAddCICDWorkflows,
  isAadManifestEnabled,
  isDeployManifestEnabled,
  AppStudioScopes,
} from "../../../../common/tools";
import {
  isBotNotificationEnabled,
  isPreviewFeaturesEnabled,
} from "../../../../common/featureFlags";
import {
  ProgrammingLanguageQuestion,
  onChangeSelectionForCapabilities,
  validateCapabilities,
} from "../../../../core/question";
import { getDefaultString, getLocalizedString } from "../../../../common/localizeUtils";
import { Constants } from "../../../resource/aad/constants";
import { PluginBot } from "../../../resource/bot/resources/strings";

export async function getQuestionsForScaffolding(
  ctx: v2.Context,
  inputs: Inputs
): Promise<Result<QTreeNode | undefined, FxError>> {
  const node = new QTreeNode({
    name: "azure-solution-group",
    type: "func",
    func: (inputs: Inputs) => {
      inputs[AzureSolutionQuestionNames.Solution] = TeamsAppSolutionNameV2;
    },
  });

  if (!isV3()) {
    node.condition = {
      containsAny: [
        TabSPFxItem.id,
        TabOptionItem.id,
        BotOptionItem.id,
        NotificationOptionItem.id,
        CommandAndResponseOptionItem.id,
        MessageExtensionItem.id,
        ...(isAadManifestEnabled() ? [TabNonSsoItem.id] : []),
        M365SsoLaunchPageOptionItem.id,
        M365SearchAppOptionItem.id,
      ],
    };
    // 1.1.1 SPFX Tab
    const spfxPlugin: v2.ResourcePlugin = Container.get<v2.ResourcePlugin>(
      ResourcePluginsV2.SpfxPlugin
    );
    if (spfxPlugin.getQuestionsForScaffolding) {
      const res = await spfxPlugin.getQuestionsForScaffolding(ctx, inputs);
      if (res.isErr()) return res;
      if (res.value) {
        const spfxNode = res.value as QTreeNode;
        spfxNode.condition = {
          validFunc: (input: any, inputs?: Inputs) => {
            if (!inputs) {
              return "Invalid inputs";
            }
            const cap = inputs[AzureSolutionQuestionNames.Capabilities] as string[];
            if (cap.includes(TabSPFxItem.id)) {
              return undefined;
            }
            return "SPFx is not selected";
          },
        };
        if (spfxNode.data) node.addChild(spfxNode);
      }
    }
  } else {
    node.condition = { containsAny: [TabOptionItem.id, BotOptionItem.id, MessageExtensionItem.id] };
  }

  // 1.1.2 Azure Tab
  const tabRes = await getTabScaffoldQuestionsV2(
    ctx,
    inputs,
    !isPreviewFeaturesEnabled() && CLIPlatforms.includes(inputs.platform) // only CLI and CLI_HELP support azure-resources question
  );
  if (tabRes.isErr()) return tabRes;
  if (tabRes.value) {
    const tabNode = tabRes.value;
    tabNode.condition = {
      validFunc: (input: any, inputs?: Inputs) => {
        if (!inputs) {
          return "Invalid inputs";
        }
        const cap = inputs[AzureSolutionQuestionNames.Capabilities] as string[];
        if (
          cap.includes(TabOptionItem.id) ||
          (isAadManifestEnabled() && cap.includes(TabNonSsoItem.id))
        ) {
          return undefined;
        }
        return "Tab is not selected";
      },
    };
    node.addChild(tabNode);
  }

  // 1.2 Bot
  const botPlugin: v2.ResourcePlugin = Container.get<v2.ResourcePlugin>(
    ResourcePluginsV2.BotPlugin
  );
  if (botPlugin.getQuestionsForScaffolding) {
    const res = await botPlugin.getQuestionsForScaffolding(ctx, inputs);
    if (res.isErr()) return res;
    if (res.value) {
      // Create a parent node of the node returned by plugin to prevent overwriting node.condition.
      const botGroup = new QTreeNode({ type: "group" });
      botGroup.addChild(res.value);
      botGroup.condition = {
        validFunc: (input: any, inputs?: Inputs) => {
          if (!inputs) {
            return "Invalid inputs";
          }
          const cap = inputs[AzureSolutionQuestionNames.Capabilities] as string[];
          if (
            cap.includes(BotOptionItem.id) ||
            cap.includes(MessageExtensionItem.id) ||
            cap.includes(NotificationOptionItem.id) ||
            cap.includes(CommandAndResponseOptionItem.id)
          ) {
            return undefined;
          }
          return "Bot/Message Extension is not selected";
        },
      };
      node.addChild(botGroup);
    }
  }

  return ok(node);
}

export async function getQuestionsForScaffoldingPreview(
  ctx: v2.Context,
  inputs: Inputs
): Promise<Result<QTreeNode | undefined, FxError>> {
  const node = new QTreeNode({
    name: "azure-solution-group",
    type: "func",
    func: (inputs: Inputs) => {
      inputs[AzureSolutionQuestionNames.Solution] = TeamsAppSolutionNameV2;
    },
  });

  if (!isV3()) {
    node.condition = {
      enum: [
        TabSPFxItem.id,
        TabOptionItem.id,
        BotOptionItem.id,
        NotificationOptionItem.id,
        CommandAndResponseOptionItem.id,
        MessageExtensionItem.id,
        ...(isAadManifestEnabled() ? [TabNonSsoItem.id] : []),
        M365SsoLaunchPageOptionItem.id,
        M365SearchAppOptionItem.id,
      ],
    };

    // 1.1.1 SPFX Tab
    const spfxPlugin: v2.ResourcePlugin = Container.get<v2.ResourcePlugin>(
      ResourcePluginsV2.SpfxPlugin
    );
    if (spfxPlugin.getQuestionsForScaffolding) {
      const res = await spfxPlugin.getQuestionsForScaffolding(ctx, inputs);
      if (res.isErr()) return res;
      if (res.value) {
        const spfxNode = res.value as QTreeNode;
        spfxNode.condition = {
          validFunc: (input: any, inputs?: Inputs) => {
            if (!inputs) {
              return "Invalid inputs";
            }
            const cap = inputs[AzureSolutionQuestionNames.Capabilities] as string;
            if (cap === TabSPFxItem.id) {
              return undefined;
            }
            return "SPFx is not selected";
          },
        };
        if (spfxNode.data) node.addChild(spfxNode);
      }
    }
  } else {
    node.condition = { enum: [TabOptionItem.id, BotOptionItem.id, MessageExtensionItem.id] };
  }

  // 1.1.2 Azure Tab
  const tabRes = await getTabScaffoldQuestionsV2(
    ctx,
    inputs,
    !isPreviewFeaturesEnabled() && CLIPlatforms.includes(inputs.platform) // only CLI and CLI_HELP support azure-resources question
  );
  if (tabRes.isErr()) return tabRes;
  if (tabRes.value) {
    const tabNode = tabRes.value;
    tabNode.condition = {
      validFunc: (input: any, inputs?: Inputs) => {
        if (!inputs) {
          return "Invalid inputs";
        }
        const cap = inputs[AzureSolutionQuestionNames.Capabilities] as string;
        if (cap === TabOptionItem.id || cap === TabNonSsoItem.id) {
          return undefined;
        }
        return "Tab is not selected";
      },
    };
    node.addChild(tabNode);
  }

  // 1.2 Bot
  const botPlugin: v2.ResourcePlugin = Container.get<v2.ResourcePlugin>(
    ResourcePluginsV2.BotPlugin
  );
  if (botPlugin.getQuestionsForScaffolding) {
    const res = await botPlugin.getQuestionsForScaffolding(ctx, inputs);
    if (res.isErr()) return res;
    if (res.value) {
      // Create a parent node of the node returned by plugin to prevent overwriting node.condition.
      const botGroup = new QTreeNode({ type: "group" });
      botGroup.addChild(res.value);
      botGroup.condition = {
        validFunc: (input: any, inputs?: Inputs) => {
          if (!inputs) {
            return "Invalid inputs";
          }
          const cap = inputs[AzureSolutionQuestionNames.Capabilities] as string;
          if (
            cap === BotOptionItem.id ||
            cap === MessageExtensionItem.id ||
            cap === NotificationOptionItem.id ||
            cap === CommandAndResponseOptionItem.id
          ) {
            return undefined;
          }
          return "Bot/Message Extension is not selected";
        },
      };
      node.addChild(botGroup);
    }
  }

  return ok(node);
}

export async function getTabScaffoldQuestionsV2(
  ctx: v2.Context,
  inputs: Inputs,
  addAzureResource: boolean
): Promise<Result<QTreeNode | undefined, FxError>> {
  const tabNode = new QTreeNode({ type: "group" });

  //Frontend plugin
  const fehostPlugin: v2.ResourcePlugin = Container.get<v2.ResourcePlugin>(
    ResourcePluginsV2.FrontendPlugin
  );
  if (fehostPlugin.getQuestionsForScaffolding) {
    const res = await fehostPlugin.getQuestionsForScaffolding(ctx, inputs);
    if (res.isErr()) return res;
    if (res.value) {
      const frontendNode = res.value as QTreeNode;
      if (frontendNode.data) tabNode.addChild(frontendNode);
    }
  }

  if (addAzureResource) {
    const azureResourceNode = new QTreeNode(AzureResourcesQuestion);
    tabNode.addChild(azureResourceNode);
    const functionPlugin: v2.ResourcePlugin = Container.get<v2.ResourcePlugin>(
      ResourcePluginsV2.FunctionPlugin
    );
    //Azure Function
    if (functionPlugin.getQuestionsForScaffolding) {
      const res = await functionPlugin.getQuestionsForScaffolding(ctx, inputs);
      if (res.isErr()) return res;
      if (res.value) {
        const azure_function = res.value as QTreeNode;
        azure_function.condition = { minItems: 1 };
        if (azure_function.data) azureResourceNode.addChild(azure_function);
      }
    }
    const sqlPlugin: v2.ResourcePlugin = Container.get<v2.ResourcePlugin>(
      ResourcePluginsV2.SqlPlugin
    );
    //Azure SQL
    if (sqlPlugin.getQuestionsForScaffolding) {
      const res = await sqlPlugin.getQuestionsForScaffolding(ctx, inputs);
      if (res.isErr()) return res;
      if (res.value) {
        const azure_sql = res.value as QTreeNode;
        azure_sql.condition = { contains: AzureResourceSQL.id };
        if (azure_sql.data) azureResourceNode.addChild(azure_sql);
      }
    }
  }
  return ok(tabNode);
}

export function getPluginCLIName(name: string): string {
  const pluginPrefix = "fx-resource-";
  if (name === ResourcePlugins.Aad) {
    return "aad-manifest";
  } else if (name === ResourcePlugins.AppStudio) {
    return "manifest";
  } else {
    return name.replace(pluginPrefix, "");
  }
}

export async function getQuestions(
  ctx: v2.Context,
  inputs: Inputs,
  envInfo: v2.DeepReadonly<v2.EnvInfoV2>,
  tokenProvider: TokenProvider
): Promise<Result<QTreeNode | undefined, FxError>> {
  const stage = inputs.stage;
  if (!stage) {
    return err(new InvalidInputError(SolutionSource, "inputs.stage", "undefined"));
  }
  const isDynamicQuestion = DynamicPlatforms.includes(inputs.platform);
  const node = new QTreeNode({ type: "group" });
  const solutionSettings = ctx.projectSetting.solutionSettings as AzureSolutionSettings;
  if (stage === Stage.provision) {
    if (isDynamicQuestion) {
      const provisioned = checkWetherProvisionSucceeded(envInfo.state);
      if (provisioned) return ok(undefined);
    }
    let plugins: v2.ResourcePlugin[] = [];
    if (isDynamicQuestion) {
      plugins = getSelectedPlugins(ctx.projectSetting);
    } else {
      plugins = getAllV2ResourcePlugins();
      node.addChild(new QTreeNode(AskSubscriptionQuestion));
    }
    for (const plugin of plugins) {
      if (plugin.getQuestions) {
        const getQuestionRes = await plugin.getQuestions(ctx, inputs, envInfo, tokenProvider);
        if (getQuestionRes.isErr()) return getQuestionRes;
        if (getQuestionRes.value) {
          const subnode = getQuestionRes.value as QTreeNode;
          node.addChild(subnode);
        }
      }
    }
  } else if (stage === Stage.deploy) {
    if (inputs.platform === Platform.VSCode && inputs[Constants.INCLUDE_AAD_MANIFEST] === "yes") {
      return ok(node);
    }

    if (isDynamicQuestion) {
      const isAzure = isAzureProject(solutionSettings);
      const provisioned = checkWetherProvisionSucceeded(envInfo.state);
      if (isAzure && !provisioned) {
        return err(
          new UserError({
            source: SolutionSource,
            name: SolutionError.CannotDeployBeforeProvision,
            message: getDefaultString("core.deploy.FailedToDeployBeforeProvision"),
            displayMessage: getLocalizedString("core.deploy.FailedToDeployBeforeProvision"),
            helpLink: HelpLinks.WhyNeedProvision,
          })
        );
      }
    }
    let plugins: v2.ResourcePlugin[] = [];
    if (isDynamicQuestion) {
      plugins = getSelectedPlugins(ctx.projectSetting);
    } else {
      plugins = getAllV2ResourcePlugins();
    }

    if (isDeployManifestEnabled() && inputs.platform === Platform.VSCode) {
      plugins = plugins.filter((plugin) => plugin.name !== ResourcePlugins.AppStudio);
    }

    if (
      isAadManifestEnabled() &&
      (inputs.platform === Platform.CLI_HELP || inputs.platform === Platform.CLI)
    ) {
      plugins = plugins.filter((plugin) => !!plugin.deploy);
    } else {
      plugins = plugins.filter((plugin) => !!plugin.deploy && plugin.displayName !== "AAD");
    }

    if (plugins.length === 0 && inputs[Constants.INCLUDE_AAD_MANIFEST] !== "yes") {
      return err(new NoCapabilityFoundError(Stage.deploy));
    }

    // On VS, users are not expected to select plugins to deploy.
    if (!isVSProject(ctx.projectSetting)) {
      const options: OptionItem[] = plugins.map((plugin) => {
        const item: OptionItem = {
          id: plugin.name,
          label: plugin.displayName,
          cliName: getPluginCLIName(plugin.name),
        };
        return item;
      });

      const selectQuestion = DeployPluginSelectQuestion;
      selectQuestion.staticOptions = options;
      selectQuestion.default = options.map((i) => i.id);
      const pluginSelection = new QTreeNode(selectQuestion);
      node.addChild(pluginSelection);

      for (const plugin of plugins) {
        if (plugin.getQuestions) {
          const getQuestionRes = await plugin.getQuestions(ctx, inputs, envInfo, tokenProvider);
          if (getQuestionRes.isErr()) return getQuestionRes;
          if (getQuestionRes.value) {
            const subnode = getQuestionRes.value as QTreeNode;
            subnode.condition = { contains: plugin.name };
            if (subnode.data) pluginSelection.addChild(subnode);
          }
        }
      }
    }
  } else if (stage === Stage.publish) {
    if (isDynamicQuestion) {
      const isAzure = isAzureProject(solutionSettings);
      const provisioned = checkWetherProvisionSucceeded(envInfo.state);
      if (!provisioned) {
        const errorMsg = isAzure
          ? getLocalizedString("core.publish.FailedToPublishBeforeProvision")
          : getLocalizedString("core.publish.SPFxAskProvisionBeforePublish");
        const defaultMsg = isAzure
          ? getDefaultString("core.publish.FailedToPublishBeforeProvision")
          : getDefaultString("core.publish.SPFxAskProvisionBeforePublish");
        return err(
          new UserError({
            source: SolutionSource,
            name: SolutionError.CannotPublishBeforeProvision,
            message: defaultMsg,
            displayMessage: errorMsg,
            helpLink: HelpLinks.WhyNeedProvision,
          })
        );
      }
    }
    let plugins: v2.ResourcePlugin[] = [];
    if (isDynamicQuestion) {
      plugins = getSelectedPlugins(ctx.projectSetting);
    } else {
      plugins = getAllV2ResourcePlugins();
    }
    plugins = plugins.filter((plugin) => !!plugin.publishApplication);
    for (const plugin of plugins) {
      if (plugin.getQuestions) {
        const getQuestionRes = await plugin.getQuestions(ctx, inputs, envInfo, tokenProvider);
        if (getQuestionRes.isErr()) return getQuestionRes;
        if (getQuestionRes.value) {
          const subnode = getQuestionRes.value as QTreeNode;
          node.addChild(subnode);
        }
      }
    }
  } else if (stage === Stage.grantPermission) {
    if (isDynamicQuestion) {
      const jsonObjectRes = await tokenProvider.m365TokenProvider.getJsonObject({
        scopes: AppStudioScopes,
      });
      const jsonObject = jsonObjectRes.isOk() ? jsonObjectRes.value : undefined;
      node.addChild(new QTreeNode(getUserEmailQuestion((jsonObject as any).upn)));
    }
  }
  return ok(node);
}

export async function getQuestionsForUserTask(
  ctx: v2.Context,
  inputs: Inputs,
  func: Func,
  envInfo: v2.DeepReadonly<v2.EnvInfoV2>,
  tokenProvider: TokenProvider
): Promise<Result<QTreeNode | undefined, FxError>> {
  const namespace = func.namespace;
  const array = namespace.split("/");
  if (func.method === "addCapability") {
    return await getQuestionsForAddCapability(ctx, inputs, func, envInfo, tokenProvider);
  }
  if (func.method === "addResource") {
    return await getQuestionsForAddResource(ctx, inputs, func, envInfo, tokenProvider);
  }
  if (func.method === "addFeature") {
    return await getQuestionsForAddFeature(ctx, inputs, func, envInfo, tokenProvider);
  }
  if (array.length == 2) {
    const pluginName = array[1];
    const pluginMap = getAllV2ResourcePluginMap();
    const plugin = pluginMap.get(pluginName);
    if (plugin && plugin.getQuestionsForUserTask) {
      return await plugin.getQuestionsForUserTask(ctx, inputs, func, envInfo, tokenProvider);
    }
  }
  return ok(undefined);
}

export async function getQuestionsForAddCapability(
  ctx: v2.Context,
  inputs: Inputs,
  func: Func,
  envInfo: v2.DeepReadonly<v2.EnvInfoV2>,
  tokenProvider: TokenProvider
): Promise<Result<QTreeNode | undefined, FxError>> {
  if (ctx.projectSetting.isM365) {
    return err(
      new UserError(
        SolutionSource,
        SolutionError.AddCapabilityNotSupport,
        getDefaultString("core.addCapability.notSupportedForM365Project"),
        getLocalizedString("core.addCapability.notSupportedForM365Project")
      )
    );
  }
  const settings = ctx.projectSetting.solutionSettings as AzureSolutionSettings | undefined;
  const addCapQuestion: MultiSelectQuestion = {
    name: AzureSolutionQuestionNames.Capabilities,
    title: isBotNotificationEnabled() ? "Capabilities" : "Choose capabilities",
    type: "multiSelect",
    staticOptions: [],
    default: [],
    validation: {
      validFunc: validateCapabilities,
    },
    onDidChangeSelection: onChangeSelectionForCapabilities,
  };
  const isDynamicQuestion = DynamicPlatforms.includes(inputs.platform);
  if (!isDynamicQuestion) {
    // For CLI_HELP
    addCapQuestion.staticOptions = [
      ...(isBotNotificationEnabled() ? [TabNewUIOptionItem] : [TabOptionItem]),
      ...[BotOptionItem],
      ...(isBotNotificationEnabled() ? [NotificationOptionItem, CommandAndResponseOptionItem] : []),
      ...(isBotNotificationEnabled() ? [MessageExtensionNewUIItem] : [MessageExtensionItem]),
      ...(isAadManifestEnabled() ? [TabNonSsoItem] : []),
    ];
    const addCapNode = new QTreeNode(addCapQuestion);
    if (isBotNotificationEnabled()) {
      // Hardcoded to call bot plugin to get notification trigger questions.
      // Originally, v2 solution will not call getQuestionForUserTask of plugins on addCapability.
      // V3 will not need this hardcoding.
      const pluginMap = getAllV2ResourcePluginMap();
      const plugin = pluginMap.get(PluginNames.BOT);
      if (plugin && plugin.getQuestionsForUserTask) {
        const result = await plugin.getQuestionsForUserTask(
          ctx,
          inputs,
          func,
          envInfo,
          tokenProvider
        );
        if (result.isErr()) {
          return result;
        }
        const botQuestionNode = result.value;
        if (botQuestionNode) {
          addCapNode.addChild(botQuestionNode);
        }
      }
    }
    return ok(addCapNode);
  }
  const canProceed = canAddCapability(settings, ctx.telemetryReporter);
  if (canProceed.isErr()) {
    return err(canProceed.error);
  }
  const appStudioPlugin = Container.get<AppStudioPluginV3>(BuiltInFeaturePluginNames.appStudio);
  const tabExceedRes = await appStudioPlugin.capabilityExceedLimit(
    ctx,
    inputs as v2.InputsWithProjectPath,
    "staticTab"
  );
  if (tabExceedRes.isErr()) {
    return err(tabExceedRes.error);
  }
  const isTabAddable = !tabExceedRes.value;
  const botExceedRes = await appStudioPlugin.capabilityExceedLimit(
    ctx,
    inputs as v2.InputsWithProjectPath,
    "Bot"
  );
  if (botExceedRes.isErr()) {
    return err(botExceedRes.error);
  }
  const isBotAddable = !botExceedRes.value;
  const meExceedRes = await appStudioPlugin.capabilityExceedLimit(
    ctx,
    inputs as v2.InputsWithProjectPath,
    "MessageExtension"
  );
  if (meExceedRes.isErr()) {
    return err(meExceedRes.error);
  }
  // for the new bot, messaging extension and other bots are mutally exclusive
  const isMEAddable = !meExceedRes.value && (!isBotNotificationEnabled() || isBotAddable);
  if (!(isTabAddable || isBotAddable || isMEAddable)) {
    ctx.userInteraction?.showMessage(
      "error",
      getLocalizedString("core.addCapability.exceedMaxLimit"),
      false
    );
    return ok(undefined);
  }
  const options = [];
  if (isBotAddable) {
    if (isBotNotificationEnabled()) {
      options.push(CommandAndResponseOptionItem);
      options.push(NotificationOptionItem);
      options.push(BotOptionItem);
    } else {
      options.push(BotOptionItem);
    }
  }
  const tabOptionItem = isBotNotificationEnabled() ? TabNewUIOptionItem : TabOptionItem;
  if (isTabAddable) {
    if (!isAadManifestEnabled()) {
      options.push(tabOptionItem);
    } else {
      if (!settings?.capabilities.includes(TabOptionItem.id)) {
        options.push(TabNonSsoItem, tabOptionItem);
      } else {
        options.push(
          settings?.capabilities.includes(TabSsoItem.id) ? tabOptionItem : TabNonSsoItem
        );
      }
    }
  }
  if (isMEAddable) {
    options.push(isBotNotificationEnabled() ? MessageExtensionNewUIItem : MessageExtensionItem);
  }

  addCapQuestion.staticOptions = options;
  const addCapNode = new QTreeNode(addCapQuestion);

  // // mini app can add SPFx tab
  // if (!settings) {
  //   options.push(TabSPFxItem);
  //   const spfxPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.SpfxPlugin);
  //   if (spfxPlugin && spfxPlugin.getQuestionsForScaffolding) {
  //     const result = await spfxPlugin.getQuestionsForScaffolding(ctx, inputs);
  //     if (result.isErr()) {
  //       return result;
  //     }
  //     const spfxQuestionNode = result.value;
  //     if (spfxQuestionNode) {
  //       spfxQuestionNode.condition = { contains: TabSPFxItem.id };
  //       addCapNode.addChild(spfxQuestionNode);
  //     }
  //   }
  // }

  if (isBotNotificationEnabled()) {
    // Hardcoded to call bot plugin to get notification trigger questions.
    // Originally, v2 solution will not call getQuestionForUserTask of plugins on addCapability.
    // V3 will not need this hardcoding.
    const pluginMap = getAllV2ResourcePluginMap();
    const plugin = pluginMap.get(PluginNames.BOT);
    if (plugin && plugin.getQuestionsForUserTask) {
      const result = await plugin.getQuestionsForUserTask(
        ctx,
        inputs,
        func,
        envInfo,
        tokenProvider
      );
      if (result.isErr()) {
        return result;
      }
      const botQuestionNode = result.value;
      if (botQuestionNode) {
        addCapNode.addChild(botQuestionNode);
      }
    }
  }

  if (!ctx.projectSetting.programmingLanguage) {
    // Language
    const programmingLanguage = new QTreeNode(ProgrammingLanguageQuestion);
    addCapNode.addChild(programmingLanguage);
  }
  return ok(addCapNode);
}

export async function getQuestionsForAddResource(
  ctx: v2.Context,
  inputs: Inputs,
  func: Func,
  envInfo: v2.DeepReadonly<v2.EnvInfoV2>,
  tokenProvider: TokenProvider
): Promise<Result<QTreeNode | undefined, FxError>> {
  const settings = ctx.projectSetting.solutionSettings as AzureSolutionSettings | undefined;
  const isDynamicQuestion = DynamicPlatforms.includes(inputs.platform);
  let addQuestion: MultiSelectQuestion;
  if (!isDynamicQuestion) {
    addQuestion = createAddAzureResourceQuestion(false, false, false, false);
  } else {
    if (!settings) {
      return err(new NoCapabilityFoundError(Stage.addResource));
    }
    const alreadyHaveFunction = settings.azureResources.includes(AzureResourceFunction.id);
    const alreadyHaveSQL = settings.azureResources.includes(AzureResourceSQL.id);
    const alreadyHaveAPIM = settings.azureResources.includes(AzureResourceApim.id);
    const alreadyHaveKeyVault = settings.azureResources.includes(AzureResourceKeyVault.id);
    addQuestion = createAddAzureResourceQuestion(
      alreadyHaveFunction,
      alreadyHaveSQL,
      alreadyHaveAPIM,
      alreadyHaveKeyVault
    );
    const canProceed = canAddResource(ctx.projectSetting, ctx.telemetryReporter);
    if (canProceed.isErr()) {
      return err(canProceed.error);
    }
  }
  const addAzureResourceNode = new QTreeNode(addQuestion);
  //traverse plugins' getQuestionsForUserTask
  const pluginsWithResources = [
    [ResourcePluginsV2.FunctionPlugin, AzureResourceFunction.id],
    [ResourcePluginsV2.SqlPlugin, AzureResourceSQL.id],
    [ResourcePluginsV2.ApimPlugin, AzureResourceApim.id],
    [ResourcePluginsV2.KeyVaultPlugin, AzureResourceKeyVault.id],
  ];
  for (const pair of pluginsWithResources) {
    const pluginName = pair[0];
    const resourceName = pair[1];
    const plugin: v2.ResourcePlugin = Container.get<v2.ResourcePlugin>(pluginName);
    if (plugin.getQuestionsForUserTask) {
      const res = await plugin.getQuestionsForUserTask(ctx, inputs, func, envInfo, tokenProvider);
      if (res.isErr()) return res;
      if (res.value) {
        const node = res.value as QTreeNode;
        node.condition = { contains: resourceName };
        if (node.data) addAzureResourceNode.addChild(node);
      }
    }
  }
  return ok(addAzureResourceNode);
}

async function getStaticOptionsForAddCapability(
  ctx: v2.Context,
  inputs: Inputs,
  settings?: AzureSolutionSettings
): Promise<Result<OptionItem[], FxError>> {
  if (inputs.platform === Platform.CLI_HELP) {
    const options: OptionItem[] = [];
    options.push(NotificationOptionItem);
    options.push(CommandAndResponseOptionItem);
    options.push(TabNewUIOptionItem, TabNonSsoItem);
    options.push(BotNewUIOptionItem);
    options.push(MessageExtensionNewUIItem);
    return ok(options);
  }
  const appStudioPlugin = Container.get<AppStudioPluginV3>(BuiltInFeaturePluginNames.appStudio);
  const tabExceedRes = await appStudioPlugin.capabilityExceedLimit(
    ctx,
    inputs as v2.InputsWithProjectPath,
    "staticTab"
  );
  if (tabExceedRes.isErr()) {
    return err(tabExceedRes.error);
  }
  const isTabAddable = !tabExceedRes.value;
  const botExceedRes = await appStudioPlugin.capabilityExceedLimit(
    ctx,
    inputs as v2.InputsWithProjectPath,
    "Bot"
  );
  if (botExceedRes.isErr()) {
    return err(botExceedRes.error);
  }

  const hasMe = settings?.capabilities.includes(MessageExtensionItem.id);
  const isScenarioBotAddable = !botExceedRes.value && !hasMe;
  const isDefaultBotAddable = !botExceedRes.value;
  const meExceedRes = await appStudioPlugin.capabilityExceedLimit(
    ctx,
    inputs as v2.InputsWithProjectPath,
    "MessageExtension"
  );
  if (meExceedRes.isErr()) {
    return err(meExceedRes.error);
  }
  // For the new bot, messaging extension and other bots are mutally exclusive.
  // For the old bot, messaging extension can be added when bot exists.
  const botCapabilities =
    ctx.projectSetting.pluginSettings?.[PluginNames.BOT]?.[PluginBot.BOT_CAPABILITIES];
  const hasNewBot = Array.isArray(botCapabilities) && botCapabilities.length > 0;
  const isMEAddable = isBotNotificationEnabled()
    ? !meExceedRes.value && !hasNewBot
    : !meExceedRes.value;
  if (!(isTabAddable || isDefaultBotAddable || isScenarioBotAddable || isMEAddable)) {
    ctx.userInteraction?.showMessage(
      "error",
      getLocalizedString("core.addCapability.exceedMaxLimit"),
      false
    );
    return ok([]);
  }

  const options: OptionItem[] = [];
  if (isScenarioBotAddable) {
    options.push(NotificationOptionItem);
    options.push(CommandAndResponseOptionItem);
  }
  if (isTabAddable) {
    if (!settings?.capabilities.includes(TabOptionItem.id)) {
      options.push(TabNewUIOptionItem, TabNonSsoItem);
    } else {
      options.push(
        settings?.capabilities.includes(TabSsoItem.id) ? TabNewUIOptionItem : TabNonSsoItem
      );
    }
  }
  if (isDefaultBotAddable) {
    options.push(BotNewUIOptionItem);
  }
  if (isMEAddable) {
    options.push(MessageExtensionNewUIItem);
  }
  return ok(options);
}

/**
 * Combines the options of AddCapability and AddResource.
 * Only works for VS Code new UI with Preview feature flag enabled.
 */
export async function getQuestionsForAddFeature(
  ctx: v2.Context,
  inputs: Inputs,
  func: Func,
  envInfo: v2.DeepReadonly<v2.EnvInfoV2>,
  tokenProvider: TokenProvider
): Promise<Result<QTreeNode | undefined, FxError>> {
  const settings = ctx.projectSetting.solutionSettings as AzureSolutionSettings | undefined;
  const options: OptionItem[] = [];
  const addFeatureQuestion: SingleSelectQuestion = {
    name: AzureSolutionQuestionNames.Features,
    title: getLocalizedString("core.addFeatureQuestion.title"),
    type: "singleSelect",
    staticOptions: [],
  };
  // check and generate capability options
  const canAddCapabilityResult = canAddCapability(settings, ctx.telemetryReporter);
  if (canAddCapabilityResult.isOk() && !ctx.projectSetting.isM365) {
    const optionsResult = await getStaticOptionsForAddCapability(ctx, inputs, settings);
    if (optionsResult.isErr()) {
      return err(optionsResult.error);
    }
    options.push(...optionsResult.value);
  }
  // check and generate cloud resource options
  if (inputs.platform === Platform.CLI_HELP) {
    options.push(...createAddCloudResourceOptions(false, false));
  } else {
    const canAddResourceResult = canAddResource(ctx.projectSetting, ctx.telemetryReporter);
    if (canAddResourceResult.isOk()) {
      // resources
      if (!settings) {
        return err(new NoCapabilityFoundError(Stage.addResource));
      }
      const alreadyHaveAPIM = settings.azureResources.includes(AzureResourceApim.id);
      const alreadyHaveKeyVault = settings.azureResources.includes(AzureResourceKeyVault.id);
      const addResourceOptions = createAddCloudResourceOptions(
        alreadyHaveAPIM,
        alreadyHaveKeyVault
      );
      options.push(...addResourceOptions);
    }
  }

  // check and generate additional feature options
  if (inputs.platform === Platform.CLI_HELP || canAddSso(ctx.projectSetting)) {
    options.push(SingleSignOnOptionItem);
  }
  const isApiConnectionAddable = canAddApiConnection(settings);
  if (inputs.platform === Platform.CLI_HELP || isApiConnectionAddable) {
    options.push(ApiConnectionOptionItem);
  }

  const isCicdAddable = await canAddCICDWorkflows(inputs, ctx);
  if (inputs.platform === Platform.CLI_HELP || isCicdAddable) {
    options.push(CicdOptionItem);
  }

  addFeatureQuestion.staticOptions = options;
  const addFeatureNode = new QTreeNode(addFeatureQuestion);

  if (inputs.platform !== Platform.CLI_HELP && !ctx.projectSetting.programmingLanguage) {
    // Language
    const programmingLanguage = new QTreeNode(ProgrammingLanguageQuestion);
    programmingLanguage.condition = {
      enum: [
        NotificationOptionItem.id,
        CommandAndResponseOptionItem.id,
        TabNewUIOptionItem.id,
        TabNonSsoItem.id,
        BotNewUIOptionItem.id,
        MessageExtensionItem.id,
        SingleSignOnOptionItem.id,
      ],
    };
    addFeatureNode.addChild(programmingLanguage);
  }

  // traverse plugins' getQuestionsForUserTask
  const pluginsWithResources = [
    [ResourcePluginsV2.BotPlugin, BotNewUIOptionItem.id],
    [ResourcePluginsV2.FunctionPlugin, AzureResourceFunction.id],
  ];
  if (inputs.platform === Platform.CLI_HELP || isCicdAddable) {
    pluginsWithResources.push([ResourcePluginsV2.CICDPlugin, CicdOptionItem.id]);
  }
  if (inputs.platform === Platform.CLI_HELP || isApiConnectionAddable) {
    pluginsWithResources.push([ResourcePluginsV2.ApiConnectorPlugin, ApiConnectionOptionItem.id]);
  }
  const alreadyHaveFunction = settings?.azureResources.includes(AzureResourceFunction.id);
  for (const pair of pluginsWithResources) {
    const pluginName = pair[0];
    const resourceName = pair[1];
    const plugin: v2.ResourcePlugin = Container.get<v2.ResourcePlugin>(pluginName);
    if (plugin.getQuestionsForUserTask) {
      const res = await plugin.getQuestionsForUserTask(ctx, inputs, func, envInfo, tokenProvider);
      if (res.isErr()) return res;
      if (res.value) {
        const node = res.value as QTreeNode;
        if (!node.condition) {
          if (resourceName !== AzureResourceFunction.id) {
            node.condition = { equals: resourceName };
          } else {
            // Azure Function question is related to APIM and SQL
            node.condition = {
              validFunc: (input: string, inputs?: Inputs) => {
                if (input === AzureResourceFunction.id) {
                  return undefined;
                }
                if (
                  !alreadyHaveFunction &&
                  (input === AzureResourceSQL.id || input === AzureResourceApim.id)
                ) {
                  return undefined;
                }
                return "Function related is not selected";
              },
            };
          }
        }
        if (node.data) {
          addFeatureNode.addChild(node);
        }
      }
    }
  }

  return ok(addFeatureNode);
}
