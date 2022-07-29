// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  CLIPlatforms,
  DynamicPlatforms,
  err,
  FxError,
  Inputs,
  InputsWithProjectPath,
  ok,
  OptionItem,
  Platform,
  Plugin,
  ProjectSettingsV3,
  QTreeNode,
  Result,
  SingleSelectQuestion,
  Stage,
  UserError,
  v2,
  v3,
} from "@microsoft/teamsfx-api";
import Container from "typedi";
import { isVSProject } from "../common/projectSettingsHelper";
import { HelpLinks } from "../common/constants";
import { getDefaultString, getLocalizedString } from "../common/localizeUtils";
import {
  hasAAD,
  hasAPIM,
  hasAzureResourceV3,
  hasBot,
  hasApi,
  hasKeyVault,
  hasTab,
} from "../common/projectSettingsHelperV3";
import { canAddCICDWorkflows } from "../common/tools";
import { ComponentNames } from "./constants";
import { ComponentName2pluginName } from "./migrate";
import { readAppManifest } from "./resource/appManifest/utils";
import { getComponent, getQuestionsV3 } from "./workflow";
import { STATIC_TABS_MAX_ITEMS } from "../plugins/resource/appstudio/constants";
import {
  createHostTypeTriggerQuestion,
  getConditionOfNotificationTriggerQuestion,
  showNotificationTriggerCondition,
} from "../plugins/resource/bot/question";
import {
  ApiConnectionOptionItem,
  AzureResourceApimNewUI,
  AzureResourceFunctionNewUI,
  AzureResourceKeyVaultNewUI,
  AzureResourceSQLNewUI,
  AzureSolutionQuestionNames,
  BotFeatureIds,
  BotNewUIOptionItem,
  CicdOptionItem,
  CommandAndResponseOptionItem,
  DeployPluginSelectQuestion,
  MessageExtensionItem,
  MessageExtensionNewUIItem,
  NotificationOptionItem,
  SingleSignOnOptionItem,
  TabFeatureIds,
  TabNewUIOptionItem,
  TabNonSsoItem,
} from "../plugins/solution/fx-solution/question";
import { getPluginCLIName } from "../plugins/solution/fx-solution/v2/getQuestions";
import { checkWetherProvisionSucceeded } from "../plugins/solution/fx-solution/v2/utils";
import { NoCapabilityFoundError } from "../core/error";
import { ProgrammingLanguageQuestion } from "../core/question";
import { createContextV3 } from "./utils";
import { isCLIDotNetEnabled } from "../common/featureFlags";
import { Runtime } from "../plugins/resource/bot/v2/enum";
import { getPlatformRuntime } from "../plugins/resource/bot/v2/mapping";
import { buildQuestionNode } from "./resource/azureSql/questions";
import { functionNameQuestion } from "../plugins/resource/function/question";
import { ApiConnectorImpl } from "../plugins/resource/apiconnector/plugin";
import { addCicdQuestion } from "./feature/cicd";

export async function getQuestionsForProvisionV3(
  inputs: Inputs
): Promise<Result<QTreeNode | undefined, FxError>> {
  if (inputs.platform === Platform.CLI_HELP) {
    return ok(buildQuestionNode());
  }
  return ok(undefined);
}

export async function getQuestionsForDeployV3(
  ctx: v2.Context,
  inputs: Inputs,
  envInfo?: v3.EnvInfoV3
): Promise<Result<QTreeNode | undefined, FxError>> {
  //VS project has no selection interaction, and will deploy all selectable components by default.
  if (isVSProject(ctx.projectSetting)) {
    return ok(undefined);
  }
  const isDynamicQuestion = DynamicPlatforms.includes(inputs.platform);
  const projectSetting = ctx.projectSetting as ProjectSettingsV3;
  const deployableComponents = [
    ComponentNames.TeamsTab,
    ComponentNames.TeamsBot,
    ComponentNames.TeamsApi,
    ComponentNames.APIM,
    ComponentNames.AppManifest,
  ];

  if (CLIPlatforms.includes(inputs.platform)) {
    deployableComponents.push(ComponentNames.AadApp);
  }

  let selectableComponents: string[];
  if (!isDynamicQuestion) {
    selectableComponents = deployableComponents;
  } else {
    const hasAzureResource = hasAzureResourceV3(projectSetting);
    const provisioned = checkWetherProvisionSucceeded(envInfo!.state);
    if (hasAzureResource && !provisioned) {
      return err(
        new UserError({
          source: "Solution",
          name: "CannotDeployBeforeProvision",
          message: getDefaultString("core.deploy.FailedToDeployBeforeProvision"),
          displayMessage: getLocalizedString("core.deploy.FailedToDeployBeforeProvision"),
          helpLink: HelpLinks.WhyNeedProvision,
        })
      );
    }
    selectableComponents = projectSetting.components
      .filter((component) => component.deploy && deployableComponents.includes(component.name))
      .map((component) => component.name) as string[];
    if (CLIPlatforms.includes(inputs.platform)) {
      deployableComponents.push(ComponentNames.AppManifest);
    }
  }
  const options = selectableComponents.map((c) => {
    const pluginName = ComponentName2pluginName(c);
    const plugin = Container.get<Plugin>(pluginName);
    const item: OptionItem = {
      id: pluginName,
      label: plugin.displayName,
      cliName: getPluginCLIName(plugin.name),
    };
    return item;
  });
  if (options.length === 0) {
    return err(new NoCapabilityFoundError(Stage.deploy));
  }
  const selectQuestion = DeployPluginSelectQuestion;
  selectQuestion.staticOptions = options;
  selectQuestion.default = options.map((i) => i.id);
  return ok(new QTreeNode(selectQuestion));
}

export async function getQuestionsForAddFeatureV3(
  ctx: v2.Context,
  inputs: Inputs
): Promise<Result<QTreeNode | undefined, FxError>> {
  const question: SingleSelectQuestion = {
    name: AzureSolutionQuestionNames.Features,
    title: getLocalizedString("core.addFeatureQuestion.title"),
    type: "singleSelect",
    staticOptions: [],
  };
  const options: OptionItem[] = [];
  question.staticOptions = options;
  if (inputs.platform === Platform.CLI_HELP) {
    options.push(NotificationOptionItem);
    options.push(CommandAndResponseOptionItem);
    options.push(BotNewUIOptionItem);
    options.push(TabNewUIOptionItem, TabNonSsoItem);
    options.push(MessageExtensionNewUIItem);
    options.push(AzureResourceApimNewUI);
    options.push(AzureResourceSQLNewUI);
    options.push(AzureResourceFunctionNewUI);
    options.push(AzureResourceKeyVaultNewUI);
    options.push(SingleSignOnOptionItem);
    options.push(ApiConnectionOptionItem);
    options.push(CicdOptionItem);
    const addFeatureNode = new QTreeNode(question);
    const triggerNodeRes = await getNotificationTriggerQuestionNode(inputs);
    if (triggerNodeRes.isErr()) return err(triggerNodeRes.error);
    if (triggerNodeRes.value) {
      addFeatureNode.addChild(triggerNodeRes.value);
    }
    return ok(addFeatureNode);
  }
  // check capability options
  const manifestRes = await readAppManifest(inputs.projectPath!);
  if (manifestRes.isErr()) return err(manifestRes.error);
  const manifest = manifestRes.value;
  const canAddTab = manifest.staticTabs!.length < STATIC_TABS_MAX_ITEMS;
  const botExceedLimit = manifest.bots!.length > 0;
  const meExceedLimit = manifest.composeExtensions!.length > 0;
  const projectSettingsV3 = ctx.projectSetting as ProjectSettingsV3;
  const teamsBot = getComponent(ctx.projectSetting as ProjectSettingsV3, ComponentNames.TeamsBot);
  const alreadyHasNewBot =
    teamsBot?.capabilities?.includes("notification") ||
    teamsBot?.capabilities?.includes("command-response");
  if (!botExceedLimit && !alreadyHasNewBot) {
    options.push(NotificationOptionItem);
    options.push(CommandAndResponseOptionItem);
    options.push(BotNewUIOptionItem);
  }
  if (canAddTab) {
    if (!hasTab(projectSettingsV3)) {
      options.push(TabNewUIOptionItem, TabNonSsoItem);
    } else {
      options.push(hasAAD(projectSettingsV3) ? TabNewUIOptionItem : TabNonSsoItem);
    }
  }
  if (!meExceedLimit && !alreadyHasNewBot) {
    options.push(MessageExtensionNewUIItem);
  }
  // check cloud resource options
  if (!hasAPIM(projectSettingsV3)) {
    options.push(AzureResourceApimNewUI);
  }
  options.push(AzureResourceSQLNewUI);
  if (!hasKeyVault(projectSettingsV3)) {
    options.push(AzureResourceKeyVaultNewUI);
  }
  if (!hasAAD(projectSettingsV3)) {
    options.push(SingleSignOnOptionItem);
  }
  if (hasBot(projectSettingsV3) || hasApi(projectSettingsV3)) {
    options.push(ApiConnectionOptionItem);
  }
  // function can always be added
  options.push(AzureResourceFunctionNewUI);
  const isCicdAddable = await canAddCICDWorkflows(inputs, ctx);
  if (isCicdAddable) {
    options.push(CicdOptionItem);
  }
  const addFeatureNode = new QTreeNode(question);
  const triggerNodeRes = await getNotificationTriggerQuestionNode(inputs);
  if (triggerNodeRes.isErr()) return err(triggerNodeRes.error);
  if (triggerNodeRes.value) {
    addFeatureNode.addChild(triggerNodeRes.value);
  }
  if (!ctx.projectSetting.programmingLanguage) {
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
        SingleSignOnOptionItem.id, // adding sso means adding sample codes
      ],
    };
    addFeatureNode.addChild(programmingLanguage);
  }
  return ok(addFeatureNode);
}

export async function getQuestionsForAddResourceV3(
  ctx: v2.Context,
  inputs: Inputs
): Promise<Result<QTreeNode | undefined, FxError>> {
  const question: SingleSelectQuestion = {
    name: AzureSolutionQuestionNames.AddResources,
    title: getLocalizedString("core.addFeatureQuestion.title"),
    type: "singleSelect",
    staticOptions: [],
  };
  const options: OptionItem[] = [];
  question.staticOptions = options;
  if (inputs.platform === Platform.CLI_HELP) {
    options.push(AzureResourceApimNewUI);
    options.push(AzureResourceSQLNewUI);
    options.push(AzureResourceFunctionNewUI);
    options.push(AzureResourceKeyVaultNewUI);
    const addFeatureNode = new QTreeNode(question);
    return ok(addFeatureNode);
  }
  const projectSettingsV3 = ctx.projectSetting as ProjectSettingsV3;
  if (!hasAPIM(projectSettingsV3)) {
    options.push(AzureResourceApimNewUI);
  }
  options.push(AzureResourceSQLNewUI);
  if (!hasKeyVault(projectSettingsV3)) {
    options.push(AzureResourceKeyVaultNewUI);
  }
  // function can always be added
  options.push(AzureResourceFunctionNewUI);
  const addFeatureNode = new QTreeNode(question);
  const triggerNodeRes = await getNotificationTriggerQuestionNode(inputs);
  if (triggerNodeRes.isErr()) return err(triggerNodeRes.error);
  if (triggerNodeRes.value) {
    addFeatureNode.addChild(triggerNodeRes.value);
  }
  if (!ctx.projectSetting.programmingLanguage) {
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
        SingleSignOnOptionItem.id, // adding sso means adding sample codes
      ],
    };
    addFeatureNode.addChild(programmingLanguage);
  }
  return ok(addFeatureNode);
}

export enum FeatureId {
  Tab = "Tab",
  TabNonSso = "TabNonSso",
  Notification = "Notification",
  CommandAndResponse = "command-bot",
  Bot = "Bot",
  MessagingExtension = "MessagingExtension",
  function = "function",
  apim = "apim",
  sql = "sql",
  keyvault = "keyvault",
  sso = "sso",
  ApiConnector = "api-connection",
  cicd = "cicd",
  M365SearchApp = "M365SearchApp",
  M365SsoLaunchPage = "M365SsoLaunchPage",
}

export const FeatureIdToComponent = {
  [FeatureId.Tab]: ComponentNames.TeamsTab,
  [FeatureId.TabNonSso]: ComponentNames.TeamsTab,
  [FeatureId.M365SsoLaunchPage]: ComponentNames.TeamsTab,
  [FeatureId.Notification]: ComponentNames.TeamsBot,
  [FeatureId.CommandAndResponse]: ComponentNames.TeamsBot,
  [FeatureId.Bot]: ComponentNames.TeamsBot,
  [FeatureId.M365SearchApp]: ComponentNames.TeamsBot,
  [FeatureId.MessagingExtension]: ComponentNames.TeamsBot,
  [FeatureId.function]: ComponentNames.TeamsApi,
  [FeatureId.apim]: ComponentNames.APIMFeature,
  [FeatureId.sql]: ComponentNames.AzureSQL,
  [FeatureId.keyvault]: ComponentNames.KeyVault,
  [FeatureId.sso]: ComponentNames.SSO,
  [FeatureId.ApiConnector]: ComponentNames.ApiConnector,
  [FeatureId.cicd]: ComponentNames.CICD,
};

export function getActionNameByFeatureId(featureId: FeatureId): string | undefined {
  const component = FeatureIdToComponent[featureId];
  if (component) {
    return `${component}.add`;
  }
}

export async function getQuestionsForAddFeatureSubCommand(
  featureId: FeatureId,
  inputs: Inputs
): Promise<Result<QTreeNode | undefined, FxError>> {
  if (BotFeatureIds.includes(featureId)) {
    return await getNotificationTriggerQuestionNode(inputs);
  } else if (TabFeatureIds.includes(featureId)) {
  } else if (featureId === AzureResourceSQLNewUI.id) {
  } else if (featureId === AzureResourceFunctionNewUI.id) {
    functionNameQuestion.validation = undefined;
    return ok(new QTreeNode(functionNameQuestion));
  } else if (featureId === AzureResourceApimNewUI.id) {
  } else if (featureId === AzureResourceKeyVaultNewUI.id) {
  } else if (featureId === CicdOptionItem.id) {
    return await addCicdQuestion(createContextV3(), inputs as InputsWithProjectPath);
  } else if (featureId === ApiConnectionOptionItem.id) {
    const apiConnectorImpl = new ApiConnectorImpl();
    return apiConnectorImpl.generateQuestion(createContextV3(), inputs);
  } else if (featureId === SingleSignOnOptionItem.id) {
  }
  return ok(undefined);
}

export async function getNotificationTriggerQuestionNode(
  inputs: Inputs
): Promise<Result<QTreeNode | undefined, FxError>> {
  const res = new QTreeNode({
    type: "group",
  });
  if (isCLIDotNetEnabled()) {
    Object.values(Runtime).forEach((runtime) => {
      const node = new QTreeNode(createHostTypeTriggerQuestion(inputs.platform, runtime));
      node.condition = getConditionOfNotificationTriggerQuestion(runtime);
      res.addChild(node);
    });
  } else {
    const runtime = getPlatformRuntime(inputs.platform);
    const node = new QTreeNode(createHostTypeTriggerQuestion(inputs.platform, runtime));
    res.addChild(node);
  }
  res.condition = showNotificationTriggerCondition;
  return ok(res);
}
