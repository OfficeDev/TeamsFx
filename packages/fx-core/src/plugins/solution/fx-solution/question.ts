// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  FuncQuestion,
  Inputs,
  MultiSelectQuestion,
  ok,
  OptionItem,
  TextInputQuestion,
  Void,
} from "@microsoft/teamsfx-api";
import { isBotNotificationEnabled } from "../../../common/featureFlags";
import { getLocalizedString } from "../../../common/localizeUtils";

export const OfficeAddinItem: OptionItem = {
  id: "office-addin",
  label: "office addin label",
  cliName: "tab",
  description: "Office Addin description",
  detail: "Office Addin detail",
};

export const TabOptionItem: OptionItem = {
  id: "Tab",
  label: getLocalizedString("core.TabOption.label"),
  cliName: "tab",
  description: getLocalizedString("core.TabOption.description"),
  detail: getLocalizedString("core.TabOption.detail"),
};

export const TabNewUIOptionItem: OptionItem = {
  id: "Tab",
  label: `$(browser) ${getLocalizedString("core.TabOption.labelNew")}`,
  cliName: "tab",
  detail: getLocalizedString("core.TabOption.detailNew"),
  groupName: getLocalizedString("core.options.separator.scenario"),
  data: "https://aka.ms/teamsfx-tab-with-sso",
  buttons: [
    {
      iconPath: "tasklist",
      tooltip: getLocalizedString("core.option.tutorial"),
      command: "fx-extension.openTutorial",
    },
  ],
};

export const BotOptionItem: OptionItem = {
  id: "Bot",
  label: "Bot",
  cliName: "bot",
  description: getLocalizedString("core.BotOption.description"),
  detail: getLocalizedString("core.BotOption.detail"),
};

export const BotNewUIOptionItem: OptionItem = {
  id: "Bot",
  label: `$(hubot) ${getLocalizedString("core.BotNewUIOption.label")}`,
  cliName: "bot",
  detail: getLocalizedString("core.BotNewUIOption.detail"),
  groupName: getLocalizedString("core.options.separator.basic"),
};

export const NotificationOptionItem: OptionItem = {
  // For default option, id and cliName must be the same
  id: "Notification",
  label: `$(hubot) ${getLocalizedString("core.NotificationOption.label")}`,
  description: getLocalizedString("core.Option.recommend"),
  cliName: "notification",
  detail: getLocalizedString("core.NotificationOption.detail"),
  groupName: getLocalizedString("core.options.separator.scenario"),
  data: "https://aka.ms/teamsfx-send-notification",
  buttons: [
    {
      iconPath: "tasklist",
      tooltip: getLocalizedString("core.option.tutorial"),
      command: "fx-extension.openTutorial",
    },
  ],
};

export const CommandAndResponseOptionItem: OptionItem = {
  // id must match cli `yargsHelp`
  id: "command-bot",
  label: `$(hubot) ${getLocalizedString("core.CommandAndResponseOption.label")}`,
  description: getLocalizedString("core.Option.recommend"),
  cliName: "command-bot",
  detail: getLocalizedString("core.CommandAndResponseOption.detail"),
  groupName: getLocalizedString("core.options.separator.scenario"),
  data: "https://aka.ms/teamsfx-create-command",
  buttons: [
    {
      iconPath: "tasklist",
      tooltip: getLocalizedString("core.option.tutorial"),
      command: "fx-extension.openTutorial",
    },
  ],
};

export const ExistingTabOptionItem: OptionItem = {
  id: "ExistingTab",
  label: `$(browser) ${getLocalizedString("core.ExistingTabOption.label")}`,
  cliName: "existing-tab",
  detail: getLocalizedString("core.ExistingTabOption.detail"),
  groupName: getLocalizedString("core.options.separator.scenario"),
  data: "https://aka.ms/teamsfx-embed-existing-web",
  buttons: [
    {
      iconPath: "tasklist",
      tooltip: getLocalizedString("core.option.tutorial"),
      command: "fx-extension.openTutorial",
    },
  ],
};

export const MessageExtensionItem: OptionItem = {
  id: "MessagingExtension",
  label: getLocalizedString("core.MessageExtensionOption.label"),
  cliName: "message-extension",
  description: getLocalizedString("core.MessageExtensionOption.description"),
  detail: getLocalizedString("core.MessageExtensionOption.detail"),
};

export const MessageExtensionNewUIItem: OptionItem = {
  id: "MessagingExtension",
  label: `$(comment-discussion) ${getLocalizedString("core.MessageExtensionOption.labelNew")}`,
  cliName: "message-extension",
  detail: getLocalizedString("core.MessageExtensionOption.detail"),
  groupName: getLocalizedString("core.options.separator.basic"),
};

export const TabSPFxItem: OptionItem = {
  id: "TabSPFx",
  label: getLocalizedString("core.TabSPFxOption.label"),
  cliName: "tab-spfx",
  description: getLocalizedString("core.TabSPFxOption.description"),
  detail: getLocalizedString("core.TabSPFxOption.detail"),
};

export const TabSPFxNewUIItem: OptionItem = {
  id: "TabSPFx",
  label: `$(browser) ${getLocalizedString("core.TabSPFxOption.labelNew")}`,
  cliName: "tab-spfx",
  detail: getLocalizedString("core.TabSPFxOption.detailNew"),
  groupName: getLocalizedString("core.options.separator.scenario"),
};

export const TabSsoItem: OptionItem = {
  id: "TabSSO",
  label: "TabSSO",
  cliName: "tab-sso",
  description: getLocalizedString("core.TabSso.description"),
  detail: getLocalizedString("core.TabSso.detail"),
  groupName: getLocalizedString("core.options.separator.scenario"),
};

export const BotSsoItem: OptionItem = {
  id: "BotSSO",
  label: "BotSSO",
  cliName: "bot-sso",
  description: getLocalizedString("core.BotSso.description"),
  detail: getLocalizedString("core.BotSso.detail"),
};

export const TabNonSsoItem: OptionItem = {
  id: "TabNonSso",
  label: `$(browser) ${getLocalizedString("core.TabNonSso.label")}`,
  cliName: "tab-non-sso",
  detail: getLocalizedString("core.TabNonSso.detail"),
  groupName: getLocalizedString("core.options.separator.basic"),
};

export const M365SsoLaunchPageOptionItem: OptionItem = {
  id: "M365SsoLaunchPage",
  label: `$(browser) ${getLocalizedString("core.M365SsoLaunchPageOptionItem.label")}`,
  cliName: "sso-launch-page",
  detail: getLocalizedString("core.M365SsoLaunchPageOptionItem.detail"),
  groupName: getLocalizedString("core.options.separator.m365"),
};

export const M365SearchAppOptionItem: OptionItem = {
  id: "M365SearchApp",
  label: `$(comment-discussion) ${getLocalizedString("core.M365SearchAppOptionItem.label")}`,
  cliName: "search-app",
  detail: getLocalizedString("core.M365SearchAppOptionItem.detail"),
  groupName: getLocalizedString("core.options.separator.m365"),
};

export enum AzureSolutionQuestionNames {
  Capabilities = "capabilities",
  TabScopes = "tab-scopes",
  HostType = "host-type",
  AzureResources = "azure-resources",
  PluginSelectionDeploy = "deploy-plugin",
  AddResources = "add-azure-resources",
  AppName = "app-name",
  AskSub = "subscription",
  ProgrammingLanguage = "programming-language",
  Solution = "solution",
  Scenarios = "scenarios",
  Features = "features",
}

export const HostTypeOptionAzure: OptionItem = {
  id: "Azure",
  label: getLocalizedString("core.HostTypeOptionAzure.label"),
  cliName: "azure",
};

export const HostTypeOptionSPFx: OptionItem = {
  id: "SPFx",
  label: getLocalizedString("core.HostTypeOptionSPFx.label"),
  cliName: "spfx",
};

export const AzureResourceSQL: OptionItem = {
  id: "sql",
  label: getLocalizedString("core.AzureResourceSQL.label"),
  description: getLocalizedString("core.AzureResourceSQL.description"),
};

export const AzureResourceSQLNewUI: OptionItem = {
  id: "sql",
  label: `$(azure) ${getLocalizedString("core.AzureResourceSQLNewUI.label")}`,
  detail: getLocalizedString("core.AzureResourceSQLNewUI.detail"),
  groupName: getLocalizedString("core.options.separator.resource"),
};

export const AzureResourceFunction: OptionItem = {
  id: "function",
  label: getLocalizedString("core.AzureResourceFunction.label"),
};

export const AzureResourceFunctionNewUI: OptionItem = {
  id: "function",
  label: `$(azure) ${getLocalizedString("core.AzureResourceFunctionNewUI.label")}`,
  detail: getLocalizedString("core.AzureResourceFunctionNewUI.detail"),
  groupName: getLocalizedString("core.options.separator.resource"),
};

export const AzureResourceApim: OptionItem = {
  id: "apim",
  label: getLocalizedString("core.AzureResourceApim.label"),
  description: getLocalizedString("core.AzureResourceApim.description"),
};

export const AzureResourceApimNewUI: OptionItem = {
  id: "apim",
  label: `$(azure) ${getLocalizedString("core.AzureResourceApimNewUI.label")}`,
  detail: getLocalizedString("core.AzureResourceApimNewUI.detail"),
  groupName: getLocalizedString("core.options.separator.resource"),
};

export const AzureResourceKeyVault: OptionItem = {
  id: "keyvault",
  label: getLocalizedString("core.AzureResourceKeyVault.label"),
  description: getLocalizedString("core.AzureResourceKeyVault.description"),
};

export const AzureResourceKeyVaultNewUI: OptionItem = {
  id: "keyvault",
  label: `$(azure) ${getLocalizedString("core.AzureResourceKeyVaultNewUI.label")}`,
  detail: getLocalizedString("core.AzureResourceKeyVaultNewUI.detail"),
  groupName: getLocalizedString("core.options.separator.resource"),
};

export const SingleSignOnOptionItem: OptionItem = {
  id: "sso",
  label: `$(unlock) ${getLocalizedString("core.SingleSignOnOption.label")}`,
  detail: getLocalizedString("core.SingleSignOnOption.detail"),
  groupName: getLocalizedString("core.options.separator.additional"),
  data: "https://aka.ms/teamsfx-add-sso",
  buttons: [
    {
      iconPath: "tasklist",
      tooltip: getLocalizedString("core.option.tutorial"),
      command: "fx-extension.openTutorial",
    },
  ],
};

export const ApiConnectionOptionItem: OptionItem = {
  id: "api-connection",
  label: `$(arrow-swap) ${getLocalizedString("core.ApiConnectionOption.label")}`,
  detail: getLocalizedString("core.ApiConnectionOption.detail"),
  groupName: getLocalizedString("core.options.separator.additional"),
  data: "https://aka.ms/teamsfx-connect-api",
  buttons: [
    {
      iconPath: "tasklist",
      tooltip: getLocalizedString("core.option.tutorial"),
      command: "fx-extension.openTutorial",
    },
  ],
};

export const CicdOptionItem: OptionItem = {
  id: "cicd",
  label: `$(sync) ${getLocalizedString("core.cicdWorkflowOption.label")}`,
  detail: getLocalizedString("core.cicdWorkflowOption.detail"),
  groupName: getLocalizedString("core.options.separator.additional"),
  data: "https://aka.ms/teamsfx-add-cicd",
  buttons: [
    {
      iconPath: "tasklist",
      tooltip: getLocalizedString("core.option.tutorial"),
      command: "fx-extension.openTutorial",
    },
  ],
};

export enum BotScenario {
  NotificationBot = "notificationBot",
  CommandAndResponseBot = "commandAndResponseBot",
}

export const BotNotificationTriggers = {
  Timer: "timer",
  Http: "http",
} as const;

export type BotNotificationTrigger =
  typeof BotNotificationTriggers[keyof typeof BotNotificationTriggers];

export const AzureResourcesQuestion: MultiSelectQuestion = {
  name: AzureSolutionQuestionNames.AzureResources,
  title: "Cloud resources",
  type: "multiSelect",
  staticOptions: [AzureResourceSQL, AzureResourceFunction],
  default: [],
  onDidChangeSelection: async function (
    currentSelectedIds: Set<string>,
    previousSelectedIds: Set<string>
  ): Promise<Set<string>> {
    if (currentSelectedIds.has(AzureResourceSQL.id)) {
      currentSelectedIds.add(AzureResourceFunction.id);
    }
    return currentSelectedIds;
  },
  placeholder: "Select a resource (optional)",
};

export function createAddAzureResourceQuestion(
  alreadyHaveFunction: boolean,
  alreadyHaveSQL: boolean,
  alreadyHaveAPIM: boolean,
  alreadyHaveKeyVault: boolean
): MultiSelectQuestion {
  const options: OptionItem[] = [AzureResourceFunction, AzureResourceSQL];
  if (!alreadyHaveAPIM) options.push(AzureResourceApim);
  if (!alreadyHaveKeyVault) options.push(AzureResourceKeyVault);
  return {
    name: AzureSolutionQuestionNames.AddResources,
    title: "Cloud resources",
    type: "multiSelect",
    staticOptions: options,
    default: [],
    onDidChangeSelection: async function (
      currentSelectedIds: Set<string>,
      previousSelectedIds: Set<string>
    ): Promise<Set<string>> {
      const hasSQL = currentSelectedIds.has(AzureResourceSQL.id);
      const hasAPIM = currentSelectedIds.has(AzureResourceApim.id);
      if ((hasSQL || hasAPIM) && !alreadyHaveFunction) {
        currentSelectedIds.add(AzureResourceFunction.id);
      }
      return currentSelectedIds;
    },
  };
}

export function createAddCloudResourceOptions(
  alreadyHaveAPIM: boolean,
  alreadyHaveKeyVault: boolean
): OptionItem[] {
  const options: OptionItem[] = [AzureResourceFunctionNewUI];
  if (!alreadyHaveAPIM) options.push(AzureResourceApimNewUI);
  options.push(AzureResourceSQLNewUI);
  if (!alreadyHaveKeyVault) options.push(AzureResourceKeyVaultNewUI);
  return options;
}

export function addCapabilityQuestion(
  alreadyHaveTab: boolean,
  alreadyHaveBot: boolean
): MultiSelectQuestion {
  const options: OptionItem[] = [];
  if (!alreadyHaveTab) options.push(TabOptionItem);
  if (!alreadyHaveBot) {
    options.push(BotOptionItem);
    options.push(MessageExtensionItem);
    options.push(NotificationOptionItem);
    options.push(CommandAndResponseOptionItem);
  }
  return {
    name: AzureSolutionQuestionNames.Capabilities,
    title: isBotNotificationEnabled()
      ? getLocalizedString("core.addCapabilityQuestion.titleNew")
      : getLocalizedString("core.addCapabilityQuestion.title"),
    type: "multiSelect",
    staticOptions: options,
    default: [],
  };
}

export const DeployPluginSelectQuestion: MultiSelectQuestion = {
  name: AzureSolutionQuestionNames.PluginSelectionDeploy,
  title: `Select resources`,
  type: "multiSelect",
  skipSingleOption: true,
  staticOptions: [],
  default: [],
};

export const AskSubscriptionQuestion: FuncQuestion = {
  name: AzureSolutionQuestionNames.AskSub,
  type: "func",
  func: async (inputs: Inputs): Promise<Void> => {
    return ok(Void);
  },
};

export function getUserEmailQuestion(currentUserEmail: string): TextInputQuestion {
  let defaultUserEmail = "";
  if (currentUserEmail && currentUserEmail.indexOf("@") > 0) {
    defaultUserEmail = "[UserName]@" + currentUserEmail.split("@")[1];
  }
  return {
    name: "email",
    type: "text",
    title: getLocalizedString("core.getUserEmailQuestion.title"),
    default: defaultUserEmail,
    validation: {
      validFunc: (input: string, previousInputs?: Inputs): string | undefined => {
        if (!input || input.trim() === "") {
          return getLocalizedString("core.getUserEmailQuestion.validation1");
        }

        input = input.trim();

        if (input === defaultUserEmail) {
          return getLocalizedString("core.getUserEmailQuestion.validation2");
        }

        const re = /\S+@\S+\.\S+/;
        if (!re.test(input)) {
          return getLocalizedString("core.getUserEmailQuestion.validation3");
        }
        return undefined;
      },
    },
  };
}

export const BotFeatureIds = [
  BotOptionItem.id,
  NotificationOptionItem.id,
  CommandAndResponseOptionItem.id,
  MessageExtensionItem.id,
];

export const TabFeatureIds = [TabOptionItem.id, TabNonSsoItem.id];
