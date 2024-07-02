// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, OptionItem, Platform } from "@microsoft/teamsfx-api";
import { FeatureFlags, featureFlagManager } from "../common/featureFlags";
import { getLocalizedString } from "../common/localizeUtils";
import { OfficeAddinProjectConfig } from "../component/generator/officeXMLAddin/projectConfig";

export enum QuestionNames {
  Scratch = "scratch",
  SctatchYes = "scratch-yes",
  AppName = "app-name",
  Folder = "folder",
  ProjectPath = "projectPath",
  ProgrammingLanguage = "programming-language",
  ProjectType = "project-type",
  Capabilities = "capabilities",
  BotTrigger = "bot-host-type-trigger",
  Runtime = "runtime",
  SPFxSolution = "spfx-solution",
  SPFxInstallPackage = "spfx-install-latest-package",
  SPFxFramework = "spfx-framework-type",
  SPFxWebpartName = "spfx-webpart-name",
  SPFxWebpartDesc = "spfx-webpart-desp",
  SPFxFolder = "spfx-folder",
  OfficeAddinFolder = "addin-project-folder",
  OfficeAddinManifest = "addin-project-manifest",
  OfficeAddinTemplate = "addin-template-select",
  OfficeAddinHost = "addin-host",
  OfficeAddinImport = "addin-import",
  OfficeAddinFramework = "office-addin-framework-type",
  Samples = "samples",
  ReplaceContentUrl = "replaceContentUrl",
  ReplaceWebsiteUrl = "replaceWebsiteUrl",
  ReplaceBotIds = "replaceBotIds",
  SafeProjectName = "safeProjectName",
  RepalceTabUrl = "tdp-tab-url",
  ValidateMethod = "validate-method",
  AppPackagePath = "appPackagePath",
  FromExistingApi = "from-existing-api", // group name for creating an App from existing api
  ApiSpecLocation = "openapi-spec-location",
  ApiOperation = "api-operation",
  MeArchitectureType = "me-architecture",
  ApiSpecApiKey = "api-key",
  ApiSpecApiKeyConfirm = "api-key-confirm",
  ApiAuth = "api-auth",
  OauthClientSecret = "oauth-client-secret",
  OauthClientId = "oauth-client-id",
  OauthConfirm = "oauth-confirm",

  CustomCopilotRag = "custom-copilot-rag",
  CustomCopilotAssistant = "custom-copilot-agent",
  LLMService = "llm-service",
  OpenAIKey = "openai-key",
  AzureOpenAIKey = "azure-openai-key",
  AzureOpenAIEndpoint = "azure-openai-endpoint",
  AzureOpenAIDeploymentName = "azure-openai-deployment-name",

  Features = "features",
  Env = "env",
  SourceEnvName = "sourceEnvName",
  TargetEnvName = "targetEnvName",
  TargetResourceGroupName = "targetResourceGroupName",
  NewResourceGroupName = "newResourceGroupName",
  NewResourceGroupLocation = "newResourceGroupLocation",
  NewTargetEnvName = "newTargetEnvName",
  ExistingTabEndpoint = "existing-tab-endpoint",
  TeamsAppManifestFilePath = "manifest-path",
  LocalTeamsAppManifestFilePath = "local-manifest-path",
  AadAppManifestFilePath = "manifest-file-path",
  TeamsAppPackageFilePath = "app-package-file-path",
  ConfirmManifest = "confirmManifest",
  ConfirmLocalManifest = "confirmLocalManifest",
  ConfirmAadManifest = "confirmAadManifest",
  OutputZipPathParamName = "output-zip-path",
  OutputManifestParamName = "output-manifest-path",
  M365Host = "m365-host",

  ManifestPath = "manifest-path",
  ManifestId = "manifest-id",
  TitleId = "title-id",
  UserEmail = "email",

  UninstallMode = "mode",
  UninstallModeManifestId = "manifest-id",
  UninstallModeEnv = "env",
  UninstallModeTitleId = "title-id",
  UninstallOptions = "options",
  UninstallOptionM365 = "m365-app",
  UninstallOptionTDP = "app-registration",
  UninstallOptionBot = "bot-framework-registration",

  collaborationAppType = "collaborationType",
  DestinationApiSpecFilePath = "destination-api-spec-location",
  PluginAvailability = "plugin-availability",
}

export const AppNamePattern =
  '^(?=(.*[\\da-zA-Z]){2})[a-zA-Z][^"<>:\\?/*&|\u0000-\u001F]*[^"\\s.<>:\\?/*&|\u0000-\u001F]$';

export enum CliQuestionName {
  Capability = "capability",
}

export enum ProgrammingLanguage {
  JS = "javascript",
  TS = "typescript",
  CSharp = "csharp",
  PY = "python",
  Common = "common",
  None = "none",
}

export const copilotPluginApiSpecOptionId = "copilot-plugin-existing-api";
export const copilotPluginExistingApiOptionIds = [copilotPluginApiSpecOptionId];
export const copilotPluginNewApiOptionId = "copilot-plugin-new-api";
export const copilotPluginOptionIds = [copilotPluginNewApiOptionId, copilotPluginApiSpecOptionId];
export const capabilitiesHavePythonOption = [
  "custom-copilot-basic",
  "custom-copilot-rag-azureAISearch",
  "custom-copilot-rag-customize",
  "custom-copilot-agent-new",
  "custom-copilot-agent-assistants-api",
];

export class RuntimeOptions {
  static NodeJS(): OptionItem {
    return {
      id: "node",
      label: "Node.js",
      detail: getLocalizedString("core.RuntimeOptionNodeJS.detail"),
    };
  }
  static DotNet(): OptionItem {
    return {
      id: "dotnet",
      label: ".NET Core",
      detail: getLocalizedString("core.RuntimeOptionDotNet.detail"),
    };
  }
}

export function getRuntime(inputs: Inputs): string {
  let runtime = RuntimeOptions.NodeJS().id;
  if (featureFlagManager.getBooleanValue(FeatureFlags.CLIDotNet)) {
    runtime = inputs[QuestionNames.Runtime] || runtime;
  } else {
    if (inputs?.platform === Platform.VS) {
      runtime = RuntimeOptions.DotNet().id;
    }
  }
  return runtime;
}

export class ScratchOptions {
  static yes(): OptionItem {
    return {
      id: "yes",
      label: getLocalizedString("core.ScratchOptionYes.label"),
      detail: getLocalizedString("core.ScratchOptionYes.detail"),
    };
  }
  static no(): OptionItem {
    return {
      id: "no",
      label: getLocalizedString("core.ScratchOptionNo.label"),
      detail: getLocalizedString("core.ScratchOptionNo.detail"),
    };
  }
  static all(): OptionItem[] {
    return [ScratchOptions.yes(), ScratchOptions.no()];
  }
}

export class ProjectTypeOptions {
  static getCreateGroupName(): string | undefined {
    return featureFlagManager.getBooleanValue(FeatureFlags.ChatParticipant)
      ? getLocalizedString("core.createProjectQuestion.projectType.createGroup.title")
      : undefined;
  }
  static tab(platform?: Platform): OptionItem {
    return {
      id: "tab-type",
      label: `${platform === Platform.VSCode ? "$(browser) " : ""}${getLocalizedString(
        "core.TabOption.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.tab.detail"),
      groupName: ProjectTypeOptions.getCreateGroupName(),
    };
  }

  static bot(platform?: Platform): OptionItem {
    return {
      id: "bot-type",
      label: `${platform === Platform.VSCode ? "$(hubot) " : ""}${getLocalizedString(
        "core.createProjectQuestion.projectType.bot.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.bot.detail"),
      groupName: ProjectTypeOptions.getCreateGroupName(),
    };
  }

  static me(platform?: Platform): OptionItem {
    return {
      id: "me-type",
      label: `${platform === Platform.VSCode ? "$(symbol-keyword) " : ""}${getLocalizedString(
        "core.MessageExtensionOption.label"
      )}`,
      detail: featureFlagManager.getBooleanValue(FeatureFlags.CopilotPlugin)
        ? getLocalizedString(
            "core.createProjectQuestion.projectType.messageExtension.copilotEnabled.detail"
          )
        : getLocalizedString("core.createProjectQuestion.projectType.messageExtension.detail"),
      groupName: ProjectTypeOptions.getCreateGroupName(),
    };
  }

  static outlookAddin(platform?: Platform): OptionItem {
    return {
      id: "outlook-addin-type",
      label: `${platform === Platform.VSCode ? "$(mail) " : ""}${getLocalizedString(
        "core.createProjectQuestion.projectType.outlookAddin.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.outlookAddin.detail"),
      groupName: ProjectTypeOptions.getCreateGroupName(),
    };
  }

  static officeAddin(platform?: Platform): OptionItem {
    return {
      id: "office-addin-type",
      label: `${platform === Platform.VSCode ? "$(extensions) " : ""}${getLocalizedString(
        "core.createProjectQuestion.projectType.officeAddin.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.officeAddin.detail"),
      groupName: ProjectTypeOptions.getCreateGroupName(),
    };
  }

  static officeAddinAllIds(platform?: Platform): string[] {
    return [
      ProjectTypeOptions.officeAddin(platform).id,
      ProjectTypeOptions.outlookAddin(platform).id,
    ];
  }

  static copilotPlugin(platform?: Platform): OptionItem {
    return {
      id: "copilot-plugin-type",
      label: `${
        platform === Platform.VSCode ? "$(teamsfx-copilot-plugin) " : ""
      }${getLocalizedString("core.createProjectQuestion.projectType.copilotPlugin.label")}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.copilotPlugin.detail"),
      groupName: ProjectTypeOptions.getCreateGroupName(),
    };
  }

  static customCopilot(platform?: Platform): OptionItem {
    return {
      id: "custom-copilot-type",
      label: `${
        platform === Platform.VSCode ? "$(teamsfx-custom-copilot) " : ""
      }${getLocalizedString("core.createProjectQuestion.projectType.customCopilot.label")}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.customCopilot.detail"),
      groupName: ProjectTypeOptions.getCreateGroupName(),
    };
  }

  static startWithGithubCopilot(): OptionItem {
    return {
      id: "start-with-github-copilot",
      label: `$(comment-discussion) ${getLocalizedString(
        "core.createProjectQuestion.projectType.copilotHelp.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.copilotHelp.detail"),
      groupName: getLocalizedString("core.createProjectQuestion.projectType.copilotGroup.title"),
    };
  }

  static customizeGpt(): OptionItem {
    return {
      id: "customize-gpt-type",
      label: getLocalizedString("core.createProjectQuestion.projectType.declarativeCopilot.label"),
      detail: getLocalizedString("core.createProjectQuestion.projectType.declarativeCopilot.title"),
      groupName: ProjectTypeOptions.getCreateGroupName(),
    };
  }
}

export class CapabilityOptions {
  // empty
  static empty(): OptionItem {
    return {
      id: "empty",
      label: "Empty",
    };
  }

  // bot
  static basicBot(): OptionItem {
    return {
      id: "bot",
      label: `${getLocalizedString("core.BotNewUIOption.label")}`,
      detail: getLocalizedString("core.BotNewUIOption.detail"),
    };
  }
  static notificationBot(): OptionItem {
    return {
      // For default option, id and cliName must be the same
      id: "notification",
      label: `${getLocalizedString("core.NotificationOption.label")}`,
      detail: getLocalizedString("core.NotificationOption.detail"),
      data: "https://aka.ms/teamsfx-send-notification",
      buttons: [
        {
          iconPath: "file-symlink-file",
          tooltip: getLocalizedString("core.option.github"),
          command: "fx-extension.openTutorial",
        },
      ],
    };
  }

  static commandBot(): OptionItem {
    return {
      // id must match cli `yargsHelp`
      id: "command-bot",
      label: `${getLocalizedString("core.CommandAndResponseOption.label")}`,
      detail: getLocalizedString("core.CommandAndResponseOption.detail"),
      data: "https://aka.ms/teamsfx-create-command",
      buttons: [
        {
          iconPath: "file-symlink-file",
          tooltip: getLocalizedString("core.option.github"),
          command: "fx-extension.openTutorial",
        },
      ],
    };
  }

  static workflowBot(inputs?: Inputs): OptionItem {
    const item: OptionItem = {
      // id must match cli `yargsHelp`
      id: "workflow-bot",
      label: `${getLocalizedString("core.WorkflowOption.label")}`,
      detail: getLocalizedString("core.WorkflowOption.detail"),
      data: "https://aka.ms/teamsfx-create-workflow",
      buttons: [
        {
          iconPath: "file-symlink-file",
          tooltip: getLocalizedString("core.option.github"),
          command: "fx-extension.openTutorial",
        },
      ],
    };
    if (inputs?.inProductDoc) {
      item.data = "cardActionResponse";
      item.buttons = [
        {
          iconPath: "file-code",
          tooltip: getLocalizedString("core.option.inProduct"),
          command: "fx-extension.openTutorial",
        },
      ];
    }
    return item;
  }

  //tab

  static nonSsoTab(): OptionItem {
    return {
      id: "tab-non-sso",
      label: `${getLocalizedString("core.TabNonSso.label")}`,
      detail: getLocalizedString("core.TabNonSso.detail"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlookM365"
      ),
    };
  }

  static tab(): OptionItem {
    return {
      id: "tab",
      label: getLocalizedString("core.TabOption.label"),
      description: getLocalizedString("core.TabOption.description"),
      detail: getLocalizedString("core.TabOption.detail"),
    };
  }

  static m365SsoLaunchPage(): OptionItem {
    return {
      id: "sso-launch-page",
      label: `${getLocalizedString("core.M365SsoLaunchPageOptionItem.label")}`,
      detail: getLocalizedString("core.M365SsoLaunchPageOptionItem.detail"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlookM365"
      ),
    };
  }

  static dashboardTab(): OptionItem {
    return {
      id: "dashboard-tab",
      label: `${getLocalizedString("core.DashboardOption.label")}`,
      detail: getLocalizedString("core.DashboardOption.detail"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlookM365"
      ),
      data: "https://aka.ms/teamsfx-dashboard-app",
      buttons: [
        {
          iconPath: "file-symlink-file",
          tooltip: getLocalizedString("core.option.github"),
          command: "fx-extension.openTutorial",
        },
      ],
    };
  }

  static SPFxTab(): OptionItem {
    return {
      id: "tab-spfx",
      label: getLocalizedString("core.TabSPFxOption.labelNew"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlookM365"
      ),
      detail: getLocalizedString("core.TabSPFxOption.detailNew"),
    };
  }

  //message extension
  static linkUnfurling(): OptionItem {
    return {
      id: "link-unfurling",
      label: `${getLocalizedString("core.LinkUnfurlingOption.label")}`,
      detail: getLocalizedString("core.LinkUnfurlingOption.detail"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlook"
      ),
    };
  }

  static m365SearchMe(): OptionItem {
    return {
      id: "search-app",
      label: `${getLocalizedString("core.M365SearchAppOptionItem.label")}`,
      detail: featureFlagManager.getBooleanValue(FeatureFlags.CopilotPlugin)
        ? getLocalizedString("core.M365SearchAppOptionItem.copilot.detail")
        : getLocalizedString("core.M365SearchAppOptionItem.detail"),
    };
  }

  static SearchMe(): OptionItem {
    return {
      id: "search-message-extension",
      label: `${getLocalizedString("core.M365SearchAppOptionItem.label")}`,
      detail: getLocalizedString("core.SearchAppOptionItem.detail"),
    };
  }

  static collectFormMe(): OptionItem {
    return {
      id: "collect-form-message-extension",
      label: `${getLocalizedString("core.MessageExtensionOption.labelNew")}`,
      detail: getLocalizedString("core.MessageExtensionOption.detail"),
    };
  }
  static me(): OptionItem {
    return {
      id: "message-extension",
      label: getLocalizedString("core.MessageExtensionOption.label"),
      description: getLocalizedString("core.MessageExtensionOption.description"),
      detail: getLocalizedString("core.MessageExtensionOption.detail"),
    };
  }
  static bots(inputs?: Inputs): OptionItem[] {
    if (inputs?.platform === Platform.VS) {
      return [
        CapabilityOptions.basicBot(),
        CapabilityOptions.aiBot(),
        CapabilityOptions.aiAssistantBot(),
        CapabilityOptions.notificationBot(),
        CapabilityOptions.commandBot(),
        CapabilityOptions.workflowBot(inputs),
      ];
    }
    return [
      CapabilityOptions.basicBot(),
      CapabilityOptions.notificationBot(),
      CapabilityOptions.commandBot(),
      CapabilityOptions.workflowBot(inputs),
    ];
  }

  static tabs(): OptionItem[] {
    return [
      CapabilityOptions.nonSsoTab(),
      CapabilityOptions.m365SsoLaunchPage(),
      CapabilityOptions.dashboardTab(),
      CapabilityOptions.SPFxTab(),
    ];
  }

  static dotnetCaps(inputs?: Inputs): OptionItem[] {
    const capabilities = [
      CapabilityOptions.empty(),
      ...CapabilityOptions.copilotPlugins(),
      ...CapabilityOptions.bots(inputs),
      CapabilityOptions.nonSsoTab(),
      CapabilityOptions.tab(),
      ...CapabilityOptions.collectMECaps(),
    ];
    if (featureFlagManager.getBooleanValue(FeatureFlags.TdpTemplateCliTest)) {
      capabilities.push(CapabilityOptions.me());
    }

    return capabilities;
  }

  /**
   * Collect all capabilities for message extension, including dotnet and nodejs.
   * @returns OptionItem[] capability list
   */
  static collectMECaps(): OptionItem[] {
    return [
      CapabilityOptions.m365SearchMe(),
      CapabilityOptions.collectFormMe(),
      CapabilityOptions.SearchMe(),
      CapabilityOptions.linkUnfurling(),
    ];
  }

  static mes(inputs?: Inputs): OptionItem[] {
    return inputs !== undefined && getRuntime(inputs) === RuntimeOptions.DotNet().id
      ? [
          CapabilityOptions.SearchMe(),
          CapabilityOptions.collectFormMe(),
          CapabilityOptions.linkUnfurling(),
        ]
      : [
          CapabilityOptions.m365SearchMe(),
          CapabilityOptions.collectFormMe(),
          CapabilityOptions.linkUnfurling(),
        ];
  }

  static officeAddinStaticCapabilities(host?: string): OptionItem[] {
    const items: OptionItem[] = [];
    for (const h of Object.keys(OfficeAddinProjectConfig)) {
      if (host && h !== host) continue;
      const hostValue = OfficeAddinProjectConfig[h];
      for (const capability of Object.keys(hostValue)) {
        const capabilityValue = hostValue[capability];
        items.push({
          id: capability,
          label: getLocalizedString(capabilityValue.title),
          detail: getLocalizedString(capabilityValue.detail),
        });
      }
    }
    return items;
  }

  static officeAddinDynamicCapabilities(projectType: string, host?: string): OptionItem[] {
    const items: OptionItem[] = [];
    const isOutlookAddin = projectType === ProjectTypeOptions.outlookAddin().id;
    const isOfficeAddin = projectType === ProjectTypeOptions.officeAddin().id;

    const pushToItems = (option: any) => {
      const capabilityValue = OfficeAddinProjectConfig.json[option];
      items.push({
        id: option,
        label: getLocalizedString(capabilityValue.title),
        detail: getLocalizedString(capabilityValue.detail),
      });
    };

    if (isOutlookAddin || isOfficeAddin) {
      pushToItems("json-taskpane");
      if (isOutlookAddin) {
        items.push(CapabilityOptions.outlookAddinImport());
      } else {
        items.push(CapabilityOptions.officeContentAddin());
        items.push(CapabilityOptions.officeAddinImport());
      }
    } else {
      if (host) {
        const hostValue = OfficeAddinProjectConfig[host];
        for (const capability of Object.keys(hostValue)) {
          const capabilityValue = hostValue[capability];
          items.push({
            id: capability,
            label: getLocalizedString(capabilityValue.title),
            detail: getLocalizedString(capabilityValue.detail),
          });
        }
      }
    }
    return items;
  }

  static copilotPlugins(): OptionItem[] {
    return [CapabilityOptions.copilotPluginNewApi(), CapabilityOptions.copilotPluginApiSpec()];
  }

  static customCopilots(): OptionItem[] {
    return [
      CapabilityOptions.customCopilotBasic(),
      CapabilityOptions.customCopilotRag(),
      CapabilityOptions.customCopilotAssistant(),
    ];
  }

  static tdpIntegrationCapabilities(): OptionItem[] {
    // templates that are used by TDP integration only
    return [
      CapabilityOptions.me(),
      CapabilityOptions.botAndMe(),
      CapabilityOptions.nonSsoTabAndBot(),
    ];
  }

  static customizeGptOptions(): OptionItem[] {
    return [CapabilityOptions.customizeGptBasic(), CapabilityOptions.customizeGptWithPlugin()];
  }

  /**
   * static capability list, which does not depend on any feature flags
   */
  static staticAll(inputs?: Inputs): OptionItem[] {
    const capabilityOptions = [
      CapabilityOptions.empty(),
      ...CapabilityOptions.bots(inputs),
      ...CapabilityOptions.tabs(),
      ...CapabilityOptions.collectMECaps(),
      ...CapabilityOptions.copilotPlugins(),
      ...CapabilityOptions.customCopilots(),
      ...CapabilityOptions.tdpIntegrationCapabilities(),
      ...CapabilityOptions.customizeGptOptions(),
    ];
    capabilityOptions.push(...CapabilityOptions.officeAddinStaticCapabilities());
    return capabilityOptions;
  }

  /**
   * dynamic capability list, which depends on feature flags
   */
  static all(inputs?: Inputs): OptionItem[] {
    const capabilityOptions = [
      ...CapabilityOptions.bots(inputs),
      ...CapabilityOptions.tabs(),
      ...CapabilityOptions.collectMECaps(),
    ];
    if (featureFlagManager.getBooleanValue(FeatureFlags.CopilotPlugin)) {
      capabilityOptions.push(...CapabilityOptions.copilotPlugins());
    }
    if (featureFlagManager.getBooleanValue(FeatureFlags.CustomizeGpt)) {
      capabilityOptions.push(...CapabilityOptions.customizeGptOptions());
    }
    capabilityOptions.push(...CapabilityOptions.customCopilots());
    if (featureFlagManager.getBooleanValue(FeatureFlags.TdpTemplateCliTest)) {
      // test templates that are used by TDP integration only
      capabilityOptions.push(...CapabilityOptions.tdpIntegrationCapabilities());
    }
    capabilityOptions.push(
      ...CapabilityOptions.officeAddinDynamicCapabilities(inputs?.projectType, inputs?.host)
    );
    return capabilityOptions;
  }

  static outlookAddinImport(): OptionItem {
    return {
      id: "outlook-addin-import",
      label: getLocalizedString("core.importAddin.label"),
      detail: getLocalizedString("core.importAddin.detail"),
    };
  }

  static officeAddinImport(): OptionItem {
    return {
      id: "office-addin-import",
      label: getLocalizedString("core.importOfficeAddin.label"),
      detail: getLocalizedString("core.importAddin.detail"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.previewOnWindow"
      ),
    };
  }

  static officeContentAddin(): OptionItem {
    return {
      id: "office-content-addin",
      label: getLocalizedString("core.officeContentAddin.label"),
      detail: getLocalizedString("core.officeContentAddin.detail"),
    };
  }

  static nonSsoTabAndBot(): OptionItem {
    return {
      id: "TabNonSsoAndBot",
      label: "", // No need to set display name as this option won't be shown in UI
    };
  }

  static botAndMe(): OptionItem {
    return {
      id: "BotAndMessageExtension",
      label: "", // No need to set display name as this option won't be shown in UI
    };
  }

  // copilot plugin
  static copilotPluginNewApi(): OptionItem {
    return {
      id: copilotPluginNewApiOptionId,
      label: getLocalizedString(
        "core.createProjectQuestion.capability.copilotPluginNewApiOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.copilotPluginNewApiOption.detail"
      ),
    };
  }

  static copilotPluginApiSpec(): OptionItem {
    return {
      id: copilotPluginApiSpecOptionId,
      label: getLocalizedString(
        "core.createProjectQuestion.capability.copilotPluginApiSpecOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.copilotPluginApiSpecOption.detail"
      ),
    };
  }

  static aiBot(): OptionItem {
    return {
      id: "ai-bot",
      label: getLocalizedString("core.aiBotOption.label"),
      detail: getLocalizedString("core.aiBotOption.detail"),
    };
  }

  static aiAssistantBot(): OptionItem {
    return {
      id: "ai-assistant-bot",
      label: getLocalizedString("core.aiAssistantBotOption.label"),
      detail: getLocalizedString("core.aiAssistantBotOption.detail"),
      description: getLocalizedString("core.createProjectQuestion.option.description.preview"),
    };
  }

  // custom copilot
  static customCopilotBasic(): OptionItem {
    return {
      id: "custom-copilot-basic",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotBasicOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotBasicOption.detail"
      ),
    };
  }

  static customCopilotRag(): OptionItem {
    return {
      id: "custom-copilot-rag",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagOption.detail"
      ),
    };
  }

  static customCopilotAssistant(): OptionItem {
    return {
      id: "custom-copilot-agent",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotAssistantOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotAssistantOption.detail"
      ),
    };
  }

  // customize GPT
  static customizeGptBasic(): OptionItem {
    return {
      id: "basic-declarative-copilot",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.declarativeCopilotBasic.title"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.declarativeCopilotBasic.detail"
      ),
    };
  }

  static customizeGptWithPlugin(): OptionItem {
    return {
      id: "declarative-copilot-with-plugin-from-scratch",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.declarativeCopilotWithPlugin.title"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.declarativeCopilotWithPlugin.detail"
      ),
    };
  }
}

export class ApiAuthOptions {
  static none(): OptionItem {
    return {
      id: "none",
      label: "None",
    };
  }
  static apiKey(): OptionItem {
    return {
      id: "api-key",
      label: "API Key",
    };
  }

  static microsoftEntra(): OptionItem {
    return {
      id: "microsoft-entra",
      label: "Microsoft Entra",
    };
  }

  static oauth(): OptionItem {
    return {
      id: "oauth",
      label: "OAuth",
    };
  }

  static all(): OptionItem[] {
    return [
      ApiAuthOptions.none(),
      ApiAuthOptions.apiKey(),
      ApiAuthOptions.microsoftEntra(),
      ApiAuthOptions.oauth(),
    ];
  }
}

export class MeArchitectureOptions {
  static botMe(): OptionItem {
    return {
      id: "bot",
      label: getLocalizedString("core.createProjectQuestion.capability.botMessageExtension.label"),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.botMessageExtension.detail"
      ),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlook"
      ),
    };
  }

  static botPlugin(): OptionItem {
    return {
      id: "bot-plugin",
      label: getLocalizedString("core.createProjectQuestion.capability.botMessageExtension.label"),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.botMessageExtension.detail"
      ),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlookCopilot"
      ),
    };
  }

  static newApi(): OptionItem {
    return {
      id: "new-api",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.copilotPluginNewApiOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.messageExtensionNewApiOption.detail"
      ),
    };
  }

  static apiSpec(): OptionItem {
    return {
      id: "api-spec",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.copilotPluginApiSpecOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.messageExtensionApiSpecOption.detail"
      ),
    };
  }

  static all(): OptionItem[] {
    return [
      MeArchitectureOptions.newApi(),
      MeArchitectureOptions.apiSpec(),
      featureFlagManager.getBooleanValue(FeatureFlags.CopilotPlugin)
        ? MeArchitectureOptions.botPlugin()
        : MeArchitectureOptions.botMe(),
    ];
  }

  static staticAll(): OptionItem[] {
    return [
      MeArchitectureOptions.newApi(),
      MeArchitectureOptions.apiSpec(),
      MeArchitectureOptions.botPlugin(),
      MeArchitectureOptions.botMe(),
    ];
  }
}

export enum HostType {
  AppService = "app-service",
  Functions = "azure-functions",
}

export const NotificationTriggers = {
  HTTP: "http",
  TIMER: "timer",
} as const;

type NotificationTrigger = typeof NotificationTriggers[keyof typeof NotificationTriggers];

interface HostTypeTriggerOptionItem extends OptionItem {
  hostType: HostType;
  triggers?: NotificationTrigger[];
}

export class NotificationTriggerOptions {
  static appService(): HostTypeTriggerOptionItem {
    return {
      id: "http-restify",
      hostType: HostType.AppService,
      label: getLocalizedString("plugins.bot.triggers.http-restify.label"),
      description: getLocalizedString("plugins.bot.triggers.http-restify.description"),
      detail: getLocalizedString("plugins.bot.triggers.http-restify.detail"),
    };
  }
  static appServiceForVS(): HostTypeTriggerOptionItem {
    return {
      id: "http-webapi",
      hostType: HostType.AppService,
      label: getLocalizedString("plugins.bot.triggers.http-webapi.label"),
      description: getLocalizedString("plugins.bot.triggers.http-webapi.description"),
      detail: getLocalizedString("plugins.bot.triggers.http-webapi.detail"),
    };
  }
  // NOTE: id must be the sample as cliName to prevent parsing error for CLI default value.
  static functionsTimerTrigger(): HostTypeTriggerOptionItem {
    return {
      id: "timer-functions",
      hostType: HostType.Functions,
      triggers: [NotificationTriggers.TIMER],
      label: getLocalizedString("plugins.bot.triggers.timer-functions.label"),
      description: getLocalizedString("plugins.bot.triggers.timer-functions.description"),
      detail: getLocalizedString("plugins.bot.triggers.timer-functions.detail"),
    };
  }

  static functionsTimerTriggerIsolated(): HostTypeTriggerOptionItem {
    return {
      id: "timer-functions-isolated",
      hostType: HostType.Functions,
      triggers: [NotificationTriggers.TIMER],
      label: getLocalizedString("plugins.bot.triggers.timer-functions.label"),
      description: getLocalizedString("plugins.bot.triggers.timer-functions.description"),
      detail: getLocalizedString("plugins.bot.triggers.timer-functions.detail"),
    };
  }

  static functionsHttpAndTimerTrigger(): HostTypeTriggerOptionItem {
    return {
      id: "http-and-timer-functions",
      hostType: HostType.Functions,
      triggers: [NotificationTriggers.HTTP, NotificationTriggers.TIMER],
      label: getLocalizedString("plugins.bot.triggers.http-and-timer-functions.label"),
      description: getLocalizedString("plugins.bot.triggers.http-and-timer-functions.description"),
      detail: getLocalizedString("plugins.bot.triggers.http-and-timer-functions.detail"),
    };
  }

  static functionsHttpAndTimerTriggerIsolated(): HostTypeTriggerOptionItem {
    return {
      id: "http-and-timer-functions-isolated",
      hostType: HostType.Functions,
      triggers: [NotificationTriggers.HTTP, NotificationTriggers.TIMER],
      label: getLocalizedString("plugins.bot.triggers.http-and-timer-functions.label"),
      description: getLocalizedString("plugins.bot.triggers.http-and-timer-functions.description"),
      detail: getLocalizedString("plugins.bot.triggers.http-and-timer-functions.detail"),
    };
  }

  static functionsHttpTrigger(): HostTypeTriggerOptionItem {
    return {
      id: "http-functions",
      hostType: HostType.Functions,
      triggers: [NotificationTriggers.HTTP],
      label: getLocalizedString("plugins.bot.triggers.http-functions.label"),
      description: getLocalizedString("plugins.bot.triggers.http-functions.description"),
      detail: getLocalizedString("plugins.bot.triggers.http-functions.detail"),
    };
  }

  static functionsHttpTriggerIsolated(): HostTypeTriggerOptionItem {
    return {
      id: "http-functions-isolated",
      hostType: HostType.Functions,
      triggers: [NotificationTriggers.HTTP],
      label: getLocalizedString("plugins.bot.triggers.http-functions.label"),
      description: getLocalizedString("plugins.bot.triggers.http-functions.description"),
      detail: getLocalizedString("plugins.bot.triggers.http-functions.detail"),
    };
  }

  static functionsTriggers(): HostTypeTriggerOptionItem[] {
    return [
      NotificationTriggerOptions.functionsHttpAndTimerTrigger(),
      NotificationTriggerOptions.functionsHttpTrigger(),
      NotificationTriggerOptions.functionsTimerTrigger(),
    ];
  }

  static all(): HostTypeTriggerOptionItem[] {
    return [
      NotificationTriggerOptions.appService(),
      NotificationTriggerOptions.appServiceForVS(),
      NotificationTriggerOptions.functionsHttpAndTimerTrigger(),
      NotificationTriggerOptions.functionsHttpTrigger(),
      NotificationTriggerOptions.functionsTimerTrigger(),
    ];
  }
}

export enum SPFxVersionOptionIds {
  installLocally = "true",
  globalPackage = "false",
}

export class CustomCopilotRagOptions {
  static customize(): OptionItem {
    return {
      id: "custom-copilot-rag-customize",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagCustomizeOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagCustomizeOption.detail"
      ),
    };
  }

  static azureAISearch(): OptionItem {
    return {
      id: "custom-copilot-rag-azureAISearch",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagAzureAISearchOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagAzureAISearchOption.detail"
      ),
    };
  }

  static customApi(): OptionItem {
    return {
      id: "custom-copilot-rag-customApi",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagCustomApiOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagCustomApiOption.detail"
      ),
      description: getLocalizedString("core.createProjectQuestion.option.description.preview"),
    };
  }

  static microsoft365(): OptionItem {
    return {
      id: "custom-copilot-rag-microsoft365",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagMicrosoft365Option.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagMicrosoft365Option.detail"
      ),
    };
  }

  static all(): OptionItem[] {
    return [
      CustomCopilotRagOptions.customize(),
      CustomCopilotRagOptions.azureAISearch(),
      CustomCopilotRagOptions.customApi(),
      CustomCopilotRagOptions.microsoft365(),
    ];
  }
}

export class CustomCopilotAssistantOptions {
  static new(): OptionItem {
    return {
      id: "custom-copilot-agent-new",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotAssistantNewOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotAssistantNewOption.detail"
      ),
    };
  }

  static assistantsApi(): OptionItem {
    return {
      id: "custom-copilot-agent-assistants-api",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotAssistantAssistantsApiOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotAssistantAssistantsApiOption.detail"
      ),
      description: getLocalizedString("core.createProjectQuestion.option.description.preview"),
    };
  }

  static all(): OptionItem[] {
    return [CustomCopilotAssistantOptions.new(), CustomCopilotAssistantOptions.assistantsApi()];
  }
}

export const recommendedLocations = [
  "South Africa North",
  "Australia East",
  "Central India",
  "East Asia",
  "Japan East",
  "Korea Central",
  "Southeast Asia",
  "Canada Central",
  "France Central",
  "Germany West Central",
  "Italy North",
  "North Europe",
  "Norway East",
  "Poland Central",
  "Sweden Central",
  "Switzerland North",
  "UK South",
  "West Europe",
  "Israel Central",
  "Qatar Central",
  "UAE North",
  "Brazil South",
  "Central US",
  "East US",
  "East US 2",
  "South Central US",
  "West US 2",
  "West US 3",
];

export class PluginAvailabilityOptions {
  static action(): OptionItem {
    return {
      id: "action",
      label: getLocalizedString("core.pluginAvailability.declarativeCopilot"),
    };
  }
  static copilotPlugin(): OptionItem {
    return {
      id: "copilot-plugin",
      label: getLocalizedString("core.pluginAvailability.copilotForM365"),
    };
  }
  static copilotPluginAndAction(): OptionItem {
    return {
      id: "copilot-plugin-and-action",
      label: getLocalizedString("core.pluginAvailability.declarativeCopilotAndM365"),
    };
  }

  static all(): OptionItem[] {
    return [
      PluginAvailabilityOptions.copilotPlugin(),
      PluginAvailabilityOptions.action(),
      PluginAvailabilityOptions.copilotPluginAndAction(),
    ];
  }
}

export class TeamsAppValidationOptions {
  static schema(): OptionItem {
    return {
      id: "validateAgainstSchema",
      label: getLocalizedString("core.selectValidateMethodQuestion.validate.schemaOption"),
    };
  }
  static package(): OptionItem {
    return {
      id: "validateAgainstPackage",
      label: getLocalizedString("core.selectValidateMethodQuestion.validate.appPackageOption"),
    };
  }
  static testCases(): OptionItem {
    return {
      id: "validateWithTestCases",
      label: getLocalizedString("core.selectValidateMethodQuestion.validate.testCasesOption"),
      description: getLocalizedString(
        "core.selectValidateMethodQuestion.validate.testCasesOptionDescription"
      ),
    };
  }
}

export enum HubTypes {
  teams = "teams",
  outlook = "outlook",
  office = "office",
}

export class HubOptions {
  static teams(): OptionItem {
    return {
      id: "teams",
      label: "Teams",
    };
  }
  static outlook(): OptionItem {
    return {
      id: "outlook",
      label: "Outlook",
    };
  }
  static office(): OptionItem {
    return {
      id: "office",
      label: "the Microsoft 365 app",
    };
  }
  static all(): OptionItem[] {
    return [this.teams(), this.outlook(), this.office()];
  }
}
