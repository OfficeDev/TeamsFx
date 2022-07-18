export class ConstantString {
  static readonly UTF8Encoding = "utf-8";
  static readonly DeploymentResourceType = "Microsoft.Resources/deployments";
  static readonly DeploymentNotFound = "DeploymentNotFound";
  static readonly RootFolder = "TeamsApps";
}

export class HelpLinks {
  static readonly WhyNeedProvision = "https://aka.ms/teamsfx/whyneedprovision";
  static readonly ArmHelpLink = "https://aka.ms/teamsfx-arm-help";

  // TODO: short link to the docs
  static readonly HowToAddCapability = "https://aka.ms/teamsfx-how-to-add-capability";
}

export class VSCodeExtensionCommand {
  static readonly showOutputChannel = "command:fx-extension.showOutputChannel";
  static readonly openFolder = "command:fx-extension.openFolder";
  static readonly openReadme = "command:fx-extension.openReadMe?%5B%22Notification%22%5D";
}

export class Bicep {
  static readonly ParameterOrchestrationFileName: string = "param.template.bicep";
  static readonly ModuleOrchestrationFileName: string = "module.template.bicep";
  static readonly OutputOrchestrationFileName: string = "output.template.bicep";
  static readonly VariablesOrchestrationFileName: string = "variables.template.bicep";
  static readonly ParameterFileName: string = "parameters.json";
  static readonly ProvisionFileName: string = "provision.template.bicep";
  static readonly ConfigFileName: string = "config.template.bicep";
}

export class TeamsClientId {
  static readonly MobileDesktop = "1fec8e78-bce4-4aaf-ab1b-5451cc387264";
  static readonly Web = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346";
}

export class OfficeClientId {
  static readonly Desktop = "0ec893e0-5785-4de6-99da-4ed124e5296c";
  static readonly Web1 = "4345a7b9-9a63-4910-a426-35363201d503";
  static readonly Web2 = "4765445b-32c6-49b0-83e6-1d93765276ca";
}

export class OutlookClientId {
  static readonly Desktop = "d3590ed6-52b3-4102-aeff-aad2292ab01c";
  static readonly Web1 = "00000002-0000-0ff1-ce00-000000000000";
  static readonly Web2 = "bc59ab01-8403-45c6-8796-ac3ef710b3e3";
}

export class ResourcePlugins {
  static readonly Aad = "fx-resource-aad-app-for-teams";
  static readonly FrontendHosting = "fx-resource-frontend-hosting";
  static readonly SimpleAuth = "fx-resource-simple-auth";
  static readonly Bot = "fx-resource-bot";
  static readonly LocalDebug = "fx-resource-local-debug";
  static readonly AzureSQL = "fx-resource-azure-sql";
  static readonly Function = "fx-resource-function";
  static readonly Identity = "fx-resource-identity";
  static readonly Apim = "fx-resource-apim";
  static readonly SPFx = "fx-resource-spfx";
  static readonly AppStudio = "fx-resource-appstudio";
}
export class PluginDisplayName {
  static readonly Solution = "Teams Toolkit";
  static readonly SpfxSolution = "SPFx";
}

export class FeatureFlagName {
  static readonly BicepEnvCheckerEnable = "TEAMSFX_BICEP_ENV_CHECKER_ENABLE";
  static readonly APIV3 = "TEAMSFX_APIV3";
  // This will default to true and this environment is only for tests. It does not expose to user.
  static readonly InsiderPreview = "__TEAMSFX_INSIDER_PREVIEW";
  static readonly VSCallingCLI = "VS_CALLING_CLI";
  static readonly ExistingTabApp = "TEAMSFX_INIT_APP";
  static readonly AadManifest = "TEAMSFX_AAD_MANIFEST";
  static readonly DebugTemplate = "TEAMSFX_DEBUG_TEMPLATE";
  static readonly YeomanScaffold = "YEOMAN_SCAFFOLD";
  static readonly BotNotification = "BOT_NOTIFICATION_ENABLED";
  static readonly M365App = "TEAMSFX_M365_APP";
  static readonly YoCheckerEnable = "TEAMSFX_YO_ENV_CHECKER_ENABLE";
  static readonly GeneratorCheckerEnable = "TEAMSFX_GENERATOR_ENV_CHECKER_ENABLE";
  static readonly ApiConnect = "TEAMSFX_API_CONNECT_ENABLE";
  static readonly DeployManifest = "TEAMSFX_DEPLOY_MANIFEST";
  static readonly Preview = "TEAMSFX_PREVIEW";
  static readonly CLIDotNet = "TEAMSFX_CLI_DOTNET";
}

export class ManifestVariables {
  static readonly DeveloperWebsiteUrl = "developerWebsiteUrl";
  static readonly DeveloperPrivacyUrl = "developerPrivacyUrl";
  static readonly DeveloperTermsOfUseUrl = "developerTermsOfUseUrl";
  static readonly TabContentUrl = "tabContentUrl";
  static readonly TabWebsiteUrl = "tabWebsiteUrl";
  static readonly TabConfigurationUrl = "tabConfigurationUrl";
  static readonly BotId = "botId";
}
