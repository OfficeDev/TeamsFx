export const CONFIGURATION_PREFIX = "fx-extension";
export enum ConfigurationKey {
  BicepEnvCheckerEnable = "prerequisiteCheck.bicep",
  AutomaticNpmInstall = "automaticNpmInstall",
  YoEnvCheckerEnable = "SPFxPrerequisiteCheck.yo",
  generatorEnvCheckerEnable = "SPFxPrerequisiteCheck.sharepointYeomanGenerator",
}

export const AzurePortalUrl = "https://portal.azure.com";
export const AzureAssignRoleHelpUrl = "https://aka.ms/teamsfx-azure-role-assignments-help-link";

export const SpfxManageSiteAdminUrl =
  "https://aka.ms/teamsfx-sharepoint-manage-site-admin-help-link";

export enum SyncedState {
  Version = "teamsToolkit:synced:version",
}

export enum UserState {
  IsExisting = "teamsToolkit:user:isExisting",
}

export enum GlobalKey {
  OpenWalkThrough = "fx-extension.openWalkThrough",
  OpenReadMe = "fx-extension.openReadMe",
  OpenSampleReadMe = "fx-extension.openSampleReadMe",
  ShowLocalDebugMessage = "ShowLocalDebugMessage",
  ShowLocalPreviewMessage = "ShowLocalPreviewMessage",
}

export enum AadManifestDeployConstants {
  INCLUDE_AAD_MANIFEST = "include-aad-manifest",
}

export const manifestConfigDataRegex = /{{config.manifest[\.a-zA-Z]+}}/g;
export const manifestStateDataRegex = /{{state\.[a-zA-Z-_]+\.\w+}}/g;

export const SUPPORTED_SPFX_VERSION = "1.15.0";
export const CLI_FOR_M365 =
  "https://pnp.github.io/cli-microsoft365/cmd/spfx/project/project-upgrade/";
