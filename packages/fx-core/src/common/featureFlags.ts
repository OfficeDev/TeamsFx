// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { FeatureFlagName } from "./constants";

// Determine whether feature flag is enabled based on environment variable setting
export function isFeatureFlagEnabled(featureFlagName: string, defaultValue = false): boolean {
  const flag = process.env[featureFlagName];
  if (flag === undefined) {
    return defaultValue; // allows consumer to set a default value when environment variable not set
  } else {
    return flag === "1" || flag.toLowerCase() === "true"; // can enable feature flag by set environment variable value to "1" or "true"
  }
}

/**
 * Update all preview feature flags.
 */
export function initializePreviewFeatureFlags(): void {
  process.env[FeatureFlagName.BotNotification] = "true";
  process.env[FeatureFlagName.M365App] = "true";
  process.env[FeatureFlagName.AadManifest] = "true";
  process.env[FeatureFlagName.ApiConnect] = "true";
  process.env[FeatureFlagName.DeployManifest] = "true";
  // Force the feature to close until it needs to be released.
  process.env[FeatureFlagName.OfficeAddin] = "false";
}

export function isCLIDotNetEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.CLIDotNet, false);
}

export function isV3Enabled(): boolean {
  return process.env.TEAMSFX_V3 ? process.env.TEAMSFX_V3 === "true" : true;
}

export function isVideoFilterEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.VideoFilter, false);
}

export function isCopilotPluginEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.CopilotPlugin, false);
}

export function isApiCopilotPluginEnabled(): boolean {
  // return isFeatureFlagEnabled(FeatureFlagName.ApiCopilotPlugin, true) && isCopilotPluginEnabled();
  return isFeatureFlagEnabled(FeatureFlagName.ApiCopilotPlugin, false) && isCopilotPluginEnabled();
}

export function enableTestToolByDefault(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.TestTool, true);
}

export function isApiKeyEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.ApiKey, false);
}

export function isMultipleParametersEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.MultipleParameters, true);
}

export function isOfficeXMLAddinEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.OfficeXMLAddin, false);
}

export function isTeamsFxRebrandingEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.TeamsFxRebranding, false);
}

export function isTdpTemplateCliTestEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.TdpTemplateCliTest, false);
}

export function isAsyncAppValidationEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.AsyncAppValidation, false);
}

export function isNewProjectTypeEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.NewProjectType, true);
}

export function isOfficeJSONAddinEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.OfficeAddin, false);
}

export function isApiMeSSOEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.ApiMeSSO, false);
}

///////////////////////////////////////////////////////////////////////////////
// Notes for Office Addin Feature flags:
// Case 1: TEAMSFX_OFFICE_ADDIN = false, TEAMSFX_OFFICE_XML_ADDIN = false
//   1.1 project-type option: `outlook-addin-type`
//   1.2 addin-host: not show but use `outlook` internally
//   1.3 capabilities options: [`jason-taskpane`, `outlook-addin-import`]
//   1.4 programming-language options: [`typescript`] (skip in UI)
//   1.5 office-addin-framework-type: not show question but use `default_old` internally
//   1.6 generator class: OfficeAddinGenerator
//   1.7 template link: config.json.json-taskpane.default_old.typescript
// Case 2: TEAMSFX_OFFICE_ADDIN = false AND TEAMSFX_OFFICE_XML_ADDIN = true
//   2.1 project-type option: `office-xml-addin-type`
//   2.2 addin-host options: [`outlook`, `word`, `excel`, `powerpoint`]
//   2.3 capabilities options:
//     if (addin-host == `outlook`) then [`json-taskpane`, `outlook-addin-import`]
//     else if (addin-host == `word`) then [`word-taskpane`, `word-xxx`, ...]
//     else if (addin-host == `excel`) then [`excel-taskpane`, `excel-xxx`, ...]
//     else if (addin-host === `powerpoint`) then [`powerpoint-taskpane`, `powerpoint-xxx`, ...]
//   2.4 programming-language options:
//     if (addin-host == `outlook`) then [`typescript`] (skip in UI)
//     else two options: [`typescript`, `javascript`]
//   2.5 office-addin-framework-type options: not show but use `default` internally
//   2.6 generator class:
//     if (addin-host == `outlook`) then OfficeAddinGenerator
//     else OfficeXMLAddinGenerator
//   2.7 template link:
//     if (addin-host == `outlook`) config.json.json-taskpane.default.[programming-language]
//     else config[addin-host].[capabilities].default.[programming-language]
// Case 3: TEAMSFX_OFFICE_ADDIN = true AND TEAMSFX_OFFICE_XML_ADDIN = true
//   3.1 project-type option: `office-addin-type`
//   3.2 addin-host: not show but will use `wxpo` internally
//   3.3 capabilities options: [`json-taskpane`, `office-addin-import`, `office-content-addin`]
//   3.4 programming-language options: [`typescript`, `javascript`]
//   3.5 office-addin-framework-type options: [`default`, `react`]
//     if (capabilities == `json-taskpane`) then [`default`, `react`]
//     else if (capabilities == `office-addin-import`) then [`default`] (skip in UI)
//     else if (capabilities == `office-content-addin`) then [`default`] (skip in UI)
//   3.6 generator class: OfficeAddinGenerator
//   3.7 template link: config.json.[capabilities].[office-addin-framework-type].[programming-language]
// case 4: TEAMSFX_OFFICE_ADDIN = true AND TEAMSFX_OFFICE_XML_ADDIN = fasle
//   4.1 project-type option: `office-addin-type`
//   4.2 addin-host: not show but will use `wxpo` internally
//   4.3 capabilities options: [`jason-taskpane`, `office-addin-import`, `office-content-addin`]
//   4.4 programming-language options: [`typescript`, `javascript`]
//   4.5 office-addin-framework-type options: [`default`, `react`]
//     if (capabilities == `json-taskpane`) then [`default`, `react`]
//     else if (capabilities == `office-addin-import`) then [`default`] (skip in UI)
//     else if (capabilities == `office-content-addin`) then [`default`] (skip in UI)
//   4.6 generator class: OfficeAddinGenerator
//   4.7 template link: config.json.[capabilities].[office-addin-framework-type].[programming-language]
///////////////////////////////////////////////////////////////////////////////////////////////////////
