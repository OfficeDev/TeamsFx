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
  if (isPreviewFeaturesEnabled()) {
    process.env[FeatureFlagName.BotNotification] = "true";
    process.env[FeatureFlagName.M365App] = "true";
    process.env[FeatureFlagName.AadManifest] = "true";
    process.env[FeatureFlagName.ApiConnect] = "true";
    process.env[FeatureFlagName.DeployManifest] = "true";
  }
}

export function isBotNotificationEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.BotNotification, false);
}

export function isPreviewFeaturesEnabled(): boolean {
  return true;
}

export function isCLIDotNetEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.CLIDotNet, false);
}

export function isOfficeAddinEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlagName.OfficeAddin, false);
}
