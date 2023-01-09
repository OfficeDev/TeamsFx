// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import path from "path";
import fs from "fs-extra";
import { MigrationContext } from "./migrationContext";
import { isObject } from "lodash";
import { FileType, namingConverterV3 } from "./MigrationUtils";
import { EOL } from "os";
import {
  AppPackageFolderName,
  AzureSolutionSettings,
  Inputs,
  Platform,
  ProjectSettings,
  ProjectSettingsV3,
} from "@microsoft/teamsfx-api";
import { CoreHookContext } from "../../types";
import semver from "semver";
import { getProjectSettingPathV3, getProjectSettingPathV2 } from "../projectSettingsLoader";
import {
  Metadata,
  MetadataV2,
  MetadataV3,
  VersionInfo,
  VersionSource,
  VersionState,
} from "../../../common/versionMetadata";
import { MANIFEST_TEMPLATE_CONSOLIDATE } from "../../../component/resource/appManifest/constants";
import { VersionForMigration } from "../types";
import { getLocalizedString } from "../../../common/localizeUtils";
import { TOOLS } from "../../globalVars";
import { load } from "js-yaml";

// read json files in states/ folder
export async function readJsonFile(context: MigrationContext, filePath: string): Promise<any> {
  const filepath = path.join(context.projectPath, filePath);
  if (await fs.pathExists(filepath)) {
    const obj = fs.readJson(filepath);
    return obj;
  }
}

// read bicep file content
export async function readBicepContent(context: MigrationContext): Promise<any> {
  const bicepFilePath = path.join(getTemplateFolderPath(context), "azure", "provision.bicep");
  const bicepFileExists = await context.fsPathExists(bicepFilePath);
  return bicepFileExists
    ? fs.readFileSync(path.join(context.projectPath, bicepFilePath), "utf8")
    : "";
}

// get template folder path
export function getTemplateFolderPath(context: MigrationContext): string {
  const inputs: Inputs = context.arguments[context.arguments.length - 1];
  return inputs.platform === Platform.VS ? "Templates" : "templates";
}

// read file names list under the given path
export function fsReadDirSync(context: MigrationContext, _path: string): string[] {
  const dirPath = path.join(context.projectPath, _path);
  return fs.readdirSync(dirPath);
}

// env variables in this list will be only convert into .env.{env} when migrating {env}.userdata
const skipList = [
  "state.fx-resource-aad-app-for-teams.clientSecret",
  "state.fx-resource-bot.botPassword",
  "state.fx-resource-apim.apimClientAADClientSecret",
  "state.fx-resource-azure-sql.adminPassword",
];

// convert any obj names if can be converted (used in states and configs migration)
export function jsonObjectNamesConvertV3(
  obj: any,
  prefix: string,
  parentKeyName: string,
  filetype: FileType,
  bicepContent: any
): string {
  let returnData = "";
  if (isObject(obj)) {
    for (const keyName of Object.keys(obj)) {
      returnData +=
        parentKeyName === ""
          ? jsonObjectNamesConvertV3(obj[keyName], prefix, prefix + keyName, filetype, bicepContent)
          : jsonObjectNamesConvertV3(
              obj[keyName],
              prefix,
              parentKeyName + "." + keyName,
              filetype,
              bicepContent
            );
    }
  } else if (!skipList.includes(parentKeyName)) {
    const res = namingConverterV3(parentKeyName, filetype, bicepContent);
    if (res.isOk()) return res.value + "=" + obj + EOL;
  } else return "";
  return returnData;
}

export async function getProjectVersion(ctx: CoreHookContext): Promise<VersionInfo> {
  const projectPath = getParameterFromCxt(ctx, "projectPath", "");
  return await getProjectVersionFromPath(projectPath);
}

export function migrationNotificationMessage(versionForMigration: VersionForMigration): string {
  if (versionForMigration.platform === Platform.VS) {
    return getLocalizedString("core.migrationV3.VS.Message", "Visual Studio 2022 17.5 Preview");
  }
  const res = getLocalizedString(
    "core.migrationV3.Message",
    MetadataV2.platformVersion[versionForMigration.platform]
  );
  return res;
}

export function getDownloadLinkByVersionAndPlatform(version: string, platform: Platform): string {
  let anchorInLink = "vscode";
  if (platform === Platform.VS) {
    anchorInLink = "visual-studio";
  } else if (platform === Platform.CLI) {
    anchorInLink = "cli";
  }
  return `${Metadata.versionMatchLink}#${anchorInLink}`;
}

export function outputCancelMessage(version: string, platform: Platform): void {
  TOOLS?.logProvider.warning(`[core] Upgrade cancelled.`);
  const link = getDownloadLinkByVersionAndPlatform(version, platform);
  if (platform === Platform.VSCode) {
    TOOLS?.logProvider.warning(
      `[core] Notice upgrade to new configuration files is a must-have to continue to use current version Teams Toolkit. If you want to upgrade, please run command (Teams: Upgrade project) or click the “Upgrade project” button on tree view to trigger the upgrade.`
    );
    TOOLS?.logProvider.warning(
      `[core]If you are not ready to upgrade and want to continue to use the old version Teams Toolkit ${MetadataV2.platformVersion[platform]}, please find it in ${link} and install it.`
    );
  } else if (platform === Platform.VS) {
    TOOLS?.logProvider.warning(
      `[core] Notice upgrade to new configuration files is a must-have to continue to use current version Teams Toolkit. If you want to upgrade, please trigger this command again.`
    );
    TOOLS?.logProvider.warning(
      `[core]If you are not ready to upgrade and want to continue to use the old version Teams Toolkit ${MetadataV2.platformVersion[platform]}, please find it in ${link} and install it.`
    );
  } else {
    TOOLS?.logProvider.warning(
      `[core] Notice upgrade to new configuration files is a must-have to continue to use current version Teams Toolkit CLI. If you want to upgrade, please trigger this command again.`
    );
    TOOLS?.logProvider.warning(
      `[core]If you are not ready to upgrade and want to continue to use the old version Teams Toolkit CLI ${MetadataV2.platformVersion[platform]}, please find it in ${link} and install it.`
    );
  }
}

export async function getProjectVersionFromPath(projectPath: string): Promise<VersionInfo> {
  const v3path = getProjectSettingPathV3(projectPath);
  if (await fs.pathExists(v3path)) {
    const settings = await fs.readFile(v3path, "utf8");
    const content = load(settings) as any;
    return {
      version: content.version || MetadataV3.projectVersion,
      source: VersionSource.teamsapp,
    };
  }
  const v2path = getProjectSettingPathV2(projectPath);
  if (await fs.pathExists(v2path)) {
    const settings = await fs.readJson(v2path);
    return {
      version: settings.version || "",
      source: VersionSource.projectSettings,
    };
  }
  return {
    version: "",
    source: VersionSource.unknown,
  };
}

export async function getTrackingIdFromPath(projectPath: string): Promise<string> {
  const v3path = getProjectSettingPathV3(projectPath);
  if (await fs.pathExists(v3path)) {
    const settings = await fs.readJson(v3path);
    return settings.trackingId || "";
  }
  const v2path = getProjectSettingPathV2(projectPath);
  if (await fs.pathExists(v2path)) {
    const settings = await fs.readJson(v2path);
    if (settings.projectId) {
      return settings.projectId || "";
    }
  }
  return "";
}

export function getVersionState(info: VersionInfo): VersionState {
  if (
    info.source === VersionSource.projectSettings &&
    semver.gte(info.version, MetadataV2.projectVersion) &&
    semver.lte(info.version, MetadataV2.projectMaxVersion)
  ) {
    return VersionState.upgradeable;
  } else if (info.source === VersionSource.teamsapp && info.version === MetadataV3.projectVersion) {
    return VersionState.compatible;
  }
  return VersionState.unsupported;
}

export function getParameterFromCxt(
  ctx: CoreHookContext,
  key: string,
  defaultValue?: string
): string {
  const inputs = ctx.arguments[ctx.arguments.length - 1] as Inputs;
  const value = (inputs[key] as string) || defaultValue || "";
  return value;
}

export function getToolkitVersionLink(platform: Platform, projectVersion: string): string {
  return Metadata.versionMatchLink;
}

export function getCapabilitySsoStatus(projectSettings: ProjectSettings): {
  TabSso: boolean;
  BotSso: boolean;
} {
  let tabSso, botSso;
  if ((projectSettings as ProjectSettingsV3).components) {
    tabSso = (projectSettings as ProjectSettingsV3).components.some((component, index, obj) => {
      return component.name === "teams-tab" && component.sso == true;
    });
    botSso = (projectSettings as ProjectSettingsV3).components.some((component, index, obj) => {
      return component.name === "teams-bot" && component.sso == true;
    });
  } else {
    // For projects that does not componentize.
    const capabilities = (projectSettings.solutionSettings as AzureSolutionSettings).capabilities;
    tabSso = capabilities.includes("TabSso");
    botSso = capabilities.includes("BotSso");
  }

  return {
    TabSso: tabSso,
    BotSso: botSso,
  };
}

export function generateAppIdUri(capabilities: { TabSso: boolean; BotSso: boolean }): string {
  if (capabilities.TabSso && !capabilities.BotSso) {
    return "api://{{state.fx-resource-frontend-hosting.domain}}/{{state.fx-resource-aad-app-for-teams.clientId}}";
  } else if (capabilities.TabSso && capabilities.BotSso) {
    return "api://{{state.fx-resource-frontend-hosting.domain}}/botid-{{state.fx-resource-bot.botId}}";
  } else if (!capabilities.TabSso && capabilities.BotSso) {
    return "api://botid-{{state.fx-resource-bot.botId}}";
  } else {
    return "api://{{state.fx-resource-aad-app-for-teams.clientId}}";
  }
}

export function replaceAppIdUri(manifest: string, appIdUri: string): string {
  const appIdUriRegex = /{{+ *state\.fx\-resource\-aad\-app\-for\-teams\.applicationIdUris *}}+/g;
  if (manifest.match(appIdUriRegex)) {
    manifest = manifest.replace(appIdUriRegex, appIdUri);
  }

  return manifest;
}

export async function readAndConvertUserdata(
  context: MigrationContext,
  filePath: string,
  bicepContent: any
): Promise<string> {
  let returnAnswer = "";

  const userdataContent = await fs.readFile(path.join(context.projectPath, filePath), "utf8");
  const lines = userdataContent.split(EOL);
  for (const line of lines) {
    if (line && line != "") {
      // in case that there are "="s in secrets
      const key_value = line.split("=");
      const res = namingConverterV3("state." + key_value[0], FileType.USERDATA, bicepContent);
      if (res.isOk()) returnAnswer += res.value + "=" + key_value.slice(1).join("=") + EOL;
    }
  }

  return returnAnswer;
}

export async function updateAndSaveManifestForSpfx(
  context: MigrationContext,
  manifest: string
): Promise<void> {
  const remoteTemplatePath = path.join(AppPackageFolderName, MANIFEST_TEMPLATE_CONSOLIDATE);
  const localTemplatePath = path.join(AppPackageFolderName, "manifest.template.local.json");

  const contentRegex = /\"\{\{\^config\.isLocalDebug\}\}.*\{\{\/config\.isLocalDebug\}\}\"/g;
  const remoteRegex = /\{\{\^config\.isLocalDebug\}\}.*\{\{\/config\.isLocalDebug\}\}\{/g;
  const localRegex = /\}\{\{\#config\.isLocalDebug\}\}.*\{\{\/config\.isLocalDebug\}\}/g;

  let remoteTemplate = manifest,
    localTemplate = manifest;

  // Replace contentUrls
  const placeholders = manifest.match(contentRegex);
  if (placeholders) {
    for (const placeholder of placeholders) {
      // Replace with local and remote url
      // Will only replace if one match found
      const remoteUrl = placeholder.match(remoteRegex);
      if (remoteUrl && remoteUrl.length == 1) {
        remoteTemplate = remoteTemplate.replace(
          placeholder,
          `"${remoteUrl[0].substring(24, remoteUrl[0].length - 25)}"`
        );
      }

      const localUrl = placeholder.match(localRegex);
      if (localUrl && localUrl.length == 1) {
        localTemplate = localTemplate.replace(
          placeholder,
          `"${localUrl[0].substring(25, localUrl[0].length - 24)}"`
        );
      }
    }
  }

  await context.fsWriteFile(remoteTemplatePath, remoteTemplate);
  await context.fsWriteFile(localTemplatePath, localTemplate);
}
