// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs-extra";
import * as path from "path";
import * as uuid from "uuid";
import * as constants from "./constants";
import { ConfigFolderName, InputConfigsFolderName, UserError } from "@microsoft/teamsfx-api";
import VsCodeLogInstance from "../commonlib/log";
import * as ExtTelemetry from "../telemetry/telemetry";
import { getTeamsAppTelemetryInfoByEnv } from "../utils/commonUtils";
import { core, getSystemInputs, showError } from "../handlers";
import * as globalVariables from "../globalVariables";
import {
  LocalEnvManager,
  FolderName,
  isV3,
  environmentManager,
  ProjectSettingsHelper,
  PluginNames,
  GLOBAL_CONFIG,
  getResourceGroupInPortal,
} from "@microsoft/teamsfx-core";
import { allRunningDebugSessions } from "./teamsfxTaskHandler";

export async function getProjectRoot(
  folderPath: string,
  folderName: string
): Promise<string | undefined> {
  const projectRoot: string = path.join(folderPath, folderName);
  const projectExists: boolean = await fs.pathExists(projectRoot);
  return projectExists ? projectRoot : undefined;
}

function getLocalEnvWithPrefix(
  env: { [key: string]: string } | undefined,
  prefix: string
): { [key: string]: string } | undefined {
  if (env === undefined) {
    return undefined;
  }
  const result: { [key: string]: string } = {};
  for (const key of Object.keys(env)) {
    if (key.startsWith(prefix) && env[key]) {
      result[key.slice(prefix.length)] = env[key];
    }
  }
  return Object.keys(result).length > 0 ? result : undefined;
}

export function getFrontendLocalEnv(
  env: { [key: string]: string } | undefined
): { [key: string]: string } | undefined {
  return getLocalEnvWithPrefix(env, constants.frontendLocalEnvPrefix);
}

export function getBackendLocalEnv(
  env: { [key: string]: string } | undefined
): { [key: string]: string } | undefined {
  return getLocalEnvWithPrefix(env, constants.backendLocalEnvPrefix);
}

export function getAuthLocalEnv(
  env: { [key: string]: string } | undefined
): { [key: string]: string } | undefined {
  // SERVICE_PATH will also be included, but it has no side effect
  return getLocalEnvWithPrefix(env, constants.authLocalEnvPrefix);
}

export function getAuthServicePath(env: { [key: string]: string } | undefined): string | undefined {
  return env ? env[constants.authServicePathEnvKey] : undefined;
}

export function getBotLocalEnv(
  env: { [key: string]: string } | undefined
): { [key: string]: string } | undefined {
  return getLocalEnvWithPrefix(env, constants.botLocalEnvPrefix);
}

export async function isFxProject(folderPath: string): Promise<boolean> {
  return fs.pathExists(path.join(folderPath, `.${ConfigFolderName}`));
}

export async function hasTeamsfxBackend(): Promise<boolean> {
  if (!globalVariables.workspaceUri) {
    return false;
  }

  if (!globalVariables.isTeamsFxProject) {
    return false;
  }

  const backendRoot = await getProjectRoot(
    globalVariables.workspaceUri.fsPath,
    FolderName.Function
  );

  return backendRoot !== undefined;
}

export async function hasTeamsfxBot(): Promise<boolean> {
  if (!globalVariables.workspaceUri) {
    return false;
  }

  if (!globalVariables.isTeamsFxProject) {
    return false;
  }

  const botRoot = await getProjectRoot(globalVariables.workspaceUri.fsPath, FolderName.Bot);

  return botRoot !== undefined;
}

export async function getDebugConfig(
  isLocalSideloadingConfiguration: boolean,
  env?: string
): Promise<{ appId: string; env?: string } | undefined> {
  try {
    if (isV3()) {
      const inputs = getSystemInputs();
      const getConfigRes = await core.getProjectConfigV3(inputs);
      if (getConfigRes.isErr()) throw getConfigRes.error;
      const config = getConfigRes.value;
      if (!config)
        throw new UserError("extension", "GetConfigError", "Failed to get project config");
      if (isLocalSideloadingConfiguration) {
        const envInfo = config.envInfos["local"];
        if (!envInfo)
          throw new UserError("extension", "EnvConfigNotExist", "Local Env config not exist");
        const appId = envInfo.state["app-manifest"].teamsAppId as string;
        return { appId: appId, env: "local" };
      } else {
        if (env === undefined) {
          const inputs = getSystemInputs();
          inputs.ignoreConfigPersist = true;
          inputs.ignoreEnvInfo = false;
          const envRes = await core.getSelectedEnv(inputs);
          if (envRes.isErr()) {
            VsCodeLogInstance.warning(`No environment selected. ${envRes.error}`);
            return undefined;
          }
          env = envRes.value;
        }
        if (!env)
          throw new UserError(
            "extension",
            "GetSelectedEnvError",
            "Failed to get selected Env name"
          );
        const envInfo = config.envInfos[env];
        if (!envInfo)
          throw new UserError("extension", "EnvConfigNotExist", `Env '${env} ' config not exist`);
        const appId = envInfo.state["app-manifest"].teamsAppId as string;
        return { appId: appId, env: env };
      }
    } else {
      if (isLocalSideloadingConfiguration) {
        const appInfo = getTeamsAppTelemetryInfoByEnv(environmentManager.getLocalEnvName());
        return { appId: appInfo?.appId as string, env: env };
      } else {
        // select env
        if (env === undefined) {
          const inputs = getSystemInputs();
          inputs.ignoreConfigPersist = true;
          inputs.ignoreEnvInfo = false;
          const envRes = await core.getSelectedEnv(inputs);
          if (envRes.isErr()) {
            VsCodeLogInstance.warning(`No environment selected. ${envRes.error}`);
            return undefined;
          }

          env = envRes.value;
        }

        // load env state
        const appInfo = getTeamsAppTelemetryInfoByEnv(env!);
        if (appInfo === undefined) {
          throw new UserError({
            name: "MissingTeamsAppId",
            message: `No teams app found in ${env} environment. Run Provision to ensure teams app is created.`,
            source: "preview",
          });
        }

        return { appId: appInfo.appId as string, env: env };
      }
    }
  } catch (error: any) {
    showError(error);
    return undefined;
  }
}

export async function getNpmInstallLogInfo(): Promise<any> {
  const localEnvManager = new LocalEnvManager(VsCodeLogInstance, ExtTelemetry.reporter);
  return await localEnvManager.getNpmInstallLogInfo();
}

export async function getPortsInUse(): Promise<number[]> {
  const localEnvManager = new LocalEnvManager(VsCodeLogInstance, ExtTelemetry.reporter);
  try {
    const projectPath = globalVariables.workspaceUri!.fsPath;
    const projectSettings = await localEnvManager.getProjectSettings(projectPath);
    return await localEnvManager.getPortsInUse(projectPath, projectSettings);
  } catch (error: any) {
    VsCodeLogInstance.warning(`Failed to check used ports. Error: ${error}`);
    return [];
  }
}

export async function getTeamsAppTenantId(): Promise<string | undefined> {
  try {
    const localEnvManager = new LocalEnvManager(VsCodeLogInstance, ExtTelemetry.reporter);
    const projectSettings = await localEnvManager.getProjectSettings(
      globalVariables.workspaceUri!.fsPath
    );
    const localEnvInfo = await localEnvManager.getLocalEnvInfo(
      globalVariables.workspaceUri!.fsPath,
      {
        projectId: projectSettings.projectId,
      }
    );
    if (localEnvInfo && localEnvInfo["state"] && localEnvInfo["state"][PluginNames.AAD]) {
      return localEnvInfo["state"][PluginNames.APPST].tenantId as string;
    }
    return undefined;
  } catch {
    // in case structure changes
    return undefined;
  }
}

export async function getLocalTeamsAppId(): Promise<string | undefined> {
  try {
    const localEnvManager = new LocalEnvManager(VsCodeLogInstance, ExtTelemetry.reporter);
    const projectSettings = await localEnvManager.getProjectSettings(
      globalVariables.workspaceUri!.fsPath
    );
    const localEnvInfo = await localEnvManager.getLocalEnvInfo(
      globalVariables.workspaceUri!.fsPath,
      {
        projectId: projectSettings.projectId,
      }
    );
    if (localEnvInfo && localEnvInfo["state"] && localEnvInfo["state"][PluginNames.APPST]) {
      return localEnvInfo["state"][PluginNames.APPST].teamsAppId as string;
    }
    return undefined;
  } catch {
    // in case structure changes
    return undefined;
  }
}

export async function getLocalBotId(): Promise<string | undefined> {
  try {
    const result = environmentManager.getEnvStateFilesPath(
      environmentManager.getLocalEnvName(),
      globalVariables.workspaceUri!.fsPath
    );
    const envJson = JSON.parse(fs.readFileSync(result.envState, "utf8"));
    return envJson[PluginNames.BOT].botId;
  } catch {
    return undefined;
  }
}

export async function getBotId(env: string): Promise<string | undefined> {
  try {
    if (env === environmentManager.getLocalEnvName()) {
      return await getLocalBotId();
    }

    const result = environmentManager.getEnvStateFilesPath(
      env,
      globalVariables.workspaceUri!.fsPath
    );
    const envJson = JSON.parse(fs.readFileSync(result.envState, "utf8"));
    return envJson[PluginNames.BOT].botId;
  } catch {
    return undefined;
  }
}

async function getResourceBaseName(env: string): Promise<string | undefined> {
  try {
    const azureParametersFilePath = path.join(
      globalVariables.workspaceUri!.fsPath,
      `.${ConfigFolderName}`,
      InputConfigsFolderName,
      `azure.parameters.${env}.json`
    );
    const azureParametersJson = JSON.parse(fs.readFileSync(azureParametersFilePath, "utf-8"));
    let result: string = azureParametersJson.parameters.provisionParameters.value.resourceBaseName;
    const placeholder = "{{state.solution.resourceNameSuffix}}";
    if (result.includes(placeholder)) {
      const envStateFilesPath = environmentManager.getEnvStateFilesPath(
        env,
        globalVariables.workspaceUri!.fsPath
      );
      const envJson = JSON.parse(fs.readFileSync(envStateFilesPath.envState, "utf8"));
      result = result.replace(placeholder, envJson[PluginNames.SOLUTION].resourceNameSuffix);
    }
    return result;
  } catch {
    return undefined;
  }
}

export async function getBotOutlookChannelLink(env: string): Promise<string> {
  const result = environmentManager.getEnvStateFilesPath(env, globalVariables.workspaceUri!.fsPath);
  const envJson = JSON.parse(fs.readFileSync(result.envState, "utf8"));
  const tenantId = envJson[GLOBAL_CONFIG].tenantId;
  const subscriptionId = envJson[GLOBAL_CONFIG].subscriptionId;
  const resourceGroupName = envJson[GLOBAL_CONFIG].resourceGroupName;

  const resourceGroupLink = getResourceGroupInPortal(subscriptionId, tenantId, resourceGroupName);
  const resourceBaseName = await getResourceBaseName(env);

  return `${resourceGroupLink}/providers/Microsoft.BotService/botServices/${resourceBaseName}/channelsReact`;
}

export async function loadPackageJson(path: string): Promise<any> {
  if (!(await fs.pathExists(path))) {
    VsCodeLogInstance.error(`Cannot load package.json from ${path}. File not found.`);
    return undefined;
  }

  const rpj = require("read-package-json-fast");
  try {
    return await rpj(path);
  } catch (error) {
    VsCodeLogInstance.error(`Cannot load package.json from ${path}. Error: ${error}`);
    return undefined;
  }
}

export class LocalDebugSession {
  static createSession() {
    const session = new LocalDebugSession(uuid.v4());
    return session;
  }
  static createInvalidSession() {
    return new LocalDebugSession();
  }

  readonly id: string;
  // Save the time when the event it sent for calculating time gaps.
  readonly eventTimes: { [eventName: string]: number | undefined } = {};
  readonly properties: { [key: string]: string } = {};
  readonly errorProps: string[] = [];
  readonly failedServices: { name: string; exitCode: number | undefined }[] = [];

  private constructor(id: string = DebugNoSessionId) {
    this.id = id;
  }
}

export const DebugNoSessionId = "no-session-id";
// Helper functions for local debug correlation-id, only used for telemetry
// Use a 2-element tuple to handle concurrent F5
const localDebugCorrelationIds: [LocalDebugSession, LocalDebugSession] = [
  LocalDebugSession.createInvalidSession(),
  LocalDebugSession.createInvalidSession(),
];
let current = 0;
export function startLocalDebugSession(): string {
  current = (current + 1) % 2;
  localDebugCorrelationIds[current] = LocalDebugSession.createSession();
  return getLocalDebugSessionId();
}

export function endLocalDebugSession() {
  localDebugCorrelationIds[current] = LocalDebugSession.createInvalidSession();
  current = (current + 1) % 2;
}

export function getLocalDebugSession(): LocalDebugSession {
  return localDebugCorrelationIds[current];
}

export function getLocalDebugSessionId(): string {
  return localDebugCorrelationIds[current].id;
}

export function checkAndSkipDebugging(): boolean {
  // skip debugging if there is already a debug session
  if (allRunningDebugSessions.size > 0) {
    VsCodeLogInstance.warning("Skip debugging because there is already a debug session.");
    endLocalDebugSession();
    return true;
  }
  return false;
}

// for telemetry use only
export async function getProjectComponents(): Promise<string | undefined> {
  const localEnvManager = new LocalEnvManager(VsCodeLogInstance, ExtTelemetry.reporter);
  try {
    const projectPath = globalVariables.workspaceUri!.fsPath;
    const projectSettings = await localEnvManager.getProjectSettings(projectPath);
    const result: { [key: string]: any } = { components: [] };
    if (ProjectSettingsHelper.isSpfx(projectSettings)) {
      result.components.push("spfx");
    }
    if (ProjectSettingsHelper.includeFrontend(projectSettings)) {
      result.components.push("frontend");
    }
    if (ProjectSettingsHelper.includeBot(projectSettings)) {
      result.components.push(`bot`);
      result.botHostType = ProjectSettingsHelper.includeFuncHostedBot(projectSettings)
        ? "azure-functions"
        : "app-service";
      result.botCapabilities = ProjectSettingsHelper.getBotCapabilities(projectSettings);
    }
    if (ProjectSettingsHelper.includeBackend(projectSettings)) {
      result.components.push("backend");
    }
    if (ProjectSettingsHelper.includeAAD(projectSettings)) {
      result.components.push("aad");
    }
    return JSON.stringify(result);
  } catch (error: any) {
    return undefined;
  }
}
