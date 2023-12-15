// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Stage, UserError } from "@microsoft/teamsfx-api";

import {
  LocalEnvManager,
  MetadataV3,
  envUtil,
  metadataUtil,
  pathUtils,
} from "@microsoft/teamsfx-core";
import * as fs from "fs-extra";
import * as path from "path";
import * as uuid from "uuid";
import * as vscode from "vscode";
import VsCodeLogInstance from "../commonlib/log";

import * as globalVariables from "../globalVariables";
import { core, getSystemInputs } from "../handlers";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import { allRunningDebugSessions } from "./teamsfxTaskHandler";

import { ExtensionErrors, ExtensionSource } from "../error";
import { VS_CODE_UI } from "../extension";

export async function getProjectRoot(
  folderPath: string,
  folderName: string
): Promise<string | undefined> {
  const projectRoot: string = path.join(folderPath, folderName);
  const projectExists: boolean = await fs.pathExists(projectRoot);
  return projectExists ? projectRoot : undefined;
}

export async function getNpmInstallLogInfo(): Promise<any> {
  const localEnvManager = new LocalEnvManager(VsCodeLogInstance, ExtTelemetry.reporter);
  return await localEnvManager.getNpmInstallLogInfo();
}

export async function getTestToolLogInfo(): Promise<string | undefined> {
  const localEnvManager = new LocalEnvManager(VsCodeLogInstance, ExtTelemetry.reporter);
  if (!globalVariables.workspaceUri?.fsPath) {
    return undefined;
  }
  return await localEnvManager.getTestToolLogInfo(globalVariables.workspaceUri?.fsPath);
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

export async function checkAndSkipDebugging(): Promise<boolean> {
  // skip debugging if there is already a debug session
  if (allRunningDebugSessions.size > 0) {
    VsCodeLogInstance.warning("Skip debugging because there is already a debug session.");
    endLocalDebugSession();
    return Promise.resolve(true);
  }
  return Promise.resolve(false);
}

export class Step {
  private currentStep: number;
  public readonly totalSteps: number;
  constructor(totalSteps: number) {
    this.currentStep = 1;
    this.totalSteps = totalSteps;
  }

  getPrefix(): string {
    return `(${this.currentStep++}/${this.totalSteps})`;
  }
}

export async function getV3TeamsAppId(projectPath: string, env: string): Promise<string> {
  const result = await envUtil.readEnv(projectPath, env, false);
  if (result.isErr()) {
    throw result.error;
  }

  const teamsAppIdKey = (await getTeamsAppKeyName(env)) || "TEAMS_APP_ID";
  const teamsAppId = result.value[teamsAppIdKey];
  if (teamsAppId === undefined) {
    throw new UserError(
      ExtensionSource,
      ExtensionErrors.TeamsAppIdNotFoundError,
      `TEAMS_APP_ID is missing in ${env} environment.`
    );
  }

  return teamsAppId;
}

export async function getTeamsAppKeyName(env?: string): Promise<string | undefined> {
  const templatePath = pathUtils.getYmlFilePath(globalVariables.workspaceUri!.fsPath, env);
  const maybeProjectModel = await metadataUtil.parse(templatePath, env);
  if (maybeProjectModel.isErr()) {
    return undefined;
  }
  const projectModel = maybeProjectModel.value;
  if (projectModel.provision?.driverDefs && projectModel.provision.driverDefs.length > 0) {
    for (const driver of projectModel.provision.driverDefs) {
      if (driver.uses === "teamsApp/create") {
        return driver.writeToEnvironmentFile?.teamsAppId;
      }
    }
  }
  return undefined;
}

export async function triggerV3Migration(): Promise<void> {
  const inputs = getSystemInputs();
  inputs.stage = Stage.debug;
  const result = await core.phantomMigrationV3(inputs);
  if (result.isErr()) {
    await vscode.debug.stopDebugging();
    throw result.error;
  }
  // reload window to terminate debugging
  await VS_CODE_UI.reload();
}

// Only work in ts/js project
export function isTestToolEnabledProject(workspacePath: string): boolean {
  const testToolYmlPath = path.join(workspacePath, MetadataV3.testToolConfigFile);
  if (fs.pathExistsSync(testToolYmlPath)) {
    return true;
  }
  return false;
}
