// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  globalStateGet,
  globalStateUpdate,
  LocalEnvManager,
  ProjectSettingsHelper,
  TaskDefinition,
} from "@microsoft/teamsfx-core";
import * as vscode from "vscode";
import { ConfigurationKey } from "../constants";
import { VS_CODE_UI } from "../extension";
import * as globalVariables from "../globalVariables";
import { getConfiguration } from "../utils/commonUtils";
import { runTask } from "./teamsfxTaskHandler";
import { createTask } from "./teamsfxTaskProvider";
import VsCodeLogInstance from "../commonlib/log";
import * as ExtTelemetry from "../telemetry/telemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/extTelemetryEvents";
import { ExtensionSource } from "../error";
import { localize } from "../utils/localizeUtils";
import { UserError } from "@microsoft/teamsfx-api";

export async function automaticNpmInstallHandler(
  excludeFrontend: boolean,
  excludeBackend: boolean,
  excludeBot: boolean
): Promise<void> {
  try {
    const state = await globalStateGet("automaticNpmInstall", false);
    if (state) {
      globalStateUpdate("automaticNpmInstall", false);
      // const configuration = getConfiguration(ConfigurationKey.AutomaticNpmInstall);
      const configuration = false;
      if (configuration && globalVariables.workspaceUri !== undefined) {
        const localEnvManager = new LocalEnvManager(
          VsCodeLogInstance,
          ExtTelemetry.reporter,
          VS_CODE_UI
        );
        const workspaceFolder = vscode.workspace.workspaceFolders![0];
        const workspacePath = workspaceFolder.uri.fsPath;
        const projectSettings = await localEnvManager.getProjectSettings(workspacePath);
        const tasks: Map<string, Promise<number | undefined>> = new Map<
          string,
          Promise<number | undefined>
        >();
        if (ProjectSettingsHelper.isSpfx(projectSettings)) {
          tasks.set(
            "spfx",
            runTask(await createTask(TaskDefinition.spfxInstall(workspacePath), workspaceFolder))
          );
        } else {
          if (!excludeFrontend && ProjectSettingsHelper.includeFrontend(projectSettings)) {
            tasks.set(
              "frontend",
              runTask(
                await createTask(TaskDefinition.frontendInstall(workspacePath), workspaceFolder)
              )
            );
          }
          if (!excludeBackend && ProjectSettingsHelper.includeBackend(projectSettings)) {
            tasks.set(
              "backend",
              runTask(
                await createTask(TaskDefinition.backendInstall(workspacePath), workspaceFolder)
              )
            );
          }
          if (!excludeBot && ProjectSettingsHelper.includeBot(projectSettings)) {
            tasks.set(
              "bot",
              runTask(await createTask(TaskDefinition.botInstall(workspacePath), workspaceFolder))
            );
          }
        }
        if (tasks.size > 0) {
          let properties: { [key: string]: string } = {};
          for (const key of tasks.keys()) {
            properties[key] = "true";
          }
          ExtTelemetry.sendTelemetryEvent(TelemetryEvent.AutomaticNpmInstallStart, properties);

          VS_CODE_UI.showMessage(
            "info",
            localize("teamstoolkit.handlers.automaticNpmInstall"),
            false,
            localize("teamstoolkit.handlers.disableAutomaticNpmInstall")
          ).then((selection) => {
            if (
              selection.isOk() &&
              selection.value === localize("teamstoolkit.handlers.disableAutomaticNpmInstall")
            ) {
              vscode.commands.executeCommand(
                "workbench.action.openSettings",
                ConfigurationKey.AutomaticNpmInstall
              );
              ExtTelemetry.sendTelemetryEvent(TelemetryEvent.ClickDisableAutomaticNpmInstall);
            }
          });

          const keys = tasks.keys();
          const exitCodes = await Promise.all(tasks.values());

          properties = {};
          for (const exitCode of exitCodes) {
            properties[keys.next().value] = exitCode + "";
          }
          const failed = exitCodes.some((exitCode) => exitCode !== 0);
          if (failed) {
            const error = new UserError(ExtensionSource, "NpmInstallFailed", "Npm install failed");
            ExtTelemetry.sendTelemetryErrorEvent(
              TelemetryEvent.AutomaticNpmInstall,
              error,
              properties
            );
          } else {
            properties[TelemetryProperty.Success] = TelemetrySuccess.Yes;
            ExtTelemetry.sendTelemetryEvent(TelemetryEvent.AutomaticNpmInstall, properties);
          }
        }
      }
    }
  } catch (error: any) {
    VsCodeLogInstance.warning(`Automatic npm install failed: ${error}`);
  }
}
