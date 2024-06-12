// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { exec } from "child_process";
import * as fs from "fs-extra";
import * as os from "os";
import * as path from "path";
import { format } from "util";
import { ConfigFolderName } from "@microsoft/teamsfx-api";
import { glob } from "glob";
import { workspace } from "vscode";
import { workspaceUri } from "../globalVariables";
import { localize } from "./localizeUtils";

export function isWindows() {
  return os.type() === "Windows_NT";
}

export function isMacOS() {
  return os.type() === "Darwin";
}

export function isLinux() {
  return os.type() === "Linux";
}

export function openFolderInExplorer(folderPath: string): void {
  const command = format('start "" "%s"', folderPath);
  exec(command);
}

export async function isM365Project(workspacePath: string): Promise<boolean> {
  const projectSettingsPath = path.resolve(
    workspacePath,
    `.${ConfigFolderName}`,
    "configs",
    "projectSettings.json"
  );

  if (await fs.pathExists(projectSettingsPath)) {
    const projectSettings = await fs.readJson(projectSettingsPath);
    return projectSettings.isM365;
  } else {
    return false;
  }
}

export function delay(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export async function hasAdaptiveCardInWorkspace(): Promise<boolean> {
  // Skip large files which are unlikely to be adaptive cards to prevent performance impact.
  const fileSizeLimit = 1024 * 1024;

  if (workspaceUri) {
    const files = await glob(workspaceUri.path + "/**/*.json", {
      ignore: ["**/node_modules/**", "./node_modules/**"],
    });
    for (const file of files) {
      let content = "";
      let fd = -1;
      try {
        fd = await fs.open(file, "r");
        const stat = await fs.fstat(fd);
        // limit file size to prevent performance impact
        if (stat.size > fileSizeLimit) {
          continue;
        }

        // avoid security issue
        // https://github.com/OfficeDev/TeamsFx/security/code-scanning/2664
        const buffer = new Uint8Array(fileSizeLimit);
        const { bytesRead } = await fs.read(fd, buffer, 0, buffer.byteLength, 0);
        content = new TextDecoder().decode(buffer.slice(0, bytesRead));
      } catch (e) {
        // skip invalid files
        continue;
      } finally {
        if (fd >= 0) {
          fs.close(fd).catch(() => {});
        }
      }

      if (isAdaptiveCard(content)) {
        return true;
      }
    }
  }

  return false;
}

function isAdaptiveCard(content: string): boolean {
  const pattern = /"type"\s*:\s*"AdaptiveCard"/;
  return pattern.test(content);
}

export async function getLocalDebugMessageTemplate(isWindows: boolean): Promise<string> {
  const enabledTestTool = await isTestToolEnabled();

  if (isWindows) {
    return enabledTestTool
      ? localize("teamstoolkit.handlers.localDebugDescription.enabledTestTool")
      : localize("teamstoolkit.handlers.localDebugDescription");
  }

  return enabledTestTool
    ? localize("teamstoolkit.handlers.localDebugDescription.enabledTestTool.fallback")
    : localize("teamstoolkit.handlers.localDebugDescription.fallback");
}

// check if test tool is enabled in scaffolded project
async function isTestToolEnabled(): Promise<boolean> {
  if (workspace.workspaceFolders && workspace.workspaceFolders.length > 0) {
    const workspaceFolder = workspace.workspaceFolders[0];
    const workspacePath: string = workspaceFolder.uri.fsPath;

    const testToolYamlPath = path.join(workspacePath, "teamsapp.testtool.yml");
    return fs.pathExists(testToolYamlPath);
  }

  return false;
}
