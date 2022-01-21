// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { LogProvider, ProjectSettings } from "@microsoft/teamsfx-api";
import * as path from "path";
import detectPort from "detect-port";

import { FolderName } from "./constants";
import { loadTeamsFxDevScript } from "./packageJsonHelper";
import { ProjectSettingsHelper } from "./projectSettingsHelper";

const frontendPortsV1 = [3000];
const frontendPorts = [53000];
const simpleAuthPorts = [55000];
const backendDebugPortRegex = /--inspect[\s]*=[\s"']*9229/im;
const backendDebugPorts = [9229];
const backendServicePortRegex = /--port[\s"']*7071/im;
const backendServicePorts = [7071];
const botDebugPortRegex = /--inspect[\s]*=[\s"']*9239/im;
const botDebugPorts = [9239];
const botServicePorts = [3978];

async function detectPortListening(port: number, logger?: LogProvider): Promise<boolean> {
  try {
    logger?.info(`Start to detect port: ${port}`);
    const portChosen = await detectPort(port);
    logger?.info(`Detect port successfully. Port is in use: ${portChosen !== port}`);
    return portChosen !== port;
  } catch (error: any) {
    // ignore any error to not block debugging
    logger?.warning(`Failed to detect port. Start-start${error?.message} `);
    return false;
  }
}

export async function getPortsInUse(
  projectPath: string,
  projectSettings: ProjectSettings,
  logger?: LogProvider,
  ignoreDebugPort?: boolean
): Promise<number[]> {
  const ports: number[] = [];

  const includeFrontend = ProjectSettingsHelper.includeFrontend(projectSettings);
  if (includeFrontend) {
    const migrateFromV1 = ProjectSettingsHelper.isMigrateFromV1(projectSettings);
    if (!migrateFromV1) {
      ports.push(...frontendPorts);
      ports.push(...simpleAuthPorts);
    } else {
      ports.push(...frontendPortsV1);
    }
  }

  const includeBackend = ProjectSettingsHelper.includeBackend(projectSettings);
  if (includeBackend) {
    ports.push(...backendServicePorts);
    if (!(ignoreDebugPort === true)) {
      const backendDevScript = await loadTeamsFxDevScript(
        path.join(projectPath, FolderName.Function)
      );
      if (backendDevScript === undefined || backendDebugPortRegex.test(backendDevScript)) {
        ports.push(...backendDebugPorts);
      }
    }
  }
  const includeBot = ProjectSettingsHelper.includeBot(projectSettings);
  if (includeBot) {
    ports.push(...botServicePorts);
    if (!(ignoreDebugPort === true)) {
      const botDevScript = await loadTeamsFxDevScript(path.join(projectPath, FolderName.Bot));
      if (botDevScript === undefined || botDebugPortRegex.test(botDevScript)) {
        ports.push(...botDebugPorts);
      }
    }
  }

  const portsInUse: number[] = [];
  for (const port of ports) {
    if (await detectPortListening(port, logger)) {
      portsInUse.push(port);
    }
  }
  return portsInUse;
}
