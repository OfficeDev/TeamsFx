// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import * as fs from "fs-extra";
import * as os from "os";
import {
  CryptoProvider,
  err,
  FxError,
  Inputs,
  Json,
  LogProvider,
  ok,
  Platform,
  Result,
  TelemetryReporter,
  v2,
  ProjectSettings,
} from "@microsoft/teamsfx-api";
import { LocalSettingsProvider } from "../../../../common/localSettingsProvider";
import { ProjectSettingsHelper } from "../../../../common/local/projectSettingsHelper";
import * as Launch from "./util/launch";
import * as LaunchNext from "./util/launchNext";
import * as Tasks from "./util/tasks";
import * as TasksNext from "./util/tasksNext";
import * as Settings from "./util/settings";
import { TelemetryEventName, TelemetryUtils } from "./util/telemetry";
import { ScaffoldLocalDebugSettingsError } from "./error";
import { BotHostTypes } from "../../../../common";

export async function scaffoldLocalDebugSettings(
  ctx: v2.Context,
  inputs: Inputs,
  localSettings?: Json,
  generateLocalSettingsFile = true
): Promise<Result<Json, FxError>> {
  return await _scaffoldLocalDebugSettings(
    ctx.projectSetting,
    inputs,
    ctx.telemetryReporter,
    ctx.logProvider,
    ctx.cryptoProvider,
    localSettings,
    generateLocalSettingsFile
  );
}

export async function _scaffoldLocalDebugSettings(
  projectSetting: ProjectSettings,
  inputs: Inputs,
  telemetryReporter: TelemetryReporter,
  logProvider: LogProvider,
  cryptoProvider: CryptoProvider,
  localSettings?: Json,
  generateLocalSettingsFile = true
): Promise<Result<Json, FxError>> {
  const isSpfx = ProjectSettingsHelper.isSpfx(projectSetting);
  const includeFrontend = ProjectSettingsHelper.includeFrontend(projectSetting);
  const includeBackend = ProjectSettingsHelper.includeBackend(projectSetting);
  const includeBot = ProjectSettingsHelper.includeBot(projectSetting);
  const includeAAD = ProjectSettingsHelper.includeAAD(projectSetting);
  const includeSimpleAuth = ProjectSettingsHelper.includeSimpleAuth(projectSetting);
  const includeFuncHostedBot = ProjectSettingsHelper.includeFuncHostedBot(projectSetting);
  const botCapabilities = ProjectSettingsHelper.getBotCapabilities(projectSetting);
  const programmingLanguage = projectSetting.programmingLanguage ?? "";
  const isM365 = projectSetting.isM365;
  const includeOfficeAddin = ProjectSettingsHelper.includeOfficeAddin(projectSetting);

  const telemetryProperties = {
    platform: inputs.platform as string,
    spfx: isSpfx ? "true" : "false",
    frontend: includeFrontend ? "true" : "false",
    function: includeBackend ? "true" : "false",
    bot: includeBot ? "true" : "false",
    auth: includeAAD && includeSimpleAuth ? "true" : "false",
    "bot-host-type": includeFuncHostedBot ? BotHostTypes.AzureFunctions : BotHostTypes.AppService,
    "bot-capabilities": JSON.stringify(botCapabilities),
    "programming-language": programmingLanguage,
  };
  TelemetryUtils.init(telemetryReporter);
  TelemetryUtils.sendStartEvent(TelemetryEventName.scaffoldLocalDebugSettings, telemetryProperties);
  try {
    // scaffold for both vscode and cli
    if (inputs.platform === Platform.VSCode || inputs.platform === Platform.CLI) {
      if (isSpfx) {
        // Only generate launch.json and tasks.json for SPFX
        const launchConfigurations = Launch.generateSpfxConfigurations();
        const launchCompounds = Launch.generateSpfxCompounds();
        const tasks = Tasks.generateSpfxTasks();
        const tasksInputs = Tasks.generateInputs();

        //TODO: save files via context api
        await fs.ensureDir(`${inputs.projectPath}/.vscode/`);
        await updateJson(
          `${inputs.projectPath}/.vscode/launch.json`,
          {
            version: "0.2.0",
            configurations: launchConfigurations,
            compounds: launchCompounds,
          },
          LaunchNext.mergeLaunches
        );

        await updateJson(
          `${inputs.projectPath}/.vscode/tasks.json`,
          {
            version: "2.0.0",
            tasks: tasks,
            inputs: tasksInputs,
          },
          TasksNext.mergeTasks
        );
      } else {
        const launchConfigurations = isM365
          ? LaunchNext.generateM365Configurations(includeFrontend, includeBackend, includeBot)
          : (await useNewTasks(inputs.projectPath))
          ? LaunchNext.generateConfigurations(
              includeFrontend,
              includeBackend,
              includeBot,
              includeOfficeAddin
            )
          : Launch.generateConfigurations(includeFrontend, includeBackend, includeBot);
        const launchCompounds = isM365
          ? LaunchNext.generateM365Compounds(includeFrontend, includeBackend, includeBot)
          : (await useNewTasks(inputs.projectPath))
          ? LaunchNext.generateCompounds(
              includeFrontend,
              includeBackend,
              includeBot,
              includeOfficeAddin
            )
          : Launch.generateCompounds(includeFrontend, includeBackend, includeBot);

        const tasks = isM365
          ? TasksNext.generateM365Tasks(
              includeFrontend,
              includeBackend,
              includeBot,
              programmingLanguage
            )
          : (await useNewTasks(inputs.projectPath))
          ? TasksNext.generateTasks(
              includeFrontend,
              includeBackend,
              includeBot,
              includeFuncHostedBot,
              programmingLanguage,
              includeOfficeAddin
            )
          : Tasks.generateTasks(
              includeFrontend,
              includeBackend,
              includeBot,
              includeSimpleAuth,
              programmingLanguage
            );

        //TODO: save files via context api
        await fs.ensureDir(`${inputs.projectPath}/.vscode/`);
        await updateJson(
          `${inputs.projectPath}/.vscode/launch.json`,
          {
            version: "0.2.0",
            configurations: launchConfigurations,
            compounds: launchCompounds,
          },
          LaunchNext.mergeLaunches
        );

        await updateJson(
          `${inputs.projectPath}/.vscode/tasks.json`,
          {
            version: "2.0.0",
            tasks: tasks,
          },
          TasksNext.mergeTasks
        );
      }

      await updateJson(
        `${inputs.projectPath}/.vscode/settings.json`,
        Settings.generateSettings(includeBackend || includeFuncHostedBot, isSpfx),
        Settings.mergeSettings
      );
    }
  } catch (error: any) {
    const systemError = ScaffoldLocalDebugSettingsError(error);
    TelemetryUtils.sendErrorEvent(TelemetryEventName.scaffoldLocalDebugSettings, systemError);
    return err(systemError);
  }

  TelemetryUtils.sendSuccessEvent(
    TelemetryEventName.scaffoldLocalDebugSettings,
    telemetryProperties
  );
  return ok(localSettings as Json);
}

async function scaffoldLocalSettingsJson(
  projectSetting: ProjectSettings,
  inputs: Inputs,
  cryptoProvider: CryptoProvider,
  localSettings?: Json
): Promise<Json> {
  const localSettingsProvider = new LocalSettingsProvider(inputs.projectPath!);

  const includeFrontend = ProjectSettingsHelper.includeFrontend(projectSetting);
  const includeBackend = ProjectSettingsHelper.includeBackend(projectSetting);
  const includeBot = ProjectSettingsHelper.includeBot(projectSetting);
  const includeAAD = ProjectSettingsHelper.includeAAD(projectSetting);
  const includeSimpleAuth = ProjectSettingsHelper.includeSimpleAuth(projectSetting);

  if (localSettings !== undefined) {
    // Add local settings for the new added capability/resource
    localSettings = localSettingsProvider.incrementalInitV2(
      localSettings,
      includeBackend,
      includeBot,
      includeFrontend,
      includeAAD,
      includeSimpleAuth
    );
    await localSettingsProvider.saveJson(localSettings, cryptoProvider);
  } else {
    // Initialize a local settings on scaffolding
    localSettings = localSettingsProvider.initV2(
      includeFrontend,
      includeBackend,
      includeBot,
      includeSimpleAuth,
      includeAAD
    );
    await localSettingsProvider.saveJson(localSettings, cryptoProvider);
  }
  return localSettings;
}

export async function useNewTasks(projectPath?: string): Promise<boolean> {
  // for new project or project with "validate-local-prerequisites", use new tasks content
  const tasksJsonPath = `${projectPath}/.vscode/tasks.json`;
  if (await fs.pathExists(tasksJsonPath)) {
    try {
      const tasksContent = await fs.readFile(tasksJsonPath, "utf-8");
      return tasksContent.includes("fx-extension.validate-local-prerequisites");
    } catch (error) {
      return false;
    }
  }

  return true;
}

export async function updateJson(
  path: string,
  newData: Record<string, unknown>,
  mergeFunc: (
    existingData: Record<string, unknown>,
    newData: Record<string, unknown>
  ) => Record<string, unknown>
): Promise<void> {
  let finalData: Record<string, unknown>;
  if (await fs.pathExists(path)) {
    try {
      const existingData = await fs.readJSON(path);
      finalData = mergeFunc(existingData, newData);
    } catch (error) {
      // If failed to parse or edit the existing file, just overwrite completely
      finalData = newData;
    }
  } else {
    finalData = newData;
  }

  await fs.writeJSON(path, finalData, {
    spaces: 4,
    EOL: os.EOL,
  });
}
