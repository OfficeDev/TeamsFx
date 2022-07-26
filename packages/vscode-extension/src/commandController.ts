// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Mutex } from "async-mutex";

import { FxError, Result } from "@microsoft/teamsfx-api";

import { VS_CODE_UI } from "./extension";
import * as ExtTelemetry from "./telemetry/telemetry";
import { TelemetryEvent, TelemetryProperty } from "./telemetry/extTelemetryEvents";
import treeViewManager from "./treeview/treeViewManager";
import { getTriggerFromProperty } from "./utils/commonUtils";
import { localize } from "./utils/localizeUtils";

type CommandHandler = (args?: unknown[]) => Promise<Result<unknown, FxError>>;

interface TeamsFxCommand {
  name: string;
  callback: CommandHandler;
  blockTooltip?: string;
}

class CommandController {
  private static instance: CommandController;

  private commandMap: Map<string, TeamsFxCommand>;
  private exclusiveCommands: Set<string>;
  private runningCommand: string | undefined;
  private mutex: Mutex;

  private constructor() {
    this.commandMap = new Map<string, TeamsFxCommand>();
    this.mutex = new Mutex();
    this.exclusiveCommands = new Set([
      "fx-extension.create",
      "fx-extension.init",
      "fx-extension.addFeature",
      "fx-extension.openManifest",
      "fx-extension.provision",
      "fx-extension.build",
      "fx-extension.deploy",
      "fx-extension.publish",
    ]);
  }

  public static getInstance() {
    if (!CommandController.instance) {
      CommandController.instance = new CommandController();
    }
    return CommandController.instance;
  }

  public registerCommand(
    commandName: string,
    commandHandler: CommandHandler,
    runningLabelKey?: string
  ) {
    let blockTooltip = "";
    if (runningLabelKey) {
      blockTooltip = localize(
        `teamstoolkit.commandsTreeViewProvider.${runningLabelKey}.blockTooltip`
      );
    }
    this.commandMap.set(commandName, {
      name: commandName,
      callback: commandHandler,
      blockTooltip,
    });
  }

  public async runCommand(commandName: string, args: unknown[]) {
    if (!this.exclusiveCommands.has(commandName)) {
      return this.runNonBlockingCommand(commandName, args);
    }
    if (this.runningCommand) {
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.TreeViewCommandConcurrentExecution, {
        ...getTriggerFromProperty(args),
        [TelemetryProperty.RunningCommand]: this.runningCommand,
        [TelemetryProperty.BlockedCommand]: commandName,
      });
      const command = this.commandMap.get(this.runningCommand);
      const blockedTooltip = command?.blockTooltip;
      if (blockedTooltip) {
        VS_CODE_UI.showMessage("warn", blockedTooltip, false);
      }
      return;
    }
    this.mutex.runExclusive(async () => await this.runBlockingCommand(commandName, args));
  }

  private runNonBlockingCommand(commandName: string, args: unknown[]) {
    const command = this.commandMap.get(commandName);
    if (command) {
      command.callback(args);
    }
  }

  private async runBlockingCommand(commandName: string, args: unknown[]) {
    this.runningCommand = commandName;
    const command = this.commandMap.get(commandName);
    const blockedCommands = [...this.exclusiveCommands.values()].filter((x) => x !== commandName);
    await treeViewManager.setRunningCommand(commandName, blockedCommands, command?.blockTooltip);
    if (command) {
      await command.callback(args);
    }
    this.runningCommand = undefined;
    await treeViewManager.restoreRunningCommand(blockedCommands);
  }

  public dispose() {}
}

export default CommandController.getInstance();
