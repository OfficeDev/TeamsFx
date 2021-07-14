// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { Argv } from "yargs";
import { YargsCommand } from "../yargsCommand";
import { FxError, Result, ok, LogLevel } from "@microsoft/teamsfx-api";
import { UserSettings, CliConfigOptions, CliConfigTelemetry } from "../userSetttings";
import CLILogProvider from "../commonlib/log";

export class ConfigGet extends YargsCommand {
  public readonly commandHead = `get`;
  public readonly command = `${this.commandHead} <option>`;
  public readonly description = "Get user settings.";

  public builder(yargs: Argv): Argv<any> {
    return yargs.positional("option", {
      description: "User settings option",
      type: "string",
      choices: [CliConfigOptions.Telemetry],
    });
  }

  public async runCommand(args: { [argName: string]: string }): Promise<Result<null, FxError>> {
    const result = UserSettings.getConfigSync();
    if (result.isErr()) {
      return result;
    }

    const config = result.value;
    switch (args.option) {
      case CliConfigOptions.Telemetry:
        CLILogProvider.necessaryLog(LogLevel.Info, JSON.stringify(config.telemetry, null, 2), true);
        return ok(null);
    }

    CLILogProvider.necessaryLog(LogLevel.Info, JSON.stringify(config, null, 2), true);
    return ok(null);
  }
}

export class ConfigSet extends YargsCommand {
  public readonly commandHead = `set`;
  public readonly command = `${this.commandHead} <option> <value>`;
  public readonly description = "Set user settings.";

  public builder(yargs: Argv): Argv<any> {
    return yargs
      .positional("option", {
        describe: "User settings option",
        type: "string",
        choices: [CliConfigOptions.Telemetry],
      })
      .positional("value", {
        describe: "Option value",
        type: "string",
        choices: [CliConfigTelemetry.On, CliConfigTelemetry.Off],
      });
  }

  public async runCommand(args: { [argName: string]: string }): Promise<Result<null, FxError>> {
    switch (args.option) {
      case CliConfigOptions.Telemetry:
        const opt = { [args.option]: args.value };
        const result = UserSettings.setConfigSync(opt);
        if (result.isErr()) {
          CLILogProvider.necessaryLog(LogLevel.Error, "Configure user settings failed");
          return result;
        }
    }

    CLILogProvider.necessaryLog(LogLevel.Info, "Configure user settings successful.");
    return ok(null);
  }
}

export default class Config extends YargsCommand {
  public readonly commandHead = `config`;
  public readonly command = `${this.commandHead} <action>`;
  public readonly description = "Configure user settings.";

  public readonly subCommands: YargsCommand[] = [new ConfigGet(), new ConfigSet()];

  public builder(yargs: Argv): Argv<any> {
    this.subCommands.forEach((cmd) => {
      yargs.command(cmd.command, cmd.description, cmd.builder.bind(cmd), cmd.handler.bind(cmd));
    });
    return yargs.version(false);
  }

  public async runCommand(args: { [argName: string]: string }): Promise<Result<null, FxError>> {
    return ok(null);
  }
}
