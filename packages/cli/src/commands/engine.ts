// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, LogLevel, Result, err, ok } from "@microsoft/teamsfx-api";
import {
  InputValidationError,
  MissingRequiredInputError,
  assembleError,
  getHashedEnv,
  isUserCancelError,
} from "@microsoft/teamsfx-core";
import { cloneDeep } from "lodash";
import { format } from "util";
import { TextType, colorize } from "../colorize";
import { logger } from "../commonlib/logger";
import { strings } from "../resource";
import CliTelemetry from "../telemetry/cliTelemetry";
import { helper } from "./helper";
import { CLICommand, CLICommandOption, CLIContext } from "./types";
import UI from "../userInteraction";
import { TelemetryProperty } from "../telemetry/cliTelemetryEvents";
import { cliSource } from "../constants";

// Licensed under the MIT license.
class CLIEngine {
  async start(rootCmd: CLICommand): Promise<void> {
    const root = cloneDeep(rootCmd);
    const args = process.argv.slice(2);

    // 1. find command
    const findRes = this.findCommand(rootCmd, args);
    const cmd = findRes.cmd;
    const remainingArgs = findRes.remainingArgs;
    // process.stdout.write("name:" + cmd.name + "\n");
    // console.log("find command:", cmd.fullName!);

    // 2. parse args
    const context = this.parseArgs(cmd, root, remainingArgs);

    // 3. --version
    if (context.globalOptionValues.version === true) {
      logger.info(rootCmd.version ?? "1.0.0");
      return;
    }

    // 4. --help
    if (context.globalOptionValues.help === true) {
      const helpText = helper.formatHelp(
        context.command,
        context.command.fullName !== root.fullName ? root : undefined
      );
      logger.info(helpText);
      return;
    }

    // 5. validate
    const validateRes = this.validateOptionsAndArguments(context.command);
    if (validateRes.isErr()) {
      this.processResult(context, validateRes.error);
      return;
    }

    // 6. run handler
    if (context.command.handler) {
      try {
        const handleRes = await context.command.handler(context);
        if (handleRes.isErr()) {
          this.processResult(context, handleRes.error);
        } else {
          this.processResult(context);
        }
      } catch (e) {
        const fxError = assembleError(e);
        this.processResult(context, fxError);
      }
    } else {
      const helpText = helper.formatHelp(rootCmd);
      logger.info(helpText);
    }
  }

  findCommand(model: CLICommand, args: string[]): { cmd: CLICommand; remainingArgs: string[] } {
    let i = 0;
    let cmd = model;
    for (; i < args.length; i++) {
      const arg = args[i];
      const command = cmd.commands?.find((c) => c.name === arg);
      if (command) {
        cmd = command;
      } else {
        break;
      }
    }
    cmd.fullName = [model.name, ...args.slice(0, i)].join(" ");
    const command = cloneDeep(cmd);
    return { cmd: command, remainingArgs: args.slice(i) };
  }

  parseArgs(command: CLICommand, rootCommand: CLICommand, args: string[]): CLIContext {
    let i = 0;
    let j = 0;
    const context: CLIContext = {
      command: command,
      optionValues: {},
      globalOptionValues: {},
      argumentValues: [],
      telemetryProperties: {},
    };
    const options = (rootCommand.options || []).concat(command.options || []);
    for (; i < args.length; i++) {
      const arg = args[i];
      if (arg.startsWith("-")) {
        const argName = arg.replace(/-/g, "");
        const option = options.find((o) => o.name === argName || o.shortName === argName);
        if (option) {
          if (option.type === "boolean") {
            if (args[i + 1] === "false") {
              option.value = false;
              ++i;
            } else if (args[i + 1] === "true") {
              option.value = true;
              ++i;
            } else {
              option.value = true;
            }
          } else {
            const value = args[++i];
            if (value) {
              option.value = value;
            }
          }
          const inputValues = command.options?.includes(option)
            ? context.optionValues
            : context.globalOptionValues;
          if (option.value !== undefined) inputValues[option.name] = option.value;
        }
      } else {
        if (command.arguments && command.arguments[j]) {
          command.arguments[j++].value = args[i];
          context.argumentValues.push(args[i]);
        }
      }
    }
    // for required options or arguments, set default value if not set
    if (command.options) {
      for (const option of command.options) {
        if (option.required && option.default !== undefined && option.value === undefined) {
          option.value = option.default;
          context.optionValues[option.name] = option.default;
        }
      }
    }
    if (command.arguments) {
      for (let i = 0; i < command.arguments.length; ++i) {
        const argument = command.arguments[i];
        if (argument.required && argument.default !== undefined && argument.value === undefined) {
          argument.value = argument.default;
          context.argumentValues[i] = argument.default as string;
        }
      }
    }

    // special process for global options
    // process interactive
    context.globalOptionValues.interactive =
      context.globalOptionValues.interactive === false ? false : true;

    // set log level
    const logLevel = context.globalOptionValues.debug ? LogLevel.Debug : LogLevel.Info;
    logger.logLevel = logLevel;

    // set root folder
    const projectPath = context.optionValues.folder as string;
    if (projectPath) {
      CliTelemetry.withRootFolder(projectPath);
    }

    UI.interactive = context.globalOptionValues.interactive as boolean;

    if (context.globalOptionValues.interactive) {
      const sameKeys = Object.keys(context.optionValues).filter(
        (k) => k !== "folder" && k in args && context.optionValues[k] !== undefined
      );
      if (sameKeys.length > 0) {
        /// only if there are intersects between parameters and arguments, show the log,
        /// because it means some parameters will be used by fx-core.
        logger.info(
          `Some arguments/options are useless because the interactive mode is opened.` +
            ` If you want to run the command non-interactively, add '--interactive false' after your command` +
            ` or set the global setting by 'teamsfx config set interactive false'.`
        );
      }
    }

    return context;
  }

  validateOptionsAndArguments(
    command: CLICommand
  ): Result<undefined, InputValidationError | MissingRequiredInputError> {
    if (command.options) {
      for (const option of command.options) {
        const res = this.validateOption(option);
        if (res.isErr()) {
          return err(res.error);
        }
      }
    }
    if (command.arguments) {
      for (const argument of command.arguments) {
        const res = this.validateOption(argument);
        if (res.isErr()) {
          return err(res.error);
        }
      }
    }
    return ok(undefined);
  }

  /**
   * validate option value
   */
  validateOption(
    option: CLICommandOption
  ): Result<undefined, InputValidationError | MissingRequiredInputError> {
    if (option.required && option.default === undefined && option.value === undefined) {
      return err(new MissingRequiredInputError(helper.formatOptionName(option, false), cliSource));
    }
    if (
      (option.type === "singleSelect" || option.type === "multiSelect") &&
      option.choices &&
      option.value !== undefined
    ) {
      if (option.type === "singleSelect") {
        if (!(option.choices as string[]).includes(option.value as string)) {
          return err(
            new InputValidationError(
              helper.formatOptionName(option, false),
              format(
                strings["error.InvalidOptionErrorReason"],
                option.value,
                option.choices.map((i) => JSON.stringify(i)).join(", ")
              )
            )
          );
        }
      } else {
        const values = option.value as string[];
        for (const v of values) {
          if (!(option.choices as string[]).includes(v)) {
            return err(
              new InputValidationError(
                helper.formatOptionName(option, false),
                format(
                  strings["error.InvalidOptionErrorReason"],
                  option.value,
                  option.choices.join(",")
                )
              )
            );
          }
        }
      }
    }
    return ok(undefined);
  }
  processResult(context: CLIContext, fxError?: FxError): void {
    if (context.command.telemetry) {
      if (context.optionValues.env) {
        context.telemetryProperties[TelemetryProperty.Env] = getHashedEnv(
          context.optionValues.env as string
        );
      }
      if (fxError) {
        CliTelemetry.sendTelemetryErrorEvent(
          context.command.telemetry.event,
          fxError,
          context.telemetryProperties
        );
      } else {
        CliTelemetry.sendTelemetryEvent(
          context.command.telemetry.event,
          context.telemetryProperties
        );
      }
    }
    if (fxError) {
      if (isUserCancelError(fxError)) {
        logger.info("User canceled.");
        return;
      }
      logger.outputError(`${fxError.source}.${fxError.name}: ${fxError.message}`);
      if ("helpLink" in fxError && fxError["helpLink"]) {
        logger.outputError(
          `Get help from `,
          colorize(fxError["helpLink"] as string, TextType.Hyperlink)
        );
      }
      if ("issueLink" in fxError && fxError["issueLink"]) {
        logger.outputError(
          `Report this issue at `,
          colorize(fxError["issueLink"] as string, TextType.Hyperlink)
        );
      }
    }
  }
}

export const engine = new CLIEngine();
