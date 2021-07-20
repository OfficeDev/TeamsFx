// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import path from "path";
import fs from "fs-extra";
import yargs from "yargs";

import { commands } from "./cmds";
import * as constants from "./constants";
import { registerPrompts } from "./prompts";
import { HelpParamGenerator } from "./helpParamGenerator";

/**
 * Registers cli and partner commands with yargs.
 * @param yargs
 */
function register(yargs: yargs.Argv): void {
  commands.forEach((command) => {
    yargs.command(
      command.command,
      command.description,
      command.builder.bind(command),
      command.handler.bind(command)
    );
  });
}

/**
 * Shows in `teamsfx -v`.
 * @returns the version of teamsfx-cli.
 */
function getVersion(): string {
  const pkgPath = path.resolve(__dirname, "..", "package.json");
  const pkgContent = fs.readJsonSync(pkgPath);
  return pkgContent.version;
}

/**
 * Starts the CLI process.
 */
export async function start() {
  registerPrompts();
  await HelpParamGenerator.initializeQuestionsForHelp();
  register(yargs);
  yargs
    .options("verbose", {
      description: "Print additional information.",
      boolean: true,
      default: false,
    })
    .options("debug", {
      description: "Print diagnostic information.",
      boolean: true,
      default: false,
    })
    .demandCommand()
    .scriptName(constants.cliName)
    .help()
    .strict()
    .showHelpOnFail(false, "Specify --help for available options")
    .alias("help", "h")
    .alias("v", "version")
    .version(getVersion())
    .wrap(Math.min(100, yargs.terminalWidth()))
    .epilogue("For more information about the Teams Toolkit - https://aka.ms/teamsfx-learn.").argv;
}
