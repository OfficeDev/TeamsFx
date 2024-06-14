// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/****************************************************************************************
 *                            NOTICE: AUTO-GENERATED                                    *
 ****************************************************************************************
 * This file is automatically generated by script "./src/question/generator.ts".        *
 * Please don't manually change its contents, as any modifications will be overwritten! *
 ***************************************************************************************/

import { CLICommandOption, CLICommandArgument } from "@microsoft/teamsfx-api";

export const UninstallOptions: CLICommandOption[] = [
  {
    name: "uninstall-mode",
    type: "string",
    description: "Choose uninstall mode",
    required: true,
    default: "uninstall-env",
    choices: ["uninstall-mode-manifest-id", "uninstall-mode-env", "uninstall-mode-title-id"],
  },
  {
    name: "manifest-id",
    type: "string",
    description: "Manifest ID",
  },
  {
    name: "env",
    type: "string",
    description: "Env",
  },
  {
    name: "projectPath",
    type: "string",
    description: "Project Path for uninstall",
    default: "./",
  },
  {
    name: "uninstall-option",
    type: "array",
    description: "Choose resources to uninstall",
    choices: [
      "uninstall-option-m365-app",
      "uninstall-option-app-registration",
      "uninstall-option-bot-framework-registration",
    ],
  },
  {
    name: "titile-id",
    type: "string",
    description: "Title ID",
  },
];
export const UninstallArguments: CLICommandArgument[] = [];
