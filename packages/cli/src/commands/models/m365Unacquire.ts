// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, err, ok } from "@microsoft/teamsfx-api";
import { UninstallInputs, QuestionNames } from "@microsoft/teamsfx-core";
import { logger } from "../../commonlib/logger";
import { MissingRequiredOptionError } from "../../error";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { m365utils, sideloadingServiceEndpoint } from "./m365Sideloading";
import { getFxCore } from "../../activate";

export const m365UnacquireCommand: CLICommand = {
  name: "uninstall",
  aliases: ["unacquire"],
  description: commands.uninstall.description,
  options: [
    {
      name: QuestionNames.UninstallMode,
      description: commands.uninstall.options["uninstall-mode"],
      type: "string",
    },
    {
      name: "title-id",
      description: commands.uninstall.options["title-id"],
      type: "string",
    },
    {
      name: "manifest-id",
      description: commands.uninstall.options["manifest-id"],
      type: "string",
    },
    {
      name: "env",
      description: commands.uninstall.options["env"],
      type: "string",
    },
    {
      name: QuestionNames.UninstallOption,
      description: commands.uninstall.options["uninstall-option"],
      type: "array",
    },
  ],
  examples: [
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} uninstall --title-id U_xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`,
      description: "Remove the acquired M365 App by Title ID",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} uninstall --manifest-id xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx -i false --m365-app --app-refistration --bot-framework-registration`,
      description: "Remove the acquired M365 App by Manifest ID",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} uninstall --env xxx -i false --m365-app --app-refistration --bot-framework-registration`,
      description: "Remove the acquired M365 App by local env",
    },
  ],
  telemetry: {
    event: TelemetryEvent.M365Unacquire,
  },
  defaultInteractiveOption: true,
  handler: async (ctx) => {
    const inputs = ctx.optionValues as UninstallInputs;
    const core = getFxCore();
    const res = await core.uninstall(inputs);
    if (res.isErr()) {
      return err(res.error);
    }
    return ok(undefined);
  },
};
