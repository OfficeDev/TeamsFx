// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, ok } from "@microsoft/teamsfx-api";
import { logger } from "../../commonlib/logger";
import { FooterText } from "../../constants";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { getVersion } from "../../utils";
import { helper } from "../helper";
import { accountCommand } from "./account";
import { addCommand } from "./add";
import { getCreateCommand } from "./create";
import { deployCommand } from "./deploy";
import { entraAppCommand } from "./entraAppUpdate";
import { envCommand } from "./env";
import { listCommand } from "./list";
import { m365LaunchInfoCommand } from "./m365LaunchInfo";
import { m365SideloadingCommand } from "./m365Sideloading";
import { m365UnacquireCommand } from "./m365Unacquire";
import { permissionCommand } from "./permission";
import { previewCommand } from "./preview";
import { provisionCommand } from "./provision";
import { teamsappDoctorCommand } from "./teamsapp/doctor";
import { teamsappPackageCommand } from "./teamsapp/package";
import { teamsappPublishCommand } from "./teamsapp/publish";
import { teamsappUpdateCommand } from "./teamsapp/update";
import { teamsappValidateCommand } from "./teamsapp/validate";
import { upgradeCommand } from "./upgrade";
import { teamsappListCommand } from "./teamsapp/listApps";
import { teamsappDeleteCommand } from "./teamsapp/delete";

export const helpCommand: CLICommand = {
  name: "help",
  description: "Show Microsoft Teams Toolkit CLI help.",
  handler: (ctx) => {
    const helpText = helper.formatHelp(rootCommand, undefined);
    logger.info(helpText);
    return ok(undefined);
  },
};
export const rootCommand: CLICommand = {
  name: "teamsapp",
  fullName: "teamsapp",
  description: "Microsoft Teams Toolkit CLI.",
  version: getVersion(),
  footer: FooterText,
  commands: [
    accountCommand,
    getCreateCommand(),
    addCommand,
    provisionCommand,
    deployCommand,
    previewCommand,
    envCommand,
    permissionCommand,
    upgradeCommand,
    listCommand,
    helpCommand,
    teamsappUpdateCommand,
    teamsappValidateCommand,
    teamsappPackageCommand,
    teamsappPublishCommand,
    teamsappDoctorCommand,
    teamsappListCommand,
    teamsappDeleteCommand,
    entraAppCommand,
    m365SideloadingCommand,
    m365UnacquireCommand,
    m365LaunchInfoCommand,
  ],
  sortCommands: true,
  options: [
    {
      type: "boolean",
      name: "version",
      shortName: "v",
      description: "Display Microsoft Teams Toolkit CLI version.",
    },
    {
      type: "boolean",
      name: "help",
      shortName: "h",
      description: "Show Microsoft Teams Toolkit CLI help.",
    },
    {
      type: "boolean",
      name: "interactive",
      shortName: "i",
      description: "Run the command in interactive mode.",
      default: true,
    },
    {
      type: "boolean",
      name: "debug",
      description: "Print debug information.",
      default: false,
    },
    {
      type: "boolean",
      name: "verbose",
      description: "Print diagnostic information.",
      default: false,
    },
    {
      type: "boolean",
      name: "telemetry",
      description: "Whether to enable telemetry.",
      default: true,
    },
  ],
  telemetry: {
    event: TelemetryEvent.RootCommand,
  },
};
