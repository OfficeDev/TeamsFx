// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, Inputs, err, ok } from "@microsoft/teamsfx-api";
import { getFxCore } from "../../../activate";
import { logger } from "../../../commonlib/logger";
import { TelemetryEvent } from "../../../telemetry/cliTelemetryEvents";
import { AppDefinition } from "@microsoft/teamsfx-core";

export const teamsappListCommand: CLICommand = {
  name: "list-apps",
  description: "List Microsoft Teams apps in Teams Developer Portal.",
  options: [
    {
      type: "string",
      name: "filter",
      description: "Filter apps by name or id.",
    },
  ],
  telemetry: {
    event: TelemetryEvent.ListTeamsApps,
  },
  defaultInteractiveOption: false,
  handler: async (ctx) => {
    const inputs = ctx.optionValues as Inputs;
    const core = getFxCore();
    const res = await core.listTeamsApps(inputs);
    if (res.isOk()) {
      let apps = res.value;
      if (inputs.filter) {
        apps = apps.filter(
          (app: AppDefinition) =>
            app.appName?.includes(inputs.filter) || app.teamsAppId?.includes(inputs.filter)
        );
      }
      logger.info(`Found ${apps.length} apps:`);
      logger.info(JSON.stringify(apps, null, 2));
      return ok(undefined);
    } else {
      return err(res.error);
    }
  },
};
