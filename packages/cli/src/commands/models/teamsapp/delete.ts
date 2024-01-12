// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, Inputs, err, ok } from "@microsoft/teamsfx-api";
import { getFxCore } from "../../../activate";
import { logger } from "../../../commonlib/logger";
import { TelemetryEvent } from "../../../telemetry/cliTelemetryEvents";
import { AppDefinition } from "@microsoft/teamsfx-core";

export const teamsappDeleteCommand: CLICommand = {
  name: "delete-apps",
  description: "Delete Microsoft Teams apps in Teams Developer Portal.",
  options: [
    {
      type: "string",
      name: "filter",
      description: "Filter apps by name or id.",
    },
  ],
  telemetry: {
    event: TelemetryEvent.DeleteTeamsApp,
  },
  defaultInteractiveOption: false,
  handler: async (ctx) => {
    const inputs = ctx.optionValues as Inputs;
    const core = getFxCore();
    const listRes = await core.listTeamsApps(inputs);
    if (listRes.isOk()) {
      let apps = listRes.value;
      if (inputs.filter) {
        apps = apps.filter(
          (app: AppDefinition) =>
            app.appName?.includes(inputs.filter) || app.teamsAppId?.includes(inputs.filter)
        );
      }
      inputs.teamsAppIds = apps.map((app: AppDefinition) => app.teamsAppId!);
      logger.info(`Found ${apps.length} apps to delete: ${JSON.stringify(inputs.teamsAppIds)}`);
      const res = await core.deleteTeamsApps(inputs);
      if (res.isErr()) {
        return err(res.error);
      }
      return ok(undefined);
    } else {
      return err(listRes.error);
    }
  },
};
