// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { Argv } from "yargs";
import { FxError, err, ok, Result, Stage } from "@microsoft/teamsfx-api";
import {} from "@microsoft/teamsfx-core";
import activate from "../activate";
import { YargsCommand } from "../yargsCommand";
import { getSystemInputs } from "../utils";
import CliTelemetry from "../telemetry/cliTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/cliTelemetryEvents";
import * as uuid from "uuid";
import * as path from "path";

export class ApplyCommand extends YargsCommand {
  public readonly commandHead = `apply`;
  public readonly command = this.commandHead;
  public readonly description = "apply a certain template";

  public builder(yargs: Argv): Argv<any> {
    this.params = {
      template: {
        alias: "t",
        describe: "path to yaml template",
        requiresArg: true,
      },
      folder: {
        alias: "f",
        describe: "path to project folder",
        requiresArg: true,
      },
      env: {
        alias: "e",
        describe: "env name",
        requiresArg: true,
      },
      lifecycle: {
        alias: "l",
        describe: "lifecycle to run",
        requiresArg: true,
        choices: ["registerApp", "configureApp", "provision", "deploy", "publish"],
      },
    };
    if (this.params) {
      yargs.options(this.params);
    }
    return yargs.version(false);
  }

  public async runCommand(args: {
    [argName: string]: string | string[];
  }): Promise<Result<null, FxError>> {
    const rootFolder = path.resolve((args.folder as string) || "./");

    const result = await activate(rootFolder);
    if (result.isErr()) {
      return err(result.error);
    }

    const core = result.value;
    const inputs = getSystemInputs(rootFolder);
    inputs.projectId = inputs.projectId ?? uuid.v4();
    inputs.folder = inputs.folder ?? rootFolder;

    const initResult = await core.apply(
      inputs,
      args["template"] as string,
      args["lifecycle"] as string
    );
    if (initResult.isErr()) {
      CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.InitProject, initResult.error);
      return err(initResult.error);
    }

    CliTelemetry.sendTelemetryEvent(TelemetryEvent.InitProject, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      [TelemetryProperty.NewProjectId]: inputs.projectId,
    });
    return ok(null);
  }
}
