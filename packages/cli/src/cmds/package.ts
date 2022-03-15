// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { Argv, Options } from "yargs";
import * as path from "path";
import { FxError, err, ok, Result, Func, Stage, Inputs } from "@microsoft/teamsfx-api";
import activate from "../activate";
import { YargsCommand } from "../yargsCommand";
import { getSystemInputs, askTargetEnvironment } from "../utils";
import CliTelemetry, { makeEnvRelatedProperty } from "../telemetry/cliTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/cliTelemetryEvents";
import HelpParamGenerator from "../helpParamGenerator";
import { environmentManager, isConfigUnifyEnabled } from "@microsoft/teamsfx-core";

export default class Package extends YargsCommand {
  public readonly commandHead = `package`;
  public readonly command = `${this.commandHead}`;
  public readonly description = "Build your Teams app into a package for publishing.";

  public builder(yargs: Argv): Argv<any> {
    this.params = HelpParamGenerator.getYargsParamForHelp(Stage.build);
    return yargs.version(false).options(this.params);
  }

  public async runCommand(args: {
    [argName: string]: string | string[];
  }): Promise<Result<null, FxError>> {
    const rootFolder = path.resolve((args.folder as string) || "./");
    CliTelemetry.withRootFolder(rootFolder).sendTelemetryEvent(TelemetryEvent.BuildStart);

    const result = await activate(rootFolder);
    if (result.isErr()) {
      CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.Build, result.error);
      return err(result.error);
    }
    const core = result.value;
    let inputs: Inputs;
    {
      const func: Func = {
        namespace: "fx-solution-azure",
        method: "buildPackage",
        params: {},
      };

      if (!args.env) {
        // include local env in interactive question
        const selectedEnv = await askTargetEnvironment(rootFolder);
        if (selectedEnv.isErr()) {
          CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.Build, selectedEnv.error);
          return err(selectedEnv.error);
        }
        args.env = selectedEnv.value;
      }

      if (args.env === environmentManager.getLocalEnvName()) {
        func.params.type = "localDebug";
        inputs = getSystemInputs(rootFolder);
        if (isConfigUnifyEnabled()) {
          inputs.ignoreEnvInfo = false;
          inputs.env = args.env;
        } else {
          inputs.ignoreEnvInfo = true;
        }
      } else {
        func.params.type = "remote";
        inputs = getSystemInputs(rootFolder, args.env as any);
        inputs.ignoreEnvInfo = false;
      }

      const result = await core.executeUserTask!(func, inputs);
      if (result.isErr()) {
        CliTelemetry.sendTelemetryErrorEvent(
          TelemetryEvent.Build,
          result.error,
          makeEnvRelatedProperty(rootFolder, inputs)
        );

        return err(result.error);
      }
    }

    CliTelemetry.sendTelemetryEvent(TelemetryEvent.Build, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      ...makeEnvRelatedProperty(rootFolder, inputs),
    });
    return ok(null);
  }
}
