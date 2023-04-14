// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";

import sinon from "sinon";
import yargs, { Options } from "yargs";
import { err, FxError, ok, Void } from "@microsoft/teamsfx-api";
import { FxCore } from "@microsoft/teamsfx-core";
import HelpParamGenerator from "../../../src/helpParamGenerator";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../../../src/telemetry/cliTelemetryEvents";
import CliTelemetry from "../../../src/telemetry/cliTelemetry";
import Upgrade from "../../../src/cmds/upgrade";
import { expect, TestFolder } from "../utils";
import { NonTeamsFxProjectFolder } from "../../../src/error";

describe("Init Command Tests", () => {
  const sandbox = sinon.createSandbox();
  let registeredCommands: string[] = [];
  let options: string[] = [];
  let telemetryEvents: string[] = [];
  let telemetryEventStatus: string | undefined = undefined;

  beforeEach(() => {
    sandbox.stub(HelpParamGenerator, "getYargsParamForHelp").returns({});
    sandbox
      .stub<any, any>(yargs, "command")
      .callsFake((command: string, description: string, builder: any, handler: any) => {
        registeredCommands.push(command);
        builder(yargs);
      });
    sandbox.stub(yargs, "options").callsFake((ops: { [key: string]: Options }) => {
      if (typeof ops === "string") {
        options.push(ops);
      } else {
        options = options.concat(...Object.keys(ops));
      }
      return yargs;
    });
    sandbox.stub(yargs, "exit").callsFake((code: number, err: Error) => {
      throw err;
    });
    sandbox
      .stub(CliTelemetry, "sendTelemetryEvent")
      .callsFake((eventName: string, options?: { [_: string]: string }) => {
        telemetryEvents.push(eventName);
        if (options && TelemetryProperty.Success in options) {
          telemetryEventStatus = options[TelemetryProperty.Success];
        }
      });
    sandbox
      .stub(CliTelemetry, "sendTelemetryErrorEvent")
      .callsFake((eventName: string, error: FxError) => {
        telemetryEvents.push(eventName);
        telemetryEventStatus = TelemetrySuccess.No;
      });
  });

  afterEach(() => {
    sandbox.restore();
  });

  beforeEach(() => {
    registeredCommands = [];
    options = [];
    telemetryEvents = [];
    telemetryEventStatus = undefined;
  });

  it("Builder Check", () => {
    const cmd = new Upgrade();
    yargs.command(cmd.command, cmd.description, cmd.builder.bind(cmd), cmd.handler.bind(cmd));
    expect(registeredCommands).deep.equals(["upgrade"]);
  });

  it("Command Running Check", async () => {
    sandbox.stub(FxCore.prototype, "phantomMigrationV3").callsFake((inputs) => {
      expect(inputs.projectPath).equals(TestFolder);
      expect(inputs.skipUserConfirm).equals(true);
      expect(inputs.nonInteractive).equals(undefined);
      return Promise.resolve(ok(Void));
    });
    const cmd = new Upgrade();
    const args = {
      folder: TestFolder,
      force: true,
    };
    await cmd.runCommand(args as any);
    expect(telemetryEvents).deep.equals([TelemetryEvent.UpgradeStart, TelemetryEvent.Upgrade]);
  });

  it("Command Running Check - error", async () => {
    sandbox.stub(FxCore.prototype, "phantomMigrationV3").callsFake((inputs) => {
      if (inputs.projectPath?.includes("fake"))
        return Promise.resolve(err(NonTeamsFxProjectFolder()));
      return Promise.resolve(ok(Void));
    });

    const cmd = new Upgrade();
    const args = {
      folder: "fake",
    };
    const result = await cmd.runCommand(args);
    expect(result.isErr());
  });
});
