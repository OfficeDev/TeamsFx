/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable @typescript-eslint/no-var-requires */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as chai from "chai";
import * as spies from "chai-spies";
import * as fs from "fs-extra";
import * as sinon from "sinon";
import { Uri } from "vscode";

import { Stage, UserError } from "@microsoft/teamsfx-api";

import * as globalVariables from "../../src/globalVariables";
import {
  TelemetryEvent,
  TelemetryEventCache,
  TelemetryProperty,
} from "../../src/telemetry/extTelemetryEvents";
import * as ExtTelemetry from "../../src/telemetry/telemetry";
import * as commonUtils from "../../src/utils/commonUtils";
import { getAllFeatureFlags } from "../../src/utils/commonUtils";

chai.use(spies);
const expect = chai.expect;
const spy = chai.spy;

const reporterSpy = spy.interface({
  sendTelemetryEvent(
    eventName: string,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {},
  sendTelemetryErrorEvent(
    eventName: string,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {},
  sendTelemetryException(
    error: Error,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {},
  dispose(): Promise<void> {
    return Promise.resolve();
  },
});
const vscReporterSpy = spy.interface({
  sendTelemetryEvent(
    eventName: string,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {},
  sendTelemetryErrorEvent(
    eventName: string,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {},
  sendTelemetryException(
    error: Error,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {},
  dispose(): Promise<void> {
    return Promise.resolve();
  },
});
const cacheSpy = spy.interface({
  addEvent(event: TelemetryEventCache): void {},
  sendEventsInCache(): void {},
  persistUncertainEventsToDiskAsync(event: TelemetryEventCache): Promise<void> {
    return Promise.resolve();
  },
});

const mock = require("mock-require");
mock("@vscode/extension-telemetry", {
  default: function (
    extensionId: string,
    extensionVersion: string,
    key: string,
    firstParty?: boolean
  ) {
    return reporterSpy;
  },
});

const featureFlags = getAllFeatureFlags()?.join(";") ?? "";

describe("telemetry", () => {
  describe("stageToEvent", () => {
    it("Stage.create", () => {
      const stage = Stage.create;
      chai.expect(ExtTelemetry.stageToEvent(stage)).equals(TelemetryEvent.CreateProject);
    });

    it("Stage.provision", () => {
      const stage = Stage.provision;
      chai.expect(ExtTelemetry.stageToEvent(stage)).equals(TelemetryEvent.Provision);
    });

    it("Stage.deploy", () => {
      const stage = Stage.deploy;
      chai.expect(ExtTelemetry.stageToEvent(stage)).equals(TelemetryEvent.Deploy);
    });

    it("Stage.publish", () => {
      const stage = Stage.publish;
      chai.expect(ExtTelemetry.stageToEvent(stage)).equals(TelemetryEvent.Publish);
    });

    it("Stage.creatEnv", () => {
      const stage = Stage.createEnv;
      chai.expect(ExtTelemetry.stageToEvent(stage)).equals(TelemetryEvent.CreateNewEnvironment);
    });

    it("unknown", () => {
      const stage = "unknown";
      chai.expect(ExtTelemetry.stageToEvent(stage as Stage)).equals(undefined);
    });
  });

  describe("Send Telemetry Functions", () => {
    const sandbox = sinon.createSandbox();
    before(() => {
      chai.util.addProperty(ExtTelemetry, "reporter", () => vscReporterSpy);
      sandbox.stub(fs, "pathExistsSync").returns(false);
      sandbox.stub(globalVariables, "workspaceUri").value(Uri.file("test"));
      sandbox.stub(globalVariables, "isSPFxProject").value(false);
      sandbox.stub(globalVariables, "isExistingUser").value("no");
    });

    after(() => {
      sandbox.restore();
    });

    it("sendTelemetryEvent", () => {
      ExtTelemetry.sendTelemetryEvent(
        "sampleEvent",
        { stringProp: "some string" },
        { numericMeasure: 123 }
      );

      chai.expect(vscReporterSpy.sendTelemetryEvent).to.have.been.called.with(
        "sampleEvent",
        {
          stringProp: "some string",
          component: "extension",
          "is-existing-user": "no",
          "is-spfx": "false",
        },
        { numericMeasure: 123 }
      );
    });

    it("sendTelemetryErrorEvent", () => {
      const error = new UserError(
        "test",
        "UserTestError",
        "test error message",
        "displayed test error message"
      );
      ExtTelemetry.sendTelemetryErrorEvent(
        "sampleEvent",
        error,
        { stringProp: "some string" },
        { numericMeasure: 123 },
        ["errorProps"]
      );

      chai.expect(vscReporterSpy.sendTelemetryErrorEvent).to.have.been.called.with(
        "sampleEvent",
        {
          stringProp: "some string",
          component: "extension",
          success: "no",
          "is-existing-user": "no",
          "is-spfx": "false",
          "error-type": "user",
          "error-message": `${error.message}${error.stack ? "\nstack:\n" + error.stack : ""}`,
          "error-code": "test.UserTestError",
        },
        { numericMeasure: 123 },
        ["errorProps"]
      );
    });

    it("sendTelemetryException", () => {
      const error = new UserError("test", "UserTestError", "test error message");
      ExtTelemetry.sendTelemetryException(
        error,
        { stringProp: "some string" },
        { numericMeasure: 123 }
      );

      chai.expect(vscReporterSpy.sendTelemetryException).to.have.been.called.with(
        error,
        {
          stringProp: "some string",
          component: "extension",
          "is-existing-user": "no",
          "is-spfx": "false",
        },
        { numericMeasure: 123 }
      );
    });

    it("dispose", async () => {
      await ExtTelemetry.dispose();

      chai.expect(vscReporterSpy.dispose).to.have.been.called();
    });
  });

  describe("VSCodeTelemetryReporter", () => {
    const sandbox = sinon.createSandbox();
    let tester: ExtTelemetry.VSCodeTelemetryReporter;

    beforeEach(() => {
      tester = new ExtTelemetry.VSCodeTelemetryReporter("test", "1.0.0-rc.1", "test");
      (tester as ExtTelemetry.VSCodeTelemetryReporter).addSharedProperty("project-id", "");
      (tester as ExtTelemetry.VSCodeTelemetryReporter).addSharedProperty(
        "programming-language",
        ""
      );
      (tester as ExtTelemetry.VSCodeTelemetryReporter).addSharedProperty("host-type", "");
      (tester as ExtTelemetry.VSCodeTelemetryReporter).addSharedProperty("is-from-sample", "");

      chai.util.addProperty(tester, "reporter", () => reporterSpy);
      chai.util.addProperty(tester, "cache", () => cacheSpy);
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("sendTelemetryEvent", () => {
      const clock = sandbox.useFakeTimers();
      tester.sendTelemetryEvent(
        "sampleEvent",
        { stringProp: "some string" },
        { numericMeasure: 123 }
      );

      expect(cacheSpy.addEvent).to.have.been.called.with({
        type: "normal",
        eventName: "sampleEvent",
        occurTime: new clock.Date(),
        properties: {
          stringProp: "some string",
          "project-id": "",
          "correlation-id": "",
          "feature-flags": featureFlags,
          "programming-language": "",
          "host-type": "",
          "is-from-sample": "",
        },
        measurements: { numericMeasure: 123 },
      } as TelemetryEventCache);
      clock.restore();
    });

    it("sendTelemetryErrorEvent", () => {
      const clock = sandbox.useFakeTimers();
      tester.sendTelemetryErrorEvent(
        "sampleErrorEvent",
        {
          stringProp: "some string",
          stackProp: "some user stack trace",
        },
        { numericMeasure: 123 },
        ["stackProp"]
      );

      expect(cacheSpy.addEvent).to.have.been.called.with({
        type: "error",
        eventName: "sampleErrorEvent",
        occurTime: new clock.Date(),
        properties: {
          stringProp: "some string",
          stackProp: "some user stack trace",
          "project-id": "",
          "correlation-id": "",
          "feature-flags": featureFlags,
          "programming-language": "",
          "host-type": "",
          "is-from-sample": "",
        },
        measurements: { numericMeasure: 123 },
      } as TelemetryEventCache);
      clock.restore();
    });

    it("sendTelemetryException", () => {
      const error = new Error("error for test");
      tester.sendTelemetryException(error, { stringProp: "some string" }, { numericMeasure: 123 });

      expect(reporterSpy.sendTelemetryException).to.have.been.called.with(
        error,
        {
          stringProp: "some string",
          "project-id": "",
          "correlation-id": "",
          "feature-flags": featureFlags,
          "programming-language": "",
          "host-type": "",
          "is-from-sample": "",
        },
        { numericMeasure: 123 }
      );
    });

    it("dispose", async () => {
      const clock = sandbox.useFakeTimers();
      sandbox.stub(ExtTelemetry, "lastCorrelationId").value("correlation-id");
      sandbox.stub(commonUtils, "getProjectId").returns("project-id");

      await tester.dispose();

      const expectedArgument = {
        type: "normal",
        occurTime: new clock.Date(),
        eventName: TelemetryEvent.Deactivate,
        properties: {
          [TelemetryProperty.CorrelationId]: "correlation-id",
          [TelemetryProperty.ProjectId]: "project-id",
        },
      };
      await clock.tickAsync(1000);
      await clock.nextAsync();

      chai
        .expect(cacheSpy.persistUncertainEventsToDiskAsync)
        .to.have.been.called.with(expectedArgument);

      clock.restore();
    });
  });
});
