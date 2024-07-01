import * as chai from "chai";
import * as spies from "chai-spies";
import { Stage, UserError } from "@microsoft/teamsfx-api";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import * as telemetryModule from "../../src/telemetry/extTelemetry";
import { TelemetryEvent } from "../../src/telemetry/extTelemetryEvents";
import sinon = require("sinon");
import * as vscTelemetryUtils from "../../src/utils/telemetryUtils";
import * as fs from "fs-extra";
import * as globalVariables from "../../src/globalVariables";
import { Uri } from "vscode";
import * as globalState from "@microsoft/teamsfx-core/build/common/globalState";
import { telemetryUtils, maskSecret } from "@microsoft/teamsfx-core";

chai.use(spies);
const spy = chai.spy;

const reporterSpy = spy.interface({
  sendTelemetryErrorEvent(
    eventName: string,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number },
    errorProps?: string[]
  ): void {},
  sendTelemetryEvent(
    eventName: string,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {},
  sendTelemetryException(
    error: Error,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {},
});

describe("ExtTelemetry", () => {
  afterEach(() => {
    // Restore the default sandbox here
    sinon.restore();
  });
  describe("setHasSentTelemetry", () => {
    it("query-expfeature", () => {
      const eventName = "query-expfeature";
      ExtTelemetry.setHasSentTelemetry(eventName);
      chai.expect(ExtTelemetry.hasSentTelemetry).equals(false);
    });

    it("other-event", () => {
      const eventName = "other-event";
      ExtTelemetry.setHasSentTelemetry(eventName);
      chai.expect(ExtTelemetry.hasSentTelemetry).equals(true);
    });
  });

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

    it("Stage.addWebpart", () => {
      const stage = Stage.addWebpart;
      chai.expect(ExtTelemetry.stageToEvent(stage)).equals(TelemetryEvent.AddWebpart);
    });

    it("Stage.copilotPluginAddAPI", () => {
      const stage = Stage.copilotPluginAddAPI;
      chai.expect(ExtTelemetry.stageToEvent(stage)).equals(TelemetryEvent.CopilotPluginAddAPI);
    });

    it("unknown", () => {
      const stage = "unknown";
      chai.expect(ExtTelemetry.stageToEvent(stage as Stage)).equals(undefined);
    });
  });

  describe("Send Telemetry", () => {
    const sandbox = sinon.createSandbox();
    beforeEach(() => {
      chai.util.addProperty(ExtTelemetry, "reporter", () => reporterSpy);
      chai.util.addProperty(ExtTelemetry, "settingsVersion", () => "1.0.0");
      sandbox.stub(fs, "pathExistsSync").returns(false);
      sandbox.stub(globalVariables, "workspaceUri").value(Uri.file("test"));
      sandbox.stub(globalVariables, "isSPFxProject").value(false);
      sandbox.stub(globalVariables, "isExistingUser").value("no");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("sendTelemetryEvent", () => {
      ExtTelemetry.sendTelemetryEvent(
        "sampleEvent",
        { stringProp: "some string" },
        { numericMeasure: 123 }
      );

      chai.expect(reporterSpy.sendTelemetryEvent).to.have.been.called.with(
        "sampleEvent",
        {
          stringProp: "some string",
          component: "extension",
          "is-existing-user": "no",
          "is-spfx": "false",
          "settings-version": "1.0.0",
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

      chai.expect(reporterSpy.sendTelemetryErrorEvent).to.have.been.called.with(
        "sampleEvent",
        {
          stringProp: "some string",
          component: "extension",
          success: "no",
          "is-existing-user": "no",
          "is-spfx": "false",
          "settings-version": "1.0.0",
          "error-type": "user",
          "error-name": "UserTestError",
          "err-message": maskSecret(error.message),
          "err-stack": telemetryUtils.extractMethodNamesFromErrorStack(error.stack),
          "error-code": "test.UserTestError",
          "error-component": "",
          "error-method": "",
          "error-source": "",
          "error-stage": "",
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

      chai.expect(reporterSpy.sendTelemetryException).to.have.been.called.with(
        error,
        {
          stringProp: "some string",
          component: "extension",
          "is-existing-user": "no",
          "is-spfx": "false",
          "settings-version": "1.0.0",
        },
        { numericMeasure: 123 }
      );
    });
  });

  describe("deactivate event", () => {
    const sandbox = sinon.createSandbox();
    afterEach(() => {
      sandbox.restore();
    });
    it("cacheTelemetryEventAsync", async () => {
      const clock = sandbox.useFakeTimers();
      let state = "";
      sandbox.stub(telemetryModule, "lastCorrelationId").value("correlation-id");
      sandbox.stub(vscTelemetryUtils, "getProjectId").resolves("project-id");
      const globalStateUpdateStub = sandbox
        .stub(globalState, "globalStateUpdate")
        .callsFake(async (key, value) => (state = value));
      const eventName = "deactivate";

      await ExtTelemetry.cacheTelemetryEventAsync(eventName);

      sandbox.assert.calledOnce(globalStateUpdateStub);
      const telemetryEvents = {
        eventName: eventName,
        properties: {
          "correlation-id": "correlation-id",
          "project-id": "project-id",
          timestamp: new clock.Date().toISOString(),
        },
      };
      const newValue = JSON.stringify(telemetryEvents);
      chai.expect(state).equals(newValue);
      clock.restore();
    });

    it("sendCachedTelemetryEventsAsync", async () => {
      const timestamp = new Date().toISOString();
      const telemetryEvents = {
        eventName: "deactivate",
        properties: {
          "correlation-id": "correlation-id",
          "project-id": "project-id",
          timestamp: timestamp,
        },
      };
      const telemetryData = JSON.stringify(telemetryEvents);
      sandbox.stub(globalState, "globalStateGet").callsFake(async () => telemetryData);
      sandbox.stub(globalState, "globalStateUpdate");
      chai.util.addProperty(ExtTelemetry, "reporter", () => reporterSpy);

      await ExtTelemetry.sendCachedTelemetryEventsAsync();

      chai.expect(reporterSpy.sendTelemetryEvent).to.have.been.called.with("deactivate", {
        "correlation-id": "correlation-id",
        "project-id": "project-id",
        timestamp: timestamp,
      });
    });
  });
});
