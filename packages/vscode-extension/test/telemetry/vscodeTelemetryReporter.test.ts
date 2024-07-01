/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable @typescript-eslint/no-var-requires */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as chai from "chai";
import * as spies from "chai-spies";
import { TelemetryReporter } from "@microsoft/teamsfx-api";
import { VSCodeTelemetryReporter } from "../../src/telemetry/vscodeTelemetryReporter";
import { getAllFeatureFlags } from "../../src/featureFlags";

chai.use(spies);
const expect = chai.expect;
const spy = chai.spy;

const reporterSpy = spy.interface({
  sendTelemetryErrorEvent(
    eventName: string,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
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
  let tester: TelemetryReporter;

  before(() => {
    tester = new VSCodeTelemetryReporter("test", "1.0.0-rc.1", "test");
    (tester as VSCodeTelemetryReporter).addSharedProperty("project-id", "");
    (tester as VSCodeTelemetryReporter).addSharedProperty("programming-language", "");
    (tester as VSCodeTelemetryReporter).addSharedProperty("host-type", "");
    (tester as VSCodeTelemetryReporter).addSharedProperty("is-from-sample", "");
    chai.util.addProperty(tester, "reporter", () => reporterSpy);
  });

  it("sendTelemetryEvent", () => {
    tester.sendTelemetryEvent(
      "sampleEvent",
      { stringProp: "some string" },
      { numericMeasure: 123 }
    );

    expect(reporterSpy.sendTelemetryEvent).to.have.been.called.with(
      "sampleEvent",
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

  it("sendTelemetryErrorEvent", () => {
    tester.sendTelemetryErrorEvent(
      "sampleErrorEvent",
      {
        stringProp: "some string",
        "error-stack": "some user stack trace at (C:/fake_path/fake_file:1:1)",
      },
      { numericMeasure: 123 },
      ["error-stack"]
    );

    expect(reporterSpy.sendTelemetryErrorEvent).to.have.been.called.with(
      "sampleErrorEvent",
      {
        stringProp: "some string",
        "error-stack": "some user stack trace at (<REDACTED: user-file-path>/fake_file:1:1)",
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
});
