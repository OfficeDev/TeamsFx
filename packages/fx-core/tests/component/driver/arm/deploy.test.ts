// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import "mocha";
import { createSandbox } from "sinon";
import { setTools } from "../../../../src/core/globalVars";
import {
  MockAzureAccountProvider,
  MockTelemetryReporter,
  MockTools,
  MockUserInteraction,
} from "../../../core/utils";
import { MockedM365Provider } from "../../../plugins/solution/util";
import { ArmDeployDriver } from "../../../../src/component/driver/arm/deploy";
import fs from "fs-extra";
import * as cpUtils from "../../../../src/common/cpUtils";
import { ArmDeployImpl } from "../../../../src/component/driver/arm/deployImpl";
import { ok } from "@microsoft/teamsfx-api";
import * as bicepChecker from "../../../../src/component/utils/depsChecker/bicepChecker";
import axios from "axios";

describe("Arm driver deploy", () => {
  const sandbox = createSandbox();
  const tools = new MockTools();
  setTools(tools);
  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    azureAccountProvider: new MockAzureAccountProvider(),
    telemetryReporter: new MockTelemetryReporter(),
    ui: new MockUserInteraction(),
  };
  const driver = new ArmDeployDriver();

  const bicepCliVersion = "v0.9.1";
  beforeEach(() => {});

  afterEach(() => {
    sandbox.restore();
  });

  it("happy path", async () => {
    sandbox.stub(fs, "readFile").resolves("{}" as any);
    sandbox.stub(cpUtils, "executeCommand").resolves("{}" as any);
    const deployRes = ok({
      mockKey: {
        type: "string",
        value: "mockValue",
      },
    });
    sandbox.stub(ArmDeployImpl.prototype, "executeDeployment").resolves(deployRes as any);
    sandbox.stub(bicepChecker, "getAvailableBicepVersions").resolves([bicepCliVersion]);
    const fakeAxiosInstance = axios.create();
    sandbox.stub(axios, "create").returns(fakeAxiosInstance);
    sandbox.stub(fakeAxiosInstance, "get").resolves({
      status: 200,
      data: "",
    });
    let deployArgs = {
      subscriptionId: "00000000-0000-0000-0000-000000000000",
      resourceGroupName: "mock-group",
      bicepCliVersion: bicepCliVersion,
      templates: [
        {
          path: "mock-template.bicep",
          parameters: "mock-parameters.json",
          deploymentName: "mock-deployment",
        },
        {
          path: "mock-template2.json",
          parameters: "mock-parameters2.json",
          deploymentName: "mock-deployment2",
        },
      ],
    };

    let res = await driver.run(deployArgs, mockedDriverContext);
    assert.isTrue(res.isOk());

    deployArgs = {
      subscriptionId: "00000000-0000-0000-0000-000000000000",
      resourceGroupName: "mock-group",
      bicepCliVersion: "",
      templates: [
        {
          path: "mock-template.json",
          parameters: "mock-parameters.json",
          deploymentName: "mock-deployment",
        },
      ],
    };
    res = await driver.run(deployArgs, mockedDriverContext);
    assert.isTrue(res.isOk());
  });

  it("invalid parameters", async () => {
    sandbox.stub(fs, "readFile").resolves("{}" as any);
    sandbox.stub(cpUtils, "executeCommand").resolves("{}" as any);
    let deployArgs = {
      subscriptionId: "",
      resourceGroupName: "",
      bicepCliVersion: "",
      templates: [
        {
          path: "mock-template",
          parameters: "mock-parameters",
          deploymentName: "",
        },
      ],
    } as any;
    let res = await driver.run(deployArgs, mockedDriverContext);
    assert.isTrue(res.isErr());

    deployArgs = {
      subscriptionId: "00000000-0000-0000-0000-000000000000",
      resourceGroupName: "mock-group",
      bicepCliVersion: "",
      templates: [],
    } as any;
    res = await driver.run(deployArgs, mockedDriverContext);
    assert.isTrue(res.isErr());
  });

  it("deploy error", async () => {
    sandbox.stub(fs, "readFile").resolves("{}" as any);
    sandbox.stub(cpUtils, "executeCommand").resolves("{}" as any);
    sandbox
      .stub(ArmDeployImpl.prototype, "executeDeployment")
      .rejects(new Error("mocked deploy error"));
    sandbox.stub(bicepChecker, "getAvailableBicepVersions").resolves([bicepCliVersion]);
    sandbox.stub(ArmDeployImpl.prototype, "ensureBicepCli").resolves();
    const deployArgs = {
      subscriptionId: "00000000-0000-0000-0000-000000000000",
      resourceGroupName: "mock-group",
      bicepCliVersion: bicepCliVersion,
      templates: [
        {
          path: "mock-template.bicep",
          parameters: "mock-parameters.json",
          deploymentName: "mock-deployment",
        },
        {
          path: "mock-template2.json",
          parameters: "mock-parameters2.json",
          deploymentName: "mock-deployment2",
        },
      ],
    };

    const res = await driver.run(deployArgs, mockedDriverContext);
    assert.isTrue(res.isErr());
  });

  it("error handle", async () => {
    sandbox.stub(ArmDeployImpl.prototype, "run").throws("mocked deploy error");

    const res = await driver.run({} as any, mockedDriverContext);
    assert.isTrue(res.isErr());
  });
});
