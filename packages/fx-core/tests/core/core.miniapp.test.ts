// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Func, Inputs, Platform } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import * as os from "os";
import * as path from "path";
import sinon from "sinon";
import { FxCore, setTools } from "../../src";
import { TabSPFxItem } from "../../src/plugins/solution/fx-solution/question";
import { deleteFolder, MockTools, randomAppName } from "./utils";
describe("Core API for mini app", () => {
  const sandbox = sinon.createSandbox();
  const tools = new MockTools();
  let projectPath: string;
  let mockedEnvRestore: RestoreFn;
  beforeEach(() => {
    setTools(tools);
    mockedEnvRestore = mockedEnv({ TEAMSFX_APIV3: "false" });
  });
  afterEach(async () => {
    sandbox.restore();
    deleteFolder(projectPath);
    mockedEnvRestore();
  });
  it("init + add spfx tab", async () => {
    const appName = randomAppName();
    projectPath = path.join(os.tmpdir(), appName);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      folder: projectPath,
      "app-name": appName,
    };
    const core = new FxCore(tools);
    const initRes = await core.init(inputs);
    assert.isTrue(initRes.isOk());
    if (initRes.isOk()) {
      const addInputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: projectPath,
        capabilities: [TabSPFxItem.id],
        "spfx-framework-type": "react",
        "spfx-webpart-name": "helloworld",
        "spfx-webpart-desp": "helloworld",
      };
      const func: Func = {
        namespace: "fx-solution-azure",
        method: "addCapability",
      };
      const addRes = await core.executeUserTaskV2(func, addInputs);
      assert.isTrue(addRes.isOk());
    }
  });
});
