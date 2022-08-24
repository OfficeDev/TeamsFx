// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import "mocha";
import { assert } from "chai";
import sinon from "sinon";
import {
  getQuestionsForAddFeatureSubCommand,
  getQuestionsForAddFeatureV3,
  getQuestionsForAddResourceV3,
  getQuestionsForDeployV3,
} from "../../src/component/questionV3";
import { manifestUtils } from "../../src/component/resource/appManifest/utils";
import {
  Inputs,
  InputsWithProjectPath,
  ok,
  OptionItem,
  Platform,
  SingleSelectQuestion,
  TeamsAppManifest,
} from "@microsoft/teamsfx-api";
import { createContextV3 } from "../../src/component/utils";
import { CicdOptionItem, newEnvInfoV3 } from "../../src";
import { FeatureId } from "../../src/component/questionV3";
import "../../src/component/core";
import * as tools from "../../src/common/tools";

describe("question for v3", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("getQuestionsForDeployV3 - CLI_HELP", async () => {
    const context = createContextV3();
    const envInfo = newEnvInfoV3();
    envInfo.state.solution = {
      provisionSucceeded: true,
    };
    const inputs: InputsWithProjectPath = {
      platform: Platform.CLI_HELP,
      projectPath: ".",
    };
    const res = await getQuestionsForDeployV3(context, inputs);
    assert.isTrue(res.isOk());
  });
  it("getQuestionsForDeployV3 - VS Code", async () => {
    const projectSettings = {
      appName: "hj070701",
      projectId: "112233",
      version: "2.1.0",
      isFromSample: false,
      components: [
        {
          name: "teams-bot",
          hosting: "azure-function",
          capabilities: ["notification"],
          build: true,
          deploy: true,
          provision: true,
          folder: "bot",
        },
        {
          name: "azure-function",
          connections: ["teams-bot"],
        },
        {
          name: "bot-service",
          provision: true,
        },
        {
          name: "apim",
          deploy: true,
          provision: true,
        },
      ],
      programmingLanguage: "javascript",
    };
    const inputs: InputsWithProjectPath = {
      platform: Platform.VSCode,
      projectPath: ".",
    };
    const context = createContextV3(projectSettings);
    const envInfo = newEnvInfoV3();
    envInfo.state.solution = {
      provisionSucceeded: true,
    };
    const res = await getQuestionsForDeployV3(context, inputs, envInfo);
    assert.isTrue(res.isOk());
  });
  it("getQuestionsForAddFeatureV3 - CLI_HELP", async () => {
    const context = createContextV3();
    const inputs: InputsWithProjectPath = {
      platform: Platform.CLI_HELP,
      projectPath: ".",
    };
    const res = await getQuestionsForAddFeatureV3(context, inputs);
    assert.isTrue(res.isOk());
  });
  it("getQuestionsForAddFeatureV3 - VS Code", async () => {
    const manifest = new TeamsAppManifest();
    manifest.staticTabs = [];
    manifest.bots = [];
    manifest.composeExtensions = [];
    sandbox.stub(manifestUtils, "readAppManifest").resolves(ok(manifest));
    const projectSettings = {
      appName: "hj070701",
      projectId: "112233",
      version: "2.1.0",
      isFromSample: false,
      components: [
        {
          name: "teams-bot",
          hosting: "azure-function",
          capabilities: ["notification"],
          build: true,
          folder: "bot",
        },
        {
          name: "azure-function",
          connections: ["teams-bot"],
        },
        {
          name: "bot-service",
          provision: true,
        },
      ],
      programmingLanguage: "javascript",
    };
    const inputs: InputsWithProjectPath = {
      platform: Platform.VSCode,
      projectPath: ".",
    };
    const context = createContextV3(projectSettings);
    const res = await getQuestionsForAddFeatureV3(context, inputs);
    assert.isTrue(res.isOk());
  });

  it("getQuestionsForAddFeatureV3 for SPFx - VS Code", async () => {
    const manifest = new TeamsAppManifest();
    manifest.staticTabs = [];
    manifest.bots = [];
    manifest.composeExtensions = [];
    sandbox.stub(manifestUtils, "readAppManifest").resolves(ok(manifest));
    sandbox.stub<any, any>(tools, "canAddCICDWorkflows").resolves(true);
    const projectSettings = {
      appName: "hj070701",
      projectId: "112233",
      version: "2.1.0",
      isFromSample: false,
      components: [],
      programmingLanguage: "javascript",
      solutionSettings: {
        name: "fx-solution-azure",
        version: "1.0.0",
        hostType: "SPFx",
        azureResources: [],
        capabilities: ["Tab"],
        activeResourcePlugins: [
          "fx-resource-spfx",
          "fx-resource-local-debug",
          "fx-resource-appstudio",
        ],
      },
    };
    const inputs: InputsWithProjectPath = {
      platform: Platform.VSCode,
      projectPath: ".",
    };
    const context = createContextV3(projectSettings);
    const res = await getQuestionsForAddFeatureV3(context, inputs);
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      const node = res.value;
      assert.isTrue(
        node && node.data && node.data.type === "singleSelect",
        "result should be singleSelect"
      );
      if (node && node.data && node.data.type === "singleSelect") {
        const options = (node.data as SingleSelectQuestion).staticOptions as OptionItem[];
        assert.deepEqual(options, [CicdOptionItem], "option item should match");
      }
    }
  });

  it("getQuestionsForAddResourceV3 - CLI_HELP", async () => {
    const context = createContextV3();
    const inputs: InputsWithProjectPath = {
      platform: Platform.CLI_HELP,
      projectPath: ".",
    };
    const res = await getQuestionsForAddResourceV3(context, inputs);
    assert.isTrue(res.isOk());
  });
  it("getQuestionsForAddResourceV3 - VS Code", async () => {
    const projectSettings = {
      appName: "hj070701",
      projectId: "112233",
      version: "2.1.0",
      isFromSample: false,
      components: [
        {
          name: "teams-bot",
          hosting: "azure-function",
          capabilities: ["notification"],
          build: true,
          folder: "bot",
        },
        {
          name: "azure-function",
          connections: ["teams-bot"],
        },
        {
          name: "bot-service",
          provision: true,
        },
      ],
      programmingLanguage: "javascript",
    };
    const inputs: InputsWithProjectPath = {
      platform: Platform.VSCode,
      projectPath: ".",
    };
    const context = createContextV3(projectSettings);
    const res = await getQuestionsForAddResourceV3(context, inputs);
    assert.isTrue(res.isOk());
  });
  it("getQuestionsForAddFeatureSubCommand", async () => {
    const inputs: Inputs = {
      platform: Platform.CLI_HELP,
    };
    for (const feature in FeatureId) {
      const res = await getQuestionsForAddFeatureSubCommand(feature as FeatureId, inputs);
      assert.isTrue(res.isOk());
    }
  });
});
