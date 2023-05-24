// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import "mocha";
import { assert } from "chai";
import {
  convertEnvStateMapV3ToV2,
  convertProjectSettingsV2ToV3,
  convertProjectSettingsV3ToV2,
} from "../../src/component/migrate";
import { InputsWithProjectPath, Platform, ProjectSettingsV3 } from "@microsoft/teamsfx-api";
import * as path from "path";
import * as os from "os";
import mockedEnv, { RestoreFn } from "mocked-env";
describe("Migration test for v3", () => {
  let mockedEnvRestore: RestoreFn;
  it("convertProjectSettingsV2ToV3", async () => {
    const projectSettings = {
      appName: "hj070701",
      projectId: "112233",
      version: "2.1.0",
      isFromSample: false,
      solutionSettings: {
        name: "fx-solution-azure",
        version: "1.0.0",
        hostType: "Azure",
        azureResources: ["function", "apim", "sql", "keyvault"],
        capabilities: ["Bot", "Tab", "TabSSO", "MessagingExtension"],
        activeResourcePlugins: [
          "fx-resource-frontend-hosting",
          "fx-resource-identity",
          "fx-resource-azure-sql",
          "fx-resource-bot",
          "fx-resource-aad-app-for-teams",
          "fx-resource-function",
          "fx-resource-local-debug",
          "fx-resource-apim",
          "fx-resource-appstudio",
          "fx-resource-key-vault",
          "fx-resource-cicd",
          "fx-resource-api-connector",
        ],
      },
      programmingLanguage: "javascript",
      pluginSettings: {
        "fx-resource-bot": {
          "host-type": "azure-functions",
          capabilities: ["notification"],
        },
      },
      defaultFunctionName: "getUserProfile",
    };
    const v3 = convertProjectSettingsV2ToV3(projectSettings, ".");
    assert.isTrue(v3.components.length > 0);
  });
  it("convertProjectSettingsV3ToV2", async () => {
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
        {
          name: "teams-tab",
          hosting: "azure-storage",
          build: true,
          provision: true,
          folder: "tabs",
          connections: ["teams-api"],
        },
        {
          name: "azure-storage",
          connections: ["teams-tab"],
          provision: true,
        },
        {
          name: "apim",
          provision: true,
          deploy: true,
          connections: ["teams-tab", "teams-bot"],
        },
        {
          name: "teams-api",
          hosting: "azure-function",
          functionNames: ["getUserProfile"],
          build: true,
          folder: "api",
        },
        {
          name: "azure-function",
          connections: ["teams-api"],
        },
        {
          name: "simple-auth",
        },
        {
          name: "key-vault",
        },
      ],
      programmingLanguage: "javascript",
    };
    const v2 = convertProjectSettingsV3ToV2(projectSettings);
    assert.isTrue(v2.solutionSettings !== undefined);
  });

  it("convertEnvStateMapV3ToV2", async () => {
    const envStateMap = new Map<string, any>();
    envStateMap.set("app-manifest", new Map<string, any>());
    const res = convertEnvStateMapV3ToV2(envStateMap);
    assert.isTrue(res.has("fx-resource-appstudio"));
  });
});
