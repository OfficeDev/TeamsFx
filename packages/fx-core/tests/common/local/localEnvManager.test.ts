// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";

import { UserError, Result, ok } from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import * as path from "path";

import { LocalEnvManager } from "../../../src/common/local/localEnvManager";
import { DepsInfo, DepsType } from "../../../src/common/deps-checker/depsChecker";
import sinon from "sinon";
import { DotnetChecker } from "../../../src/common/deps-checker/internal/dotnetChecker";
import { NgrokChecker } from "../../../src/common/deps-checker/internal/ngrokChecker";
import { FuncToolChecker } from "../../../src/common/deps-checker/internal/funcToolChecker";
import { DepsCheckerError } from "../../../src/common/deps-checker/depsError";
import mockedEnv, { RestoreFn } from "mocked-env";
import { environmentManager } from "../../../src";

chai.use(chaiAsPromised);

describe("LocalEnvManager", () => {
  const projectSettings0 = {
    appName: "unit-test0",
    projectId: "11111111-1111-1111-1111-111111111111",
    version: "2.0.0",
    programmingLanguage: "javascript",
  };
  const localSettings0 = {
    teamsApp: {
      tenantId: "22222222-2222-2222-2222-222222222222",
      teamsAppId: "33333333-3333-3333-3333-333333333333",
    },
    auth: {
      clientId: "44444444-4444-4444-4444-444444444444",
      // encrypted text: "password-placeholder"
      clientSecret:
        "crypto_025d3c0a85c31e192ff0d8b8d0c9f44e3d5044fa95e642ce6c46d8ee5c4e2ad6b90c3ab385589e7c0d52862898efea47433586f4a14c9f899a7769b3ec73eff372161bbe4b98eb8ba928d58a4ad942bfc880585fe0de737c2f3e5d1a0509e844a4adaf55fa8dd0ecd1e6b3f52dc9812cf6bebb0e",
    },
    frontend: {
      tabDomain: "localhost",
      tabEndpoint: "https://localhost:53000",
    },
  };
  const projectPath = path.resolve(__dirname, "data");
  const configFolder = path.resolve(projectPath, ".fx/configs");
  const localEnvManager = new LocalEnvManager();

  beforeEach(() => {
    fs.ensureDirSync(path.resolve(__dirname, "data"));
  });

  describe("getProjectSettings()", () => {
    it("happy path", async () => {
      await fs.ensureDir(configFolder);
      await fs.writeFile(
        path.resolve(configFolder, "projectSettings.json"),
        JSON.stringify(projectSettings0)
      );

      const projectSettings = await localEnvManager.getProjectSettings(projectPath);

      chai.assert.isDefined(projectSettings);
      chai.assert.equal(projectSettings.appName, "unit-test0");
      chai.assert.equal(projectSettings.projectId, "11111111-1111-1111-1111-111111111111");
      chai.assert.equal(projectSettings.version, "2.0.0");
      chai.assert.equal(projectSettings.programmingLanguage, "javascript");
    });

    it("missing field", async () => {
      await fs.ensureDir(configFolder);
      await fs.writeFile(path.resolve(configFolder, "projectSettings.json"), "{}");

      const projectSettings = await localEnvManager.getProjectSettings(projectPath);

      chai.assert.isDefined(projectSettings);
      chai.assert.isUndefined(projectSettings.appName);
      chai.assert.isDefined(projectSettings.projectId);
    });

    it("missing file", async () => {
      await fs.ensureDir(configFolder);
      await fs.emptyDir(configFolder);

      let error: UserError | undefined = undefined;
      try {
        await localEnvManager.getProjectSettings(projectPath);
      } catch (e: any) {
        error = e as UserError;
      }

      chai.assert.isDefined(error);
      chai.assert.equal(error!.name, "FileNotFoundError");
    });
  });

  describe("getLocalSettings()", () => {
    it("happy path", async () => {
      await fs.ensureDir(configFolder);
      await fs.writeFile(
        path.resolve(configFolder, "projectSettings.json"),
        JSON.stringify(projectSettings0)
      );
      await fs.writeFile(
        path.resolve(configFolder, "localSettings.json"),
        JSON.stringify(localSettings0)
      );

      const projectSettings = await localEnvManager.getProjectSettings(projectPath);
      const localSettings = await localEnvManager.getLocalSettings(projectPath, {
        projectId: projectSettings.projectId,
      });

      chai.assert.isDefined(localSettings);
      chai.assert.isDefined(localSettings!.teamsApp);
      chai.assert.equal(localSettings!.teamsApp.tenantId, "22222222-2222-2222-2222-222222222222");
      chai.assert.equal(localSettings!.teamsApp.teamsAppId, "33333333-3333-3333-3333-333333333333");
      chai.assert.isDefined(localSettings!.auth);
      chai.assert.equal(localSettings!.auth.clientId, "44444444-4444-4444-4444-444444444444");
      chai.assert.equal(localSettings!.auth.clientSecret, "password-placeholder");
      chai.assert.isDefined(localSettings!.frontend);
      chai.assert.equal(localSettings!.frontend.tabDomain, "localhost");
      chai.assert.equal(localSettings!.frontend.tabEndpoint, "https://localhost:53000");
    });

    it("missing field", async () => {
      await fs.ensureDir(configFolder);
      await fs.writeFile(
        path.resolve(configFolder, "projectSettings.json"),
        JSON.stringify(projectSettings0)
      );
      await fs.writeFile(path.resolve(configFolder, "localSettings.json"), "{}");

      const projectSettings = await localEnvManager.getProjectSettings(projectPath);
      const localSettings = await localEnvManager.getLocalSettings(projectPath, {
        projectId: projectSettings.projectId,
      });

      chai.assert.isDefined(localSettings);
      chai.assert.isUndefined(localSettings!.teamsApp);
    });

    it("missing file", async () => {
      await fs.ensureDir(configFolder);
      await fs.emptyDir(configFolder);
      await fs.writeFile(
        path.resolve(configFolder, "projectSettings.json"),
        JSON.stringify(projectSettings0)
      );

      const projectSettings = await localEnvManager.getProjectSettings(projectPath);
      const localSettings = await localEnvManager.getLocalSettings(projectPath, {
        projectId: projectSettings.projectId,
      });

      chai.assert.isUndefined(localSettings);
    });
  });

  const testData: {
    message: string;
    solutionSettings: any;
    pluginSettings?: any;
    depsTypes: DepsType[];
  }[] = [
    {
      message: "tab",
      solutionSettings: {
        name: "fx-solution-azure",
        hostType: "Azure",
        capabilities: ["Tab"],
        activeResourcePlugins: ["fx-resource-simple-auth"],
      },
      depsTypes: [DepsType.AzureNode, DepsType.Dotnet],
    },
    {
      message: "tab without simple auth",
      solutionSettings: {
        name: "fx-solution-azure",
        hostType: "Azure",
        capabilities: ["Tab"],
      },
      depsTypes: [DepsType.AzureNode],
    },
    {
      message: "tab + function",
      solutionSettings: {
        name: "fx-solution-azure",
        hostType: "Azure",
        capabilities: ["Tab"],
        azureResources: ["function"],
        activeResourcePlugins: ["fx-resource-simple-auth"],
      },
      depsTypes: [DepsType.FunctionNode, DepsType.Dotnet, DepsType.FuncCoreTools],
    },
    {
      message: "bot",
      solutionSettings: {
        name: "fx-solution-azure",
        hostType: "Azure",
        capabilities: ["Bot"],
      },
      depsTypes: [DepsType.AzureNode, DepsType.Ngrok],
    },
    {
      message: "tab + bot",
      solutionSettings: {
        name: "fx-solution-azure",
        hostType: "Azure",
        capabilities: ["Tab", "Bot"],
        activeResourcePlugins: ["fx-resource-simple-auth"],
      },
      depsTypes: [DepsType.AzureNode, DepsType.Dotnet, DepsType.Ngrok],
    },
    {
      message: "tab + bot + function",
      solutionSettings: {
        name: "fx-solution-azure",
        hostType: "Azure",
        capabilities: ["Tab", "Bot", "MessagingExtension"],
        azureResources: ["function"],
        activeResourcePlugins: ["fx-resource-simple-auth"],
      },
      depsTypes: [DepsType.FunctionNode, DepsType.Dotnet, DepsType.Ngrok, DepsType.FuncCoreTools],
    },
    {
      message: "spfx",
      solutionSettings: {
        name: "fx-solution-azure",
        hostType: "SPFx",
      },
      depsTypes: [DepsType.SpfxNode],
    },
    {
      message: "azure functions hosted bot",
      solutionSettings: {
        name: "fx-solution-azure",
        hostType: "Azure",
        capabilities: ["Bot"],
      },
      pluginSettings: {
        "fx-resource-bot": {
          "host-type": "azure-functions",
        },
      },
      depsTypes: [DepsType.FunctionNode, DepsType.FuncCoreTools, DepsType.Ngrok],
    },
    {
      message: "app service hosted bot",
      solutionSettings: {
        name: "fx-solution-azure",
        hostType: "Azure",
        capabilities: ["Bot"],
      },
      pluginSettings: {
        "fx-resource-bot": {
          "host-type": "app-service",
        },
      },
      depsTypes: [DepsType.AzureNode, DepsType.Ngrok],
    },
  ];

  describe("getActiveDependencies()", () => {
    const sandbox = sinon.createSandbox();
    afterEach(() => {
      sandbox.restore();
    });

    testData.forEach((data) => {
      it(data.message, async () => {
        const projectSettings = {
          appName: "",
          projectId: "",
          solutionSettings: data.solutionSettings,
          pluginSettings: data.pluginSettings,
        };
        const result = localEnvManager.getActiveDependencies(projectSettings);
        chai.assert.sameDeepMembers(data.depsTypes, result);
      });
    });
  });

  describe("getLocalEnvInfo()", () => {
    const sandbox = sinon.createSandbox();
    let mockedEnvRestore: RestoreFn;

    beforeEach(() => {
      sandbox.restore();
      mockedEnvRestore = mockedEnv({ TEAMSFX_APIV3: "true" });
    });

    afterEach(() => {
      sandbox.restore();
      mockedEnvRestore();
    });

    it("getLocalEnvInfo() happy path", async () => {
      sandbox.stub(environmentManager, "loadEnvInfo").resolves(
        ok({
          envName: "local",
          config: {},
          state: { solution: { key: "value" } },
        })
      );
      const res = await localEnvManager.getLocalEnvInfo(projectPath, { projectId: "123" });
      chai.assert.isDefined(res);
      chai.assert.deepEqual(res, {
        envName: "local",
        config: {},
        state: { solution: { key: "value" } },
      });
    });
  });
});
