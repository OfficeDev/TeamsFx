// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  ConfigFolderName,
  ConfigMap,
  FxError,
  InputConfigsFolderName,
  Inputs,
  ok,
  Platform,
  ProjectSettingsFileName,
  Result,
  v2,
  v3,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import * as dotenv from "dotenv";
import fs from "fs-extra";
import "mocha";
import * as os from "os";
import * as path from "path";
import sinon from "sinon";
import Container from "typedi";
import {
  environmentManager,
  newEnvInfo,
  newEnvInfoV3,
  separateSecretDataV3,
  setTools,
} from "../../../src";
import { LocalCrypto } from "../../../src/core/crypto";
import {
  ContextInjectorMW,
  EnvInfoLoaderMW,
  EnvInfoWriterMW,
  ErrorHandlerMW,
  newSolutionContext,
  ProjectSettingsLoaderMW,
  ProjectSettingsWriterMW,
} from "../../../src/core/middleware";
import { CoreHookContext } from "../../../src/core/types";
import { MockProjectSettings, MockTools, randomAppName } from "../utils";

describe("Middleware - EnvInfoWriterMW, EnvInfoLoaderMW", async () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("successfully write EnvInfo and load it with encrypting and decrypting userdata", async () => {
    const appName = randomAppName();
    const inputs: Inputs = { platform: Platform.VSCode };
    const projectPath = path.join(os.tmpdir(), appName);
    inputs.projectPath = projectPath;

    const projectSettings = MockProjectSettings(appName);
    const cryptoProvider = new LocalCrypto(projectSettings.projectId);

    const tools = new MockTools();
    setTools(tools);
    tools.cryptoProvider = cryptoProvider;

    const solutionContext = await newSolutionContext(tools, inputs);
    const configMap = new ConfigMap();
    const pluginName = "fx-resource-aad-app-for-teams";
    const secretName = "clientSecret";
    const secretText = "test";
    configMap.set(secretName, secretText);
    solutionContext.envInfo.state.set("solution", new ConfigMap());
    solutionContext.envInfo.state.set(pluginName, configMap);

    solutionContext.projectSettings = projectSettings;
    solutionContext.cryptoProvider = cryptoProvider;

    const contextV2: v2.Context = {
      userInteraction: tools.ui,
      logProvider: tools.logProvider,
      telemetryReporter: tools.telemetryReporter!,
      cryptoProvider: cryptoProvider,
      permissionRequestProvider: tools.permissionRequestProvider,
      projectSetting: projectSettings,
    };
    const envInfoV1 = newEnvInfo();
    envInfoV1.state = solutionContext.envInfo.state;
    const envInfoV2: v2.EnvInfoV2 = {
      envName: envInfoV1.envName,
      config: envInfoV1.config,
      state: {},
    };
    for (const key of envInfoV1.state.keys()) {
      const map = envInfoV1.state.get(key) as ConfigMap;
      const value = map?.toJSON();
      if (value) {
        envInfoV2.state[key] = value;
      }
    }
    class MyClass {
      tools = tools;
      async setEnvInfoIntoContext(
        inputs: Inputs,
        ctx?: CoreHookContext
      ): Promise<Result<any, FxError>> {
        if (ctx) {
          ctx.projectSettings = projectSettings;
          ctx.contextV2 = contextV2;
          ctx.envInfoV2 = envInfoV2;
        }
        return ok("");
      }
      async getContext(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
        return ok(ctx);
      }
    }

    hooks(MyClass, {
      setEnvInfoIntoContext: [
        ErrorHandlerMW,
        ContextInjectorMW,
        ProjectSettingsWriterMW,
        EnvInfoWriterMW(),
      ],
      getContext: [
        ErrorHandlerMW,
        ProjectSettingsLoaderMW,
        EnvInfoLoaderMW(false),
        ContextInjectorMW,
      ],
    });
    const fileMap = new Map<string, any>();
    sandbox.stub<any, any>(fs, "writeFile").callsFake(async (file: string, data: any) => {
      fileMap.set(file, data);
    });
    sandbox.stub(fs, "pathExists").resolves(true);
    const envName = environmentManager.getDefaultEnvName();
    const envConfigFile = environmentManager.getEnvConfigPath(envName, projectPath);
    const envFiles = environmentManager.getEnvStateFilesPath(envName, projectPath);
    const userdataFile = envFiles.userDataFile;
    const envJsonFile = envFiles.envState;
    const confFolderPath = path.resolve(projectPath, `.${ConfigFolderName}`);
    const settingsFiles = [
      path.resolve(confFolderPath, "settings.json"),
      path.resolve(confFolderPath, InputConfigsFolderName, ProjectSettingsFileName),
    ];
    const my = new MyClass();
    const setRes = await my.setEnvInfoIntoContext(inputs);
    assert.isTrue(setRes.isOk());
    const content = fileMap.get(userdataFile);
    assert.isTrue(content !== undefined);
    const userdata = dotenv.parse(content);
    const secretValue = userdata[`${pluginName}.${secretName}`];
    assert.isTrue(secretValue.startsWith("crypto_"));
    const decryptedRes = cryptoProvider.decrypt(secretValue);
    assert.isTrue(decryptedRes.isOk() && decryptedRes.value === secretText);
    sandbox.stub<any, any>(fs, "readJson").callsFake(async (file: string) => {
      if (envJsonFile === file) return JSON.parse(fileMap.get(envJsonFile));
      if (settingsFiles.includes(file)) return JSON.parse(fileMap.get(file));
      if (envConfigFile === file) return envInfoV1.config;
      return {};
    });
    sandbox.stub<any, any>(fs, "readFile").callsFake(async (file: string) => {
      if (userdataFile === file) return content;
      if (envJsonFile === file) return fileMap.get(envJsonFile);
      if (envConfigFile === file) return JSON.stringify(envInfoV1.config);
      return {};
    });
    const configsFolder = environmentManager.getEnvConfigsFolder(projectPath);
    sandbox.stub<any, any>(fs, "readdir").callsFake(async (file: string) => {
      if (configsFolder === file) return [`config.${envName}.json`];
      return [];
    });
    inputs.env = envName;
    const getRes = await my.getContext(inputs);
    assert.isTrue(getRes.isOk());
    if (getRes.isOk()) {
      const ctx: CoreHookContext = getRes.value as CoreHookContext;
      assert.isTrue(ctx !== undefined);
      if (ctx) {
        assert.isTrue(
          ctx.envInfoV2 &&
            ctx.envInfoV2.state &&
            ctx.envInfoV2.state[pluginName][secretName] === secretText
        );
      }
    }
  });

  it("newEnvInfoV3()", async () => {
    newEnvInfoV3();
  });

  it("separateSecretDataV3()", async () => {
    const data: v3.ResourceStates = {
      solution: {},
      r1: {
        field1: "123456",
        secretFields: ["field1"],
      },
      r2: {
        field2: "789012",
        secretFields: ["field2"],
      },
    };
    Container.set("r1", { secretKeys: ["field1"] });
    Container.set("r2", { secretKeys: ["field2"] });
    const secret = separateSecretDataV3(data);
    assert.deepEqual(secret, {
      "r1.field1": "123456",
      "r2.field2": "789012",
    });
    assert.deepEqual(data, {
      solution: {},
      r1: {
        field1: "{{r1.field1}}",
      },
      r2: {
        field2: "{{r2.field2}}",
      },
    });
  });
});
