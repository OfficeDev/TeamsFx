// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import * as chai from "chai";
import * as path from "path";
import fs from "fs-extra";
import * as dotenv from "dotenv";
import { expect } from "chai";
import { EnvHandler } from "../../../../src/plugins/resource/apiconnector/envHandler";
import {
  AuthType,
  ComponentType,
  Constants,
} from "../../../../src/plugins/resource/apiconnector/constants";
import {
  ApiConnectorConfiguration,
  BasicAuthConfig,
} from "../../../../src/plugins/resource/apiconnector/config";
import { LocalEnvProvider, LocalEnvs } from "../../../../src/common/local/localEnvProvider";

describe("EnvHandler", () => {
  const fakeProjectPath = path.join(__dirname, "test-api-connector");
  const botPath = path.join(fakeProjectPath, "bot");
  const apiPath = path.join(fakeProjectPath, "api");
  const localEnvFileName = ".env.teamsfx.local";
  beforeEach(async () => {
    await fs.ensureDir(fakeProjectPath);
    await fs.ensureDir(botPath);
    await fs.ensureDir(apiPath);
  });
  afterEach(async () => {
    await fs.remove(fakeProjectPath);
  });

  it("should create .env.teamsfx.local if not exist with empty api envs", async () => {
    const service: ComponentType = ComponentType.BOT;
    const envHandler = new EnvHandler(fakeProjectPath, service);
    expect(await fs.pathExists(path.join(botPath, localEnvFileName))).to.be.false;
    await envHandler.saveLocalEnvFile();
    expect(await fs.pathExists(path.join(botPath, localEnvFileName))).to.be.true;
    const provider: LocalEnvProvider = new LocalEnvProvider(fakeProjectPath);
    const envs: LocalEnvs = await provider.loadBotLocalEnvs();
    for (const item in envs.customizedLocalEnvs) {
      expect(item.startsWith("API_")).to.be.false;
    }
  });

  it("env save to .env.teamsfx.local first time", async () => {
    const service: ComponentType = ComponentType.BOT;
    const envHandler = new EnvHandler(fakeProjectPath, service);
    const fakeConfig: ApiConnectorConfiguration = {
      ComponentPath: ["bot"],
      APIName: "FAKE",
      EndPoint: "fake_endpoint",
      AuthConfig: {
        AuthType: AuthType.BASIC,
        UserName: "fake_api_user_name",
      } as BasicAuthConfig,
    };
    envHandler.updateEnvs(fakeConfig);
    expect(await fs.pathExists(path.join(botPath, localEnvFileName))).to.be.false;
    await envHandler.saveLocalEnvFile();
    const envs = dotenv.parse(await fs.readFile(path.join(botPath, localEnvFileName)));
    chai.assert.strictEqual(envs[Constants.envPrefix + "FAKE_ENDPOINT"], "fake_endpoint");
    chai.assert.strictEqual(envs[Constants.envPrefix + "FAKE_USERNAME"], "fake_api_user_name");
    chai.assert.exists(envs[Constants.envPrefix + "FAKE_PASSWORD"]);
  });

  it("env update in .env.teamsfx.local", async () => {
    const service: ComponentType = ComponentType.BOT;
    const envHandler = new EnvHandler(fakeProjectPath, service);
    const fakeConfig: ApiConnectorConfiguration = {
      ComponentPath: ["bot"],
      APIName: "FAKE",
      EndPoint: "fake_endpoint",
      AuthConfig: {
        AuthType: AuthType.BASIC,
        UserName: "fake_api_user_name",
      } as BasicAuthConfig,
    };
    envHandler.updateEnvs(fakeConfig);
    expect(await fs.pathExists(path.join(botPath, localEnvFileName))).to.be.false;
    await envHandler.saveLocalEnvFile();
    let envs = dotenv.parse(await fs.readFile(path.join(botPath, localEnvFileName)));
    chai.assert.strictEqual(envs[Constants.envPrefix + "FAKE_ENDPOINT"], "fake_endpoint");
    chai.assert.strictEqual(envs[Constants.envPrefix + "FAKE_USERNAME"], "fake_api_user_name");
    chai.assert.exists(envs[Constants.envPrefix + "FAKE_PASSWORD"]);

    const fakeConfig2: ApiConnectorConfiguration = {
      ComponentPath: ["bot"],
      APIName: "FAKE",
      EndPoint: "fake_endpoint2",
      AuthConfig: {
        AuthType: AuthType.BASIC,
        UserName: "fake_api_user_name2",
      } as BasicAuthConfig,
    };
    envHandler.updateEnvs(fakeConfig2);
    await envHandler.saveLocalEnvFile();
    envs = dotenv.parse(await fs.readFile(path.join(botPath, localEnvFileName)));
    chai.assert.strictEqual(envs[Constants.envPrefix + "FAKE_ENDPOINT"], "fake_endpoint2");
    chai.assert.strictEqual(envs[Constants.envPrefix + "FAKE_USERNAME"], "fake_api_user_name2");
  });
});
