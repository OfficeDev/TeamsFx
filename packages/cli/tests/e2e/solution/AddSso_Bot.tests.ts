// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Bowen Song <bowsong@microsoft.com>
 */

import path from "path";
import {
  getTestFolder,
  getUniqueAppName,
  cleanUp,
  execAsync,
  readContextMultiEnv,
} from "../commonUtils";
import { CliHelper } from "../../commonlib/cliHelper";
import { Capability, PluginId, ProjectSettingKey, TestFilePath } from "../../commonlib/constants";
import fs from "fs-extra";
import { expect } from "chai";
import { AadValidator, BotValidator } from "../../commonlib";
import M365Login from "../../../src/commonlib/m365Login";
import mockedEnv, { RestoreFn } from "mocked-env";

describe("Add SSO", () => {
  const testFolder = getTestFolder();
  let appName: string | undefined;
  let projectPath: string | undefined;
  let mockedEnvRestore: RestoreFn | undefined;

  afterEach(async () => {
    if (appName && projectPath) {
      await cleanUp(appName, projectPath, true, true, false);
    }
    if (mockedEnvRestore) {
      mockedEnvRestore();
    }
  });
  for (const v3flag of ["false", "true"]) {
    it(`Add SSO to non SSO Bot project (v3=${v3flag})`, async () => {
      mockedEnvRestore = mockedEnv({
        TEAMSFX_APIV3: v3flag,
        TEAMSFX_AAD_MANIFEST: "true",
        TEAMSFX_CONFIG_UNIFY: "true",
      });
      appName = getUniqueAppName();
      projectPath = path.resolve(testFolder, appName);
      // Arrange
      await CliHelper.createProjectWithCapability(appName, testFolder, Capability.Bot);

      // Assert
      {
        const projectSettings = await fs.readJSON(
          path.join(projectPath, TestFilePath.configFolder, TestFilePath.projectSettingsFileName)
        );
        const activeResourcePlugins =
          projectSettings[ProjectSettingKey.solutionSettings][
            ProjectSettingKey.activeResourcePlugins
          ];
        const capabilities =
          projectSettings[ProjectSettingKey.solutionSettings][ProjectSettingKey.capabilities];
        expect(activeResourcePlugins.includes(PluginId.Aad)).to.be.false;
        expect(capabilities.includes(Capability.BotSso)).to.be.false;
      }

      // Act
      await execAsync(`teamsfx add sso`, {
        cwd: projectPath,
        env: process.env,
        timeout: 0,
      });

      // Assert
      {
        const projectSettings = await fs.readJSON(
          path.join(projectPath, TestFilePath.configFolder, TestFilePath.projectSettingsFileName)
        );
        const activeResourcePlugins =
          projectSettings[ProjectSettingKey.solutionSettings][
            ProjectSettingKey.activeResourcePlugins
          ];
        const capabilities =
          projectSettings[ProjectSettingKey.solutionSettings][ProjectSettingKey.capabilities];
        expect(activeResourcePlugins.includes(PluginId.Aad)).to.be.true;
        expect(capabilities.includes(Capability.BotSso)).to.be.true;

        const readmeFilePath = path.join(projectPath, "auth", "bot", "README.md");
        const readmeExists = await fs.pathExists(readmeFilePath);
        expect(readmeExists).to.be.true;
      }

      await CliHelper.provisionProject(projectPath, "");

      const context = await readContextMultiEnv(projectPath, "dev");
      // Validate Aad App
      const aad = AadValidator.init(context, false, M365Login);
      await AadValidator.validate(aad);

      // Validate Bot Provision
      const bot = new BotValidator(context, projectPath, "dev");
      await bot.validateProvision();

      // Act
      try {
        await execAsync(`teamsfx add sso`, {
          cwd: projectPath,
          env: process.env,
          timeout: 0,
        });
      } catch (error) {
        expect(error.toString()).to.contains("SsoEnabled");
      }
    });
  }
});
