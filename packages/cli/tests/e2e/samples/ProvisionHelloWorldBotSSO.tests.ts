// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { expect } from "chai";
import fs from "fs-extra";
import path from "path";
import { it } from "@microsoft/extra-shot-mocha";
import { getTestFolder, cleanUp, readContextMultiEnvV3, getUniqueAppName } from "../commonUtils";
import { AadValidator, BotValidator } from "../../commonlib";
import { TemplateProject } from "../../commonlib/constants";
import { Executor } from "../../utils/executor";
import m365Login from "../../../src/commonlib/m365Login";
import { environmentManager } from "@microsoft/teamsfx-core/build/core/environment";

describe("teamsfx new template", function () {
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);
  const env = environmentManager.getDefaultEnvName();

  it(`${TemplateProject.HelloWorldBotSSO}`, { testPlanCaseId: 15277464 }, async function () {
    await Executor.openTemplateProject(appName, testFolder, TemplateProject.HelloWorldBotSSO);
    expect(fs.pathExistsSync(projectPath)).to.be.true;
    expect(fs.pathExistsSync(path.resolve(projectPath, "infra"))).to.be.true;

    // Provision
    await Executor.provision(projectPath);

    // Validate Provision
    const context = await readContextMultiEnvV3(projectPath, env);

    // Validate Bot Provision
    const bot = new BotValidator(context, projectPath, env);
    await bot.validateProvisionV3(false);

    // deploy
    await Executor.deploy(projectPath);

    // Validate deployment
    {
      // Get context
      const context = await readContextMultiEnvV3(projectPath, env);

      // Validate Aad App
      const aad = AadValidator.init(context, false, m365Login);
      await AadValidator.validate(aad);

      // Validate Bot Deploy
      const bot = new BotValidator(context, projectPath, env);
      await bot.validateDeploy();
    }

    // test (validate)
    await Executor.validate(projectPath);

    // package
    await Executor.package(projectPath);
  });

  after(async () => {
    await cleanUp(appName, projectPath, true, true, false);
  });
});
