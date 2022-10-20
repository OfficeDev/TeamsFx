// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { expect } from "chai";
import fs from "fs-extra";
import path from "path";
import { it } from "@microsoft/extra-shot-mocha";
import {
  execAsync,
  getTestFolder,
  cleanUp,
  setSimpleAuthSkuNameToB1Bicep,
  getSubscriptionId,
  readContextMultiEnv,
  validateTabAndBotProjectProvision,
  getUniqueAppName,
} from "../commonUtils";
import { SqlValidator, FunctionValidator } from "../../commonlib";
import { getUuid } from "../../commonlib/utilities";
import { TemplateProject } from "../../commonlib/constants";
import { CliHelper } from "../../commonlib/cliHelper";
import { environmentManager } from "@microsoft/teamsfx-core/build/core/environment";

describe("teamsfx new template", function () {
  const testFolder = getTestFolder();
  const subscription = getSubscriptionId();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);
  const env = environmentManager.getDefaultEnvName();

  it(`${TemplateProject.ShareNow}`, { testPlanCaseId: 15277467 }, async function () {
    await CliHelper.createTemplateProject(
      appName,
      testFolder,
      TemplateProject.ShareNow,
      TemplateProject.ShareNow
    );

    expect(fs.pathExistsSync(projectPath)).to.be.true;
    expect(fs.pathExistsSync(path.resolve(projectPath, ".fx"))).to.be.true;

    // Provision
    await setSimpleAuthSkuNameToB1Bicep(projectPath, env);
    await CliHelper.setSubscription(subscription, projectPath);
    await CliHelper.provisionProject(
      projectPath,
      `--sql-admin-name Abc123321 --sql-password Cab232332${getUuid().substring(0, 6)}`
    );

    // Validate Provision
    await validateTabAndBotProjectProvision(projectPath, env);

    await execAsync(`Set EXPO_DEBUG=true && npm config set package-lock false`, {
      cwd: path.join(projectPath, "tabs"),
      env: process.env,
      timeout: 0,
    });

    const result = await execAsync(`npm i @types/node -D`, {
      cwd: path.join(projectPath, "tabs"),
      env: process.env,
      timeout: 0,
    });
    if (!result.stderr) {
      console.log("success to run cmd: npm i @types/node -D");
    } else {
      console.log("[failed] ", result.stderr);
    }

    // deploy
    await CliHelper.deployAll(projectPath);

    // Assert
    {
      const context = await readContextMultiEnv(projectPath, env);

      // Validate Function App
      const functionValidator = new FunctionValidator(context, projectPath, env);
      await functionValidator.validateProvision();
      await functionValidator.validateDeploy();

      // Validate sql
      await SqlValidator.init(context);
      await SqlValidator.validateSql();
    }
  });

  after(async () => {
    await cleanUp(appName, projectPath, true, true, false);
  });
});
