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
import { AadValidator, FrontendValidator } from "../../commonlib";
import { TemplateProject } from "../../commonlib/constants";
import { Executor } from "../../utils/executor";
import m365Login from "../../../src/commonlib/m365Login";
import { environmentManager } from "@microsoft/teamsfx-core/build/core/environment";
import { assert } from "chai";

describe("teamsfx new template", function () {
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);
  const env = environmentManager.getDefaultEnvName();

  it(`${TemplateProject.OneProductivityHub}`, { testPlanCaseId: 15277463 }, async function () {
    await Executor.openTemplateProject(appName, testFolder, TemplateProject.OneProductivityHub);
    expect(fs.pathExistsSync(projectPath)).to.be.true;
    expect(fs.pathExistsSync(path.resolve(projectPath, "infra"))).to.be.true;

    // Provision
    {
      const { success, stderr } = await Executor.provision(projectPath);
      if (!success) {
        console.log(stderr);
        assert.fail("Provision failed");
      }
    }

    // Validate Provision
    const context = await readContextMultiEnvV3(projectPath, env);

    // Validate Aad App
    const aad = AadValidator.init(context, false, m365Login);
    await AadValidator.validate(aad);

    // Validate Tab Frontend
    const frontend = FrontendValidator.init(context);
    await FrontendValidator.validateProvision(frontend);

    // deploy
    {
      const { success, stderr } = await Executor.deploy(projectPath);
      if (!success) {
        console.log(stderr);
        assert.fail("Deploy failed");
      }
    }
  });

  after(async () => {
    await cleanUp(appName, projectPath, true, false, false);
  });
});
