// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Zhaofeng Xu <zhaofengxu@microsoft.com>
 */

import { it } from "@microsoft/extra-shot-mocha";
import { isV3Enabled } from "@microsoft/teamsfx-core/build/common/tools";
import * as chai from "chai";
import { describe } from "mocha";
import * as path from "path";
import { CliHelper } from "../../commonlib/cliHelper";
import { Cleaner } from "../../utils/cleaner";
import { TemplateProject } from "../../commonlib/constants";
import { Executor } from "../../utils/executor";
import { getTestFolder, getUniqueAppName } from "../commonUtils";

describe("upgrade", () => {
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);

  afterEach(async function () {
    await Cleaner.clean(projectPath);
  });

  it("sample incoming webhook notification", { testPlanCaseId: 19298763 }, async function () {
    if (!isV3Enabled()) {
      return;
    }

    {
      Executor.installCLI(testFolder, "1.2.5", true);
      const env = Object.assign({}, process.env);
      env["TEAMSFX_V3"] = "false";
      // new projiect
      await CliHelper.createTemplateProject(
        appName,
        testFolder,
        TemplateProject.IncomingWebhook,
        env
      );
    }

    {
      // provision
      const result = await Executor.provision(projectPath);
      chai.assert.isFalse(result.success);
      chai.assert.include(
        result.stderr,
        "This command only works for project created by Teams Toolkit"
      );
    }
  });
});
