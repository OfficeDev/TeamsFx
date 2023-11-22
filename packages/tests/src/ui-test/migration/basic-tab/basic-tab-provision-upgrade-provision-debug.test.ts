// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { MigrationTestContext } from "../migrationContext";
import { Timeout, Capability, Notification } from "../../../utils/constants";
import { it } from "../../../utils/it";
import { Env } from "../../../utils/env";
import {
  initPage,
  validateTabNoneSSO,
} from "../../../utils/playwrightOperation";
import { CliHelper } from "../../cliHelper";
import {
  validateNotification,
  validateUpgrade,
  upgradeByCommandPalette,
} from "../../../utils/vscodeOperation";
import * as dotenv from "dotenv";
import { CLIVersionCheck } from "../../../utils/commonUtils";

dotenv.config();

describe("Migration Tests", function () {
  this.timeout(Timeout.testAzureCase);
  let mirgationDebugTestContext: MigrationTestContext;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);

    mirgationDebugTestContext = new MigrationTestContext(
      Capability.TabNonSso,
      "javascript"
    );
    await mirgationDebugTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    await mirgationDebugTestContext.after(false, false, "dev");
  });

  it(
    "[auto] Basic Tab app with sso migrate test - js",
    {
      testPlanCaseId: 17184120,
      author: "v-helzha@microsoft.com",
    },
    async () => {
      // create v2 project using CLI
      await mirgationDebugTestContext.createProjectCLI(false);
      // verify popup
      try {
        await validateNotification(Notification.Upgrade);
      } catch (error) {
        await validateNotification(Notification.Upgrade_dicarded);
      }

      // v2 provision
      await mirgationDebugTestContext.provisionWithCLI("dev", false);

      // upgrade
      await upgradeByCommandPalette();
      // verify upgrade
      await validateUpgrade();

      // enable cli v3
      await CliHelper.installCLI(
        "alpha",
        false,
        mirgationDebugTestContext.projectPath
      );
      CliHelper.setV3Enable();

      // v3 provision
      await mirgationDebugTestContext.provisionWithCLI("dev", true);
      await CLIVersionCheck("V3", mirgationDebugTestContext.projectPath);
      // v3 deploy
      await mirgationDebugTestContext.deployWithCLI("dev");

      // UI verify
      const teamsAppId = await mirgationDebugTestContext.getTeamsAppId("dev");
      const page = await initPage(
        mirgationDebugTestContext.context!,
        teamsAppId,
        Env.username,
        Env.password
      );
      await validateTabNoneSSO(page);
    }
  );
});
