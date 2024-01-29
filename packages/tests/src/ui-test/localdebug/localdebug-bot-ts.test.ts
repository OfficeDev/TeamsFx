// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Xiaofu Huang <xiaofu.huang@microsoft.com>
 */
import * as path from "path";
import { startDebugging, waitForTerminal } from "../../utils/vscodeOperation";
import {
  initPage,
  reopenPage,
  validateEchoBot,
} from "../../utils/playwrightOperation";
import { LocalDebugTestContext } from "./localdebugContext";
import {
  Timeout,
  LocalDebugTaskLabel,
  DebugItemSelect,
  LocalDebugError,
} from "../../utils/constants";
import { Env } from "../../utils/env";
import { it } from "../../utils/it";
import { validateFileExist } from "../../utils/commonUtils";
import { ChildProcessWithoutNullStreams } from "child_process";
import { Executor } from "../../utils/executor";
import { expect } from "chai";
import { VSBrowser } from "vscode-extension-tester";
import { getScreenshotName } from "../../utils/nameUtil";

describe("Local Debug Tests", function () {
  this.timeout(Timeout.testCase);
  let localDebugTestContext: LocalDebugTestContext;
  let devtunnelProcess: ChildProcessWithoutNullStreams | null;
  let debugProcess: ChildProcessWithoutNullStreams | null;
  let tunnelName = "";
  let successFlag = true;
  let errorMessage = "";

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    localDebugTestContext = new LocalDebugTestContext("bot", "typescript");
    await localDebugTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    if (debugProcess) {
      setTimeout(() => {
        debugProcess?.kill("SIGTERM");
      }, 2000);
    }

    if (tunnelName) {
      setTimeout(() => {
        devtunnelProcess?.kill("SIGTERM");
      }, 2000);
      Executor.deleteTunnel(
        tunnelName,
        (data) => {
          if (data) {
            console.log(data);
          }
        },
        (error) => {
          console.log(error);
        }
      );
    }
    await localDebugTestContext.after(false, true);
    this.timeout(Timeout.finishAzureTestCase);
  });

  it(
    "[auto] [Typescript] Local Debug for bot project",
    {
      testPlanCaseId: 9729308,
      author: "xiaofu.huang@microsoft.com",
    },
    async function () {
      try {
        const projectPath = path.resolve(
          localDebugTestContext.testRootFolder,
          localDebugTestContext.appName
        );
        validateFileExist(projectPath, "index.ts");

        // local debug
        console.log("======= debug with ttk ========");
        try {
          await startDebugging(DebugItemSelect.DebugInTeamsUsingChrome);
          await waitForTerminal(LocalDebugTaskLabel.StartLocalTunnel);
          await waitForTerminal(LocalDebugTaskLabel.StartBotApp, "Bot Started");
        } catch (error) {
          const errorMsg = error.toString();
          if (
            // skip can't find element
            errorMsg.includes(LocalDebugError.ElementNotInteractableError) ||
            // skip timeout
            errorMsg.includes(LocalDebugError.TimeoutError) ||
            // skip node 16 warning
            errorMsg.includes(LocalDebugError.FilePermission)
          ) {
            console.log("[skip error] ", error);
          } else {
            expect.fail(errorMsg);
          }
        }

        const teamsAppId = await localDebugTestContext.getTeamsAppId();
        expect(teamsAppId).to.not.be.empty;
        {
          const page = await initPage(
            localDebugTestContext.context!,
            teamsAppId,
            Env.username,
            Env.password
          );
          await localDebugTestContext.validateLocalStateForBot();
          await validateEchoBot(page);
        }

        // cli preview
        const res = await Executor.cliPreview(projectPath, true);
        devtunnelProcess = res.devtunnelProcess;
        tunnelName = res.tunnelName;
        debugProcess = res.debugProcess;
        {
          const page = await reopenPage(
            localDebugTestContext.context!,
            teamsAppId,
            Env.username,
            Env.password
          );
          await localDebugTestContext.validateLocalStateForBot();
          await validateEchoBot(page);
        }
      } catch (error) {
        successFlag = false;
        errorMessage = "[Error]: " + error;
        await VSBrowser.instance.takeScreenshot(getScreenshotName("error"));
        await VSBrowser.instance.driver.sleep(Timeout.playwrightDefaultTimeout);
      }
      expect(successFlag, errorMessage).to.true;
      console.log("debug finish!");
    }
  );
});
