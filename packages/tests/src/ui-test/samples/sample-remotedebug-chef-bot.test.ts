// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject, ValidationContent } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";
import { SampledebugContext } from "./sampledebugContext";
import { validateWelcomeAndReplyBot } from "../../utils/playwrightOperation";
import * as path from "path";
import * as fs from "fs";
import * as os from "os";

class ChefBotTestCase extends CaseFactory {
  public override async onAfterCreate(
    sampledebugContext: SampledebugContext,
    env: "local" | "dev"
  ): Promise<void> {
    const envFile = path.resolve(
      sampledebugContext.projectPath,
      ".env",
    );
    // create .env file
    fs.writeFileSync(envFile, "OPENAI_API_KEY=yourapikey");
    console.log(`add OPENAI_API_KEY=yourapikey to .env file`);
  }
  override async onValidate(page: Page): Promise<void> {
    console.log("Mocked api key. Only verify happy path...");
    return await validateWelcomeAndReplyBot(page, {
      hasCommandReplyValidation: true,
      botCommand: "helloWorld",
      expectedReplyMessage: ValidationContent.AiBotErrorMessage,
    });
  }
}

new ChefBotTestCase(
  TemplateProject.ChefBot,
  24409842,
  "v-ivanchen@microsoft.com",
  "dev",
  [],
  { testRootFolder: path.resolve(os.homedir(), "resourse") } // fix yarn error
).test();
