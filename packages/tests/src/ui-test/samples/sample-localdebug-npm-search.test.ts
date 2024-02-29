// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { validateNpm } from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";

class NpmSearchTestCase extends CaseFactory {
  override async onValidate(
    page: Page,
    options?: { npmName: string }
  ): Promise<void> {
    return await validateNpm(page, { npmName: options?.npmName });
  }

  override async onCliValidate(
    page: Page,
    options?: { npmName: string }
  ): Promise<void> {
    return await validateNpm(page, { npmName: options?.npmName });
  }
}

new NpmSearchTestCase(
  TemplateProject.NpmSearch,
  12664761,
  "v-ivanchen@microsoft.com",
  "local",
  [LocalDebugTaskLabel.StartLocalTunnel, LocalDebugTaskLabel.StartBotApp],
  { npmName: "axios", debug: "cli" }
).test();
