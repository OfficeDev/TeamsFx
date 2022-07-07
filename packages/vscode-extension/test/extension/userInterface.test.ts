// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { ExtensionContext } from "vscode";

import { sleep } from "../../src/utils/commonUtils";
import { VsCodeUI } from "../../src/qm/vsc_ui";

describe("UI Unit Tests", async () => {
  before(() => {
    // Mock user input.
  });

  describe("Manually", () => {
    it("Show Progress 2", async function (this: Mocha.Context) {
      this.timeout(0);
      const VS_CODE_UI = new VsCodeUI(<ExtensionContext>{});
      const handler = VS_CODE_UI.createProgressBar("Test Progress Bar", 3);

      await handler.start("Prepare");
      await sleep(2 * 1000);

      await handler.next("First step");
      await sleep(2 * 1000);

      await handler.next("Second step");
      await sleep(2 * 1000);

      await handler.next("Third step");
      await sleep(2 * 1000);

      await handler.end(true);
    });
  });
});
