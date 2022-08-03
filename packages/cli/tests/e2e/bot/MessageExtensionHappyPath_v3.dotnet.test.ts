// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { happyPathTest } from "./BotHappyPathCommon";
import { Runtime } from "../../commonlib/constants";
import { it } from "../../commonlib/it";
import mockedEnv from "mocked-env";
describe("Provision message extension Dotnet", () => {
  let mockedEnvRestore: () => void;
  before(async () => {
    mockedEnvRestore = mockedEnv({
      TEAMSFX_APIV3: "true",
    });
  });
  after(async () => {
    mockedEnvRestore();
  });
  it("Provision Resource: message extension dotnet", async function () {
    await happyPathTest(Runtime.Dotnet, "message-extension");
  });
});
