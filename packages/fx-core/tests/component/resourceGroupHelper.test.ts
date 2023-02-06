import { ResourceManagementClient } from "@azure/arm-resources";
import { ok } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import * as sinon from "sinon";
import { resourceGroupHelper } from "../../src/component/utils/ResourceGroupHelper";
import { setTools } from "../../src/core/globalVars";
import { MockTools } from "../core/utils";
import { MyTokenCredential } from "../plugins/solution/util";

describe("resouce group helper test", () => {
  const sandbox = sinon.createSandbox();
  const tools = new MockTools();
  setTools(tools);
  let mockedEnvRestore: RestoreFn | undefined;

  afterEach(() => {
    sandbox.restore();
    if (mockedEnvRestore) {
      mockedEnvRestore();
    }
  });

  beforeEach(() => {
    mockedEnvRestore = mockedEnv({
      TEAMSFX_V3: "true",
    });
  });

  it("askResourceGroupInfoV3", async () => {
    sandbox.stub(resourceGroupHelper, "listResourceGroups").resolves(ok([["rg1", "loc1"]]));
    sandbox.stub(resourceGroupHelper, "getLocations").resolves(ok(["loc1"]));
    sandbox.stub(tools.ui, "selectOption").resolves(ok({ type: "success", result: "rg1" }));
    const mockResourceManagementClient = new ResourceManagementClient(
      new MyTokenCredential(),
      "id"
    );
    const res = await resourceGroupHelper.askResourceGroupInfoV3(
      tools.tokenProvider.azureAccountProvider,
      mockResourceManagementClient,
      "rg1"
    );
    assert.isTrue(res.isOk());
  });
});
