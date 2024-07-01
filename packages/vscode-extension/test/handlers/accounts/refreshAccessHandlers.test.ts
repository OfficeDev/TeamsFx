import * as sinon from "sinon";
import * as chai from "chai";
import { ok } from "@microsoft/teamsfx-api";
import {
  refreshCopilotCallback,
  refreshSideloadingCallback,
} from "../../../src/handlers/accounts/refreshAccessHandlers";
import M365TokenInstance from "../../../src/commonlib/m365Login";
import accountTreeViewProviderInstance from "../../../src/treeview/account/accountTreeViewProvider";

describe("refreshAccessHandlers", () => {
  describe("refreshSideloadingCallback", async () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("Happy path", async () => {
      const status = {
        status: "success",
        token: "test-token",
      };
      sandbox.stub(M365TokenInstance, "getStatus").resolves(ok(status));
      const updateChecksStub = sandbox.stub(
        accountTreeViewProviderInstance.m365AccountNode,
        "updateChecks"
      );
      await refreshSideloadingCallback();
      chai.assert(updateChecksStub.calledOnceWithExactly("test-token", true, false));
    });

    it("No token", async () => {
      const status = {
        status: "success",
      };
      sandbox.stub(M365TokenInstance, "getStatus").resolves(ok(status));
      const updateChecksStub = sandbox.stub(
        accountTreeViewProviderInstance.m365AccountNode,
        "updateChecks"
      );
      await refreshSideloadingCallback();
      chai.assert(updateChecksStub.notCalled);
    });
  });

  describe("refreshCopilotCallback", async () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("Happy path", async () => {
      const status = {
        status: "success",
        token: "test-token",
      };
      sandbox.stub(M365TokenInstance, "getStatus").resolves(ok(status));
      const updateChecksStub = sandbox.stub(
        accountTreeViewProviderInstance.m365AccountNode,
        "updateChecks"
      );
      await refreshCopilotCallback();
      chai.assert(updateChecksStub.calledOnceWithExactly("test-token", false, true));
    });

    it("No token", async () => {
      const status = {
        status: "success",
      };
      sandbox.stub(M365TokenInstance, "getStatus").resolves(ok(status));
      const updateChecksStub = sandbox.stub(
        accountTreeViewProviderInstance.m365AccountNode,
        "updateChecks"
      );
      await refreshCopilotCallback();
      chai.assert(updateChecksStub.notCalled);
    });
  });
});
