import { UserError, err, ok } from "@microsoft/teamsfx-api";
import "mocha";
import * as sinon from "sinon";
import * as fs from "fs-extra";
import * as global from "../../../src/globalVariables";
import {
  checkProjectTypeAndSendTelemetry,
  isTestToolEnabledProject,
} from "../../../src/utils/projectChecker";
import { MockCore } from "../../mocks/mockCore";
import * as vscode from "vscode";
import { ExtTelemetry } from "../../../src/telemetry/extTelemetry";

describe("projectChecker", () => {
  describe("checkProjectTypeAndSendTelemetry", () => {
    const sandbox = sinon.createSandbox();
    const core = new MockCore();

    afterEach(() => {
      sandbox.restore();
    });

    it("happy", async () => {
      sandbox.stub(global, "workspaceUri").value(vscode.Uri.file("./"));
      sandbox.stub(global, "core").value(core);
      sandbox.stub(core, "checkProjectType").resolves(
        ok({
          isTeamsFx: true,
          hasTeamsManifest: true,
          dependsOnTeamsJs: false,
          lauguages: ["ts"],
        })
      );
      sandbox.stub(ExtTelemetry, "addSharedProperty");
      await checkProjectTypeAndSendTelemetry();
    });

    it("error", async () => {
      sandbox.stub(global, "workspaceUri").value(vscode.Uri.file("./"));
      sandbox.stub(global, "core").value(core);
      sandbox.stub(core, "checkProjectType").resolves(err(new UserError({})));
      await checkProjectTypeAndSendTelemetry();
    });

    it("workspaceUri is undefined", async () => {
      sandbox.stub(global, "workspaceUri").value(undefined);
      await checkProjectTypeAndSendTelemetry();
    });
  });

  describe("isTestToolEnabledProject", () => {
    const sandbox = sinon.createSandbox();

    afterEach(async () => {
      sandbox.restore();
    });

    it("test tool yaml exist", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(true);
      const res = isTestToolEnabledProject("testPath");
      chai.assert.isTrue(res);
    });

    it("test tool yaml not exist", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(false);
      const res = isTestToolEnabledProject("testPath");
      chai.assert.isFalse(res);
    });
  });
});
