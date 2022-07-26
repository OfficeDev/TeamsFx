// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";

import { AzureSolutionSettings, ok, ProjectSettings } from "@microsoft/teamsfx-api";
import * as globalState from "@microsoft/teamsfx-core/build/common/globalState";
import { LocalEnvManager } from "@microsoft/teamsfx-core";

import * as commonUtils from "../../src/utils/commonUtils";
import * as extension from "../../src/extension";
import * as globalVariables from "../../src/globalVariables";
import * as teamsfxTaskHandler from "../../src/debug/teamsfxTaskHandler";
import { automaticNpmInstallHandler } from "../../src/debug/npmInstallHandler";
import * as ExtTelemetry from "../../src/telemetry/telemetry";

describe("npmInstallHandler", () => {
  describe("automaticNpmInstallHandler", () => {
    const workspaceFolder: vscode.WorkspaceFolder = {
      uri: vscode.Uri.file("test"),
      name: "test",
      index: 0,
    };
    sinon.stub(globalVariables, "workspaceUri").value(workspaceFolder.uri);
    let state: any;
    const solutionSettings: AzureSolutionSettings = {
      name: "fx-solution-azure",
      version: "1.0.0",
      hostType: "Azure",
      capabilities: [],
      azureResources: [],
      activeResourcePlugins: [],
    };
    const projectSettings: ProjectSettings = {
      appName: "test",
      version: "2.0.0",
      projectId: "11111111-1111-1111-1111-111111111111",
      solutionSettings,
    };

    let globalStateGetStub: sinon.SinonStub;
    let globalStateUpdateStub: sinon.SinonStub;
    let runTaskStub: sinon.SinonStub;
    let showMessageCalledCount: number;

    beforeEach(() => {
      sinon.restore();
      sinon.stub(vscode.workspace, "workspaceFolders").value([workspaceFolder]);
      globalStateGetStub = sinon.stub(globalState, "globalStateGet").callsFake(async () => state);
      globalStateUpdateStub = sinon
        .stub(globalState, "globalStateUpdate")
        .callsFake(async (key, value) => (state = value));
      sinon.stub(commonUtils, "getConfiguration").returns(true);
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      solutionSettings.hostType = "Azure";
      solutionSettings.capabilities = [];
      solutionSettings.azureResources = [];
      sinon
        .stub(LocalEnvManager.prototype, "getProjectSettings")
        .returns(Promise.resolve(projectSettings));
      showMessageCalledCount = 0;
      sinon.stub(extension, "VS_CODE_UI").value({
        showMessage: async () => {
          showMessageCalledCount += 1;
          return ok(undefined);
        },
      });
      runTaskStub = sinon.stub(teamsfxTaskHandler, "runTask").callsFake(async () => undefined);
    });

    afterEach(() => {
      sinon.restore();
    });

    it("Create SPFx", async () => {
      state = true;
      solutionSettings.hostType = "SPFx";
      solutionSettings.capabilities = ["TabSPFx"];
      await automaticNpmInstallHandler(false, false, false);
      sinon.assert.calledOnce(globalStateGetStub);
      sinon.assert.calledOnce(globalStateUpdateStub);
      chai.expect(state).false;
      chai.expect(showMessageCalledCount).equals(0);
      sinon.assert.notCalled(runTaskStub);
    });

    it("Create Tab", async () => {
      state = true;
      solutionSettings.capabilities = ["Tab"];
      await automaticNpmInstallHandler(false, false, false);
      sinon.assert.calledOnce(globalStateGetStub);
      sinon.assert.calledOnce(globalStateUpdateStub);
      chai.expect(state).false;
      chai.expect(showMessageCalledCount).equals(0);
      sinon.assert.notCalled(runTaskStub);
    });

    it("Create Bot", async () => {
      state = true;
      solutionSettings.capabilities = ["Bot"];
      await automaticNpmInstallHandler(false, false, false);
      sinon.assert.calledOnce(globalStateGetStub);
      sinon.assert.calledOnce(globalStateUpdateStub);
      chai.expect(state).false;
      chai.expect(showMessageCalledCount).equals(0);
      sinon.assert.notCalled(runTaskStub);
    });

    it("Create Message extension", async () => {
      state = true;
      solutionSettings.capabilities = ["MessagingExtension"];
      await automaticNpmInstallHandler(false, false, false);
      sinon.assert.calledOnce(globalStateGetStub);
      sinon.assert.calledOnce(globalStateUpdateStub);
      chai.expect(state).false;
      chai.expect(showMessageCalledCount).equals(0);
      sinon.assert.notCalled(runTaskStub);
    });

    it("Create Tab + Bot", async () => {
      state = true;
      solutionSettings.capabilities = ["Tab", "Bot"];
      await automaticNpmInstallHandler(false, false, false);
      sinon.assert.calledOnce(globalStateGetStub);
      sinon.assert.calledOnce(globalStateUpdateStub);
      chai.expect(state).false;
      chai.expect(showMessageCalledCount).equals(0);
      sinon.assert.notCalled(runTaskStub);
    });

    it("Create Tab + Message extension", async () => {
      state = true;
      solutionSettings.capabilities = ["Tab", "MessagingExtension"];
      await automaticNpmInstallHandler(false, false, false);
      sinon.assert.calledOnce(globalStateGetStub);
      sinon.assert.calledOnce(globalStateUpdateStub);
      chai.expect(state).false;
      chai.expect(showMessageCalledCount).equals(0);
      sinon.assert.notCalled(runTaskStub);
    });

    it("Create Tab + Bot + Message extension", async () => {
      state = true;
      solutionSettings.capabilities = ["Tab", "Bot", "MessagingExtension"];
      await automaticNpmInstallHandler(false, false, false);
      sinon.assert.calledOnce(globalStateGetStub);
      sinon.assert.calledOnce(globalStateUpdateStub);
      chai.expect(state).false;
      chai.expect(showMessageCalledCount).equals(0);
      sinon.assert.notCalled(runTaskStub);
    });

    it("Tab add Function", async () => {
      state = true;
      solutionSettings.capabilities = ["Tab"];
      solutionSettings.azureResources = ["function"];
      await automaticNpmInstallHandler(true, false, false);
      sinon.assert.calledOnce(globalStateGetStub);
      sinon.assert.calledOnce(globalStateUpdateStub);
      chai.expect(state).false;
      chai.expect(showMessageCalledCount).equals(0);
      sinon.assert.notCalled(runTaskStub);
    });

    it("Bot add Function", async () => {
      state = true;
      solutionSettings.capabilities = ["Bot"];
      solutionSettings.azureResources = ["function"];
      await automaticNpmInstallHandler(false, false, true);
      sinon.assert.calledOnce(globalStateGetStub);
      sinon.assert.calledOnce(globalStateUpdateStub);
      chai.expect(state).false;
      chai.expect(showMessageCalledCount).equals(0);
      sinon.assert.notCalled(runTaskStub);
    });

    it("Message extension add Function", async () => {
      state = true;
      solutionSettings.capabilities = ["MessagingExtension"];
      solutionSettings.azureResources = ["function"];
      await automaticNpmInstallHandler(false, false, true);
      sinon.assert.calledOnce(globalStateGetStub);
      sinon.assert.calledOnce(globalStateUpdateStub);
      chai.expect(state).false;
      chai.expect(showMessageCalledCount).equals(0);
      sinon.assert.notCalled(runTaskStub);
    });

    it("Tab + Bot add Function", async () => {
      state = true;
      solutionSettings.capabilities = ["Tab", "Bot"];
      solutionSettings.azureResources = ["function"];
      await automaticNpmInstallHandler(true, false, true);
      sinon.assert.calledOnce(globalStateGetStub);
      sinon.assert.calledOnce(globalStateUpdateStub);
      chai.expect(state).false;
      chai.expect(showMessageCalledCount).equals(0);
      sinon.assert.notCalled(runTaskStub);
    });

    it("Tab + Function add Function", async () => {
      state = true;
      solutionSettings.capabilities = ["Tab"];
      solutionSettings.azureResources = ["function"];
      await automaticNpmInstallHandler(true, true, false);
      sinon.assert.calledOnce(globalStateGetStub);
      sinon.assert.calledOnce(globalStateUpdateStub);
      chai.expect(state).false;
      chai.expect(showMessageCalledCount).equals(0);
      sinon.assert.notCalled(runTaskStub);
    });

    it("Tab add Bot", async () => {
      state = true;
      solutionSettings.capabilities = ["Tab", "Bot"];
      await automaticNpmInstallHandler(true, false, false);
      sinon.assert.calledOnce(globalStateGetStub);
      sinon.assert.calledOnce(globalStateUpdateStub);
      chai.expect(state).false;
      chai.expect(showMessageCalledCount).equals(0);
      sinon.assert.notCalled(runTaskStub);
    });

    it("Bot add Tab", async () => {
      state = true;
      solutionSettings.capabilities = ["Tab", "Bot"];
      await automaticNpmInstallHandler(false, false, true);
      sinon.assert.calledOnce(globalStateGetStub);
      sinon.assert.calledOnce(globalStateUpdateStub);
      chai.expect(state).false;
      chai.expect(showMessageCalledCount).equals(0);
      sinon.assert.notCalled(runTaskStub);
    });
  });
});
