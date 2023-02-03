// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author Ning Tang <nintan@microsoft.com>
 */

import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";

import * as commonTools from "@microsoft/teamsfx-core/build/common/tools";

import commandController from "../../src/commandController";
import TreeViewManagerInstance from "../../src/treeview/treeViewManager";

describe("Command Controller", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {});

  afterEach(() => {
    sandbox.restore();
  });

  it("directly call command callback in V3", async () => {
    const commandName = "fx-extension.provision";
    const commandCallback = sandbox.stub();

    sandbox.stub(commonTools, "isV3Enabled").returns(true);
    commandController.registerCommand(commandName, commandCallback);
    await commandController.runCommand(commandName, []);

    chai.assert.isTrue(commandCallback.calledOnce);
  });

  it("refresh UI when receiving lock events", async () => {
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();
    const setRunningCommandStub = sandbox.stub(TreeViewManagerInstance, "setRunningCommand");

    await commandController.lockedByOperation("provisionResources");

    chai.assert.isTrue(
      executeCommandStub.calledOnceWithExactly("setContext", "fx-extension.commandLocked", true)
    );
    chai.assert.isTrue(setRunningCommandStub.calledOnce);
  });

  it("refresh UI when receiving unlock events", async () => {
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();
    const setRunningCommandStub = sandbox.stub(TreeViewManagerInstance, "setRunningCommand");

    await commandController.unlockedByOperation("provisionResources");

    chai.assert.isTrue(
      executeCommandStub.calledOnceWithExactly("setContext", "fx-extension.commandLocked", true)
    );
    chai.assert.isTrue(setRunningCommandStub.calledOnce);
  });
});
