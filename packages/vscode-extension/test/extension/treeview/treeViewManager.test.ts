import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";

import * as globalVariables from "../../../src/globalVariables";
import { CommandsTreeViewProvider } from "../../../src/treeview/commandsTreeViewProvider";
import treeViewManager from "../../../src/treeview/treeViewManager";

describe("TreeViewManager", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("registerTreeViews", () => {
    treeViewManager.registerTreeViews({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);
    chai.assert.isDefined(treeViewManager.getTreeView("teamsfx-accounts"));

    const lifecycleTreeView = treeViewManager.getTreeView("teamsfx-lifecycle");
    chai.assert.isDefined(lifecycleTreeView);
    chai.assert.equal((lifecycleTreeView as any).commands.length, 3);
    chai.assert.equal((lifecycleTreeView as any).commands[0].commandId, "fx-extension.provision");
  });

  it("registerTreeViews", () => {
    sandbox.stub(globalVariables, "context").value({ extensionPath: "" });
    sandbox.stub(globalVariables, "isSPFxProject").value(false);
    treeViewManager.registerTreeViews({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);

    const developmentTreeview = treeViewManager.getTreeView("teamsfx-development");
    chai.assert.isDefined(developmentTreeview);
    chai.assert.equal((developmentTreeview as any).commands.length, 5);
  });

  it("setRunningCommand", () => {
    treeViewManager.registerTreeViews({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);
    const command = (treeViewManager as any).commandMap.get("fx-extension.create");
    const setStatusStub = sinon.stub(command, "setStatus");
    treeViewManager.setRunningCommand("fx-extension.create", ["fx-extension.openSamples"]);

    chai.assert.equal(setStatusStub.callCount, 1);

    treeViewManager.restoreRunningCommand(["fx-extension.openSamples"]);
    chai.assert.equal(setStatusStub.callCount, 2);
  });

  it("updateTreeViewsOnSPFxChanged", () => {
    sandbox.stub(globalVariables, "isSPFxProject").value(false);
    treeViewManager.registerTreeViews({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);
    const developmentTreeviewProvider = treeViewManager.getTreeView(
      "teamsfx-development"
    ) as CommandsTreeViewProvider;

    const commands = developmentTreeviewProvider.getCommands();
    chai.assert.equal(commands.length, 5);

    sandbox.stub(globalVariables, "isSPFxProject").value(true);
    treeViewManager.updateTreeViewsOnSPFxChanged();

    chai.assert.equal(commands.length, 6);
  });
});
