import { Err, Ok, SystemError } from "@microsoft/teamsfx-api";
import { PackageService } from "@microsoft/teamsfx-core";
import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import M365TokenInstance from "../../../src/commonlib/m365Login";
import { infoIcon, passIcon, warningIcon } from "../../../src/treeview/account/common";
import { CopilotNode } from "../../../src/treeview/account/copilotNode";
import { DynamicNode } from "../../../src/treeview/dynamicNode";
import * as checkAccessCallback from "../../../src/handlers/accounts/checkAccessCallback";

describe("copilotNode", () => {
  const sandbox = sinon.createSandbox();
  const eventEmitter = new vscode.EventEmitter<DynamicNode | undefined | void>();

  afterEach(() => {
    sandbox.restore();
  });

  it("getTreeItem with empty string", async () => {
    const copilotNode = new CopilotNode(eventEmitter, "");
    const treeItem = await copilotNode.getTreeItem();

    chai.assert.equal(treeItem.iconPath, infoIcon);
  });

  it("getTreeItem with check false", async () => {
    sandbox
      .stub(M365TokenInstance, "getAccessToken")
      .returns(Promise.resolve(new Ok("test-token")));
    sandbox.stub(PackageService, "GetSharedInstance").returns(new PackageService("endpoint"));
    sandbox.stub(PackageService.prototype, "getCopilotStatus").resolves(false);
    sandbox.stub(checkAccessCallback, "checkCopilotCallback");
    const copilotNode = new CopilotNode(eventEmitter, "token");
    const treeItem = await copilotNode.getTreeItem();

    chai.assert.equal(treeItem.iconPath, warningIcon);
  });

  it("getTreeItem with check true", async () => {
    sandbox
      .stub(M365TokenInstance, "getAccessToken")
      .returns(Promise.resolve(new Ok("test-token")));
    sandbox.stub(PackageService, "GetSharedInstance").returns(new PackageService("endpoint"));
    sandbox.stub(PackageService.prototype, "getCopilotStatus").resolves(true);
    const copilotNode = new CopilotNode(eventEmitter, "token");
    const treeItem = await copilotNode.getTreeItem();

    chai.assert.equal(treeItem.iconPath, passIcon);
  });

  it("getTreeItem with check error", async () => {
    sandbox
      .stub(M365TokenInstance, "getAccessToken")
      .returns(Promise.resolve(new Ok("test-token")));
    sandbox.stub(PackageService, "GetSharedInstance").returns(new PackageService("endpoint"));
    sandbox
      .stub(PackageService.prototype, "getCopilotStatus")
      .returns(Promise.reject(new Error("test-error")));
    const copilotNode = new CopilotNode(eventEmitter, "token");
    const treeItem = await copilotNode.getTreeItem();

    chai.assert.equal(treeItem.iconPath, infoIcon);
  });

  it("getTreeItem with token error", async () => {
    sandbox
      .stub(M365TokenInstance, "getAccessToken")
      .returns(Promise.resolve(new Err(new SystemError("test-source", "test-name", "test-error"))));
    const copilotNode = new CopilotNode(eventEmitter, "token");
    const treeItem = await copilotNode.getTreeItem();

    chai.assert.equal(treeItem.iconPath, infoIcon);
  });

  it("getTreeItem with empty token", async () => {
    sandbox.stub(M365TokenInstance, "getAccessToken").returns(Promise.resolve(new Ok("")));
    const copilotNode = new CopilotNode(eventEmitter, "token");
    const treeItem = await copilotNode.getTreeItem();

    chai.assert.equal(treeItem.iconPath, infoIcon);
  });

  it("getChildren", () => {
    const copilotNode = new CopilotNode(eventEmitter, "token");
    chai.assert.isNull(copilotNode.getChildren());
  });
});
