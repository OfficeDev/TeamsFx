import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";

import * as handlers from "../../src/copilotChatHandlers";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import * as extTelemetryEvents from "../../src/telemetry/extTelemetryEvents";
import VsCodeLogInstance from "../../src/commonlib/log";

describe("invokeTeamsAgent", async () => {
  const sandbox = sinon.createSandbox();
  let clock: sinon.SinonFakeTimers;
  afterEach(() => {
    sandbox.restore();
    if (clock) {
      clock.restore();
    }
  });

  beforeEach(() => {
    sandbox.stub(ExtTelemetry, "dispose");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    sandbox.stub(VsCodeLogInstance, "outputChannel").value({
      name: "name",
      append: (value: string) => {},
      appendLine: (value: string) => {},
      replace: (value: string) => {},
      clear: () => {},
      show: (...params: any[]) => {},
      hide: () => {},
      dispose: () => {},
    });
  });
  it("no need to install Github Copilot", async () => {
    sandbox.stub(vscode.extensions, "getExtension").returns({ name: "github.copilot" } as any);
    sandbox.stub(vscode.commands, "executeCommand").resolves();

    const res = await handlers.invokeTeamsAgent([
      extTelemetryEvents.TelemetryTriggerFrom.CreateAppQuestionFlow,
    ]);

    chai.assert.isTrue(res.isOk());
  });

  it("install Github Copilot and invoke Teams Agent", async () => {
    clock = sandbox.useFakeTimers();
    sandbox.stub(vscode, "version").value("1.88.0-insiders");
    sandbox
      .stub(vscode.extensions, "getExtension")
      .onFirstCall()
      .returns(undefined)
      .onSecondCall()
      .returns({ name: "github.copilot" } as any);
    const commandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();
    sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves("Install" as unknown as vscode.MessageItem);

    const job = handlers.invokeTeamsAgent([extTelemetryEvents.TelemetryTriggerFrom.TreeView]);
    await clock.tickAsync(6000);
    const res = await job;

    if (res.isErr()) {
      console.log(res.error);
    }

    chai.assert.isTrue(res.isOk());
    chai.assert.equal(commandStub.callCount, 3);
    chai.assert.equal(commandStub.getCall(2).args[1].query, "@teams");
  });

  it("install Github Copilot, wait and invoke Teams Agent", async () => {
    clock = sandbox.useFakeTimers();
    sandbox.stub(vscode, "version").value("1.88.0-insiders");
    sandbox
      .stub(vscode.extensions, "getExtension")
      .onFirstCall()
      .returns(undefined)
      .onSecondCall()
      .returns(undefined)
      .onThirdCall()
      .returns({ name: "github.copilot" } as any);
    const commandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();
    sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves("Install" as unknown as vscode.MessageItem);

    const job = handlers.invokeTeamsAgent();
    await clock.tickAsync(6000);
    const res = await job;

    chai.assert.isTrue(res.isOk());
    chai.assert.equal(commandStub.callCount, 3);
  });

  it("Install github copilot extension error", async () => {
    sandbox.stub(vscode, "version").value("1.88.0");
    sandbox.stub(vscode.extensions, "getExtension").onFirstCall().returns(undefined);
    const commandStub = sandbox
      .stub(vscode.commands, "executeCommand")
      .callsFake(async (command: string) => {
        if (command === "workbench.extensions.installExtension") {
          throw new Error("Install Error");
        } else {
          return {};
        }
      });
    sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves("Install" as unknown as vscode.MessageItem);
    sandbox.stub(VsCodeLogInstance, "error").resolves();

    const res = await handlers.invokeTeamsAgent();

    chai.assert.isTrue(res.isErr());
    if (res.isErr()) {
      chai.assert.equal(res.error.source, "installCopilotChat");
    }
    chai.assert.equal(commandStub.callCount, 1);
  });

  it("Install github copilot extension cancel", async () => {
    sandbox.stub(vscode, "version").value("1.88.0");
    const loggerStub = sandbox.stub(VsCodeLogInstance, "error").resolves();
    sandbox
      .stub(vscode.extensions, "getExtension")
      .onFirstCall()
      .returns(undefined)
      .onSecondCall()
      .returns(undefined)
      .onThirdCall()
      .returns({ name: "github.copilot" } as any);
    const commandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();

    sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves("Cancel" as unknown as vscode.MessageItem);

    const res = await handlers.invokeTeamsAgent();

    chai.assert.isTrue(res.isErr());
    if (res.isErr()) {
      chai.assert.equal(res.error.name, "UserCancelError");
    }
    chai.assert.equal(commandStub.callCount, 0);
    chai.assert.equal(loggerStub.callCount, 0);
  });

  it("Verify installation error", async () => {
    clock = sandbox.useFakeTimers();
    sandbox.stub(vscode, "version").value("1.88.0-insiders");
    sandbox.stub(vscode.extensions, "getExtension").returns(undefined);
    const commandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();
    sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves("Install" as unknown as vscode.MessageItem);

    const job = handlers.invokeTeamsAgent();
    await clock.tickAsync(30000);
    const res = await job;

    chai.assert.isTrue(res.isErr());
    if (res.isErr()) {
      chai.assert.equal(res.error.name, "CannotVerifyGithubCopilotChat");
    }
    chai.assert.equal(commandStub.callCount, 1);
  });

  it("invoke Copilot chat error", async () => {
    sandbox.stub(vscode, "version").value("1.88.0");
    sandbox.stub(vscode.extensions, "getExtension").returns({ name: "github.copilot" } as any);
    const commandStub = sandbox
      .stub(vscode.commands, "executeCommand")
      .callsFake(async (command: string) => {
        if (command === "workbench.panel.chat.view.copilot.focus") {
          throw new Error("Install Error");
        } else {
          return {};
        }
      });
    sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves("Install" as unknown as vscode.MessageItem);
    const loggerError = sandbox.stub(VsCodeLogInstance, "error").resolves();

    const res = await handlers.invokeTeamsAgent();

    chai.assert.isTrue(res.isErr());
    if (res.isErr()) {
      chai.assert.equal(res.error.source, "openCopilotChat");
    }
    chai.assert.equal(commandStub.callCount, 1);
    chai.assert.equal(loggerError.callCount, 2);
  });
});
