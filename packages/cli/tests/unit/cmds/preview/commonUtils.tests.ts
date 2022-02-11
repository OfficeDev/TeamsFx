import { IProgressHandler, err, ok, returnUserError } from "@microsoft/teamsfx-api";
import sinon from "sinon";
import {
  createTaskStartCb,
  createTaskStopCb,
  getAutomaticNpmInstallSetting,
} from "../../../../src/cmds/preview/commonUtils";
import { expect } from "../../utils";
import { UserSettings } from "../../../../src/userSetttings";
import { cliSource } from "../../../../src/constants";

describe("commonUtils", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });

  describe("createTaskStartCb", () => {
    it("happy path", async () => {
      const progressHandler = sinon.createStubInstance(MockProgressHandler);
      const taskStartCallback = createTaskStartCb(progressHandler, "start message");
      await taskStartCallback("start", true);
      expect(progressHandler.start.calledOnce).to.be.true;
    });
  });
  describe("createTaskStopCb", () => {
    it("happy path", async () => {
      const progressHandler = sinon.createStubInstance(MockProgressHandler);
      const taskStopCallback = createTaskStopCb(progressHandler);
      await taskStopCallback("stop", true, {
        command: "command",
        success: true,
        stdout: [],
        stderr: [],
        exitCode: null,
      });
      expect(progressHandler.end.calledOnce).to.be.true;
    });
  });

  describe("getAutomaticNpmInstallSetting", () => {
    const automaticNpmInstallOption = "automatic-npm-install";

    afterEach(() => {
      sinon.restore();
    });

    it("on", () => {
      sinon.stub(UserSettings, "getConfigSync").returns(
        ok({
          [automaticNpmInstallOption]: "on",
        })
      );
      expect(getAutomaticNpmInstallSetting()).to.be.true;
    });

    it("off", () => {
      sinon.stub(UserSettings, "getConfigSync").returns(
        ok({
          [automaticNpmInstallOption]: "off",
        })
      );
      expect(getAutomaticNpmInstallSetting()).to.be.false;
    });

    it("others", () => {
      sinon.stub(UserSettings, "getConfigSync").returns(
        ok({
          [automaticNpmInstallOption]: "others",
        })
      );
      expect(getAutomaticNpmInstallSetting()).to.be.true;
    });

    it("none", () => {
      sinon.stub(UserSettings, "getConfigSync").returns(ok({}));
      expect(getAutomaticNpmInstallSetting()).to.be.false;
    });

    it("getConfigSync error", () => {
      const error = returnUserError(new Error("Test"), cliSource, "Test");
      sinon.stub(UserSettings, "getConfigSync").returns(err(error));
      expect(getAutomaticNpmInstallSetting()).to.be.false;
    });

    it("getConfigSync exception", () => {
      sinon.stub(UserSettings, "getConfigSync").throws("Test");
      expect(getAutomaticNpmInstallSetting()).to.be.false;
    });
  });
});

class MockProgressHandler implements IProgressHandler {
  start(detail?: string): Promise<void> {
    return Promise.resolve();
  }
  next(detail?: string): Promise<void> {
    return Promise.resolve();
  }
  end(success: boolean): Promise<void> {
    return Promise.resolve();
  }
}
