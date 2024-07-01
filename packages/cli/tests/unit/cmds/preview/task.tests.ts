// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as sinon from "sinon";
import { EventEmitter } from "events";
import cp from "child_process";
import { Readable } from "stream";
import { Task } from "../../../../src/cmds/preview/task";
import { expect } from "../../utils";

function delay(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

describe("Task", () => {
  describe("wait", () => {
    let sandbox: sinon.SinonSandbox;

    afterEach(() => {
      sandbox.restore();
    });

    it("happy path", async () => {
      sandbox = sinon.createSandbox();
      const startCallback = sinon.stub().resolves();
      const stopCallback = sinon.stub().resolves(null);
      const spawnEvent = <cp.ChildProcess>new EventEmitter();
      spawnEvent.stdout = <Readable>new EventEmitter();
      spawnEvent.stderr = <Readable>new EventEmitter();
      sandbox.stub(cp, "spawn").callsFake(() => {
        return spawnEvent;
      });

      const task = new Task("taskTitle", false, "command");
      const promise = task.wait(startCallback, stopCallback);

      await delay(10);

      spawnEvent.stdout.emit("data", "stdout1");
      spawnEvent.stdout.emit("data", "stdout2");
      spawnEvent.stderr.emit("data", "stderr1");
      spawnEvent.stderr.emit("data", "stderr2");
      spawnEvent.emit("exit", 0);

      const resultRes = await promise;

      expect(resultRes.isOk()).to.be.true;
      const result = (resultRes as any).value;
      expect(result.success).to.be.true;
      expect(result.stdout).to.deep.equals(["stdout1", "stdout2"]);
      expect(result.stderr).to.deep.equals(["stderr1", "stderr2"]);
      expect(result.exitCode).to.equals(0);
    });
  });

  describe("waitFor", () => {
    let sandbox: sinon.SinonSandbox;

    afterEach(() => {
      sandbox.restore();
    });

    it("happy path: match stdout", async () => {
      sandbox = sinon.createSandbox();
      const startCallback = sinon.stub().resolves();
      const stopCallback = sinon.stub().resolves(null);
      const spawnEvent = <cp.ChildProcess>new EventEmitter();
      spawnEvent.stdout = <Readable>new EventEmitter();
      spawnEvent.stderr = <Readable>new EventEmitter();
      sandbox.stub(cp, "spawn").callsFake(() => {
        return spawnEvent;
      });

      const task = new Task("taskTitle", true, "command");
      const promise = task.waitFor(new RegExp("started", "i"), startCallback, stopCallback);

      await delay(5);

      spawnEvent.stdout.emit("data", "stdout1");
      spawnEvent.stdout.emit("data", "stdout2");
      spawnEvent.stdout.emit("data", "xxx started xxx");
      spawnEvent.stderr.emit("data", "stderr1");
      spawnEvent.stderr.emit("data", "stderr2");

      const resultRes = await promise;

      expect(resultRes.isOk()).to.be.true;
      const result = (resultRes as any).value;
      expect(result.success).to.be.true;
      expect(result.stdout).to.deep.equals(["stdout1", "stdout2", "xxx started xxx"]);
      expect(result.stderr).to.deep.equals(["stderr1", "stderr2"]);
      expect(result.exitCode).to.equals(null);
    });

    it("happy path: match stderr", async () => {
      sandbox = sinon.createSandbox();
      const startCallback = sinon.stub().resolves();
      const stopCallback = sinon.stub().resolves(null);
      const spawnEvent = <cp.ChildProcess>new EventEmitter();
      spawnEvent.stdout = <Readable>new EventEmitter();
      spawnEvent.stderr = <Readable>new EventEmitter();
      sandbox.stub(cp, "spawn").callsFake(() => {
        return spawnEvent;
      });

      const task = new Task("taskTitle", true, "command");
      const promise = task.waitFor(new RegExp("started", "i"), startCallback, stopCallback);

      await delay(5);

      spawnEvent.stdout.emit("data", "stdout1");
      spawnEvent.stdout.emit("data", "stdout2");
      spawnEvent.stderr.emit("data", "stderr1");
      spawnEvent.stderr.emit("data", "stderr2");
      spawnEvent.stderr.emit("data", "xxx started xxx");

      const resultRes = await promise;

      expect(resultRes.isOk()).to.be.true;
      const result = (resultRes as any).value;
      expect(result.success).to.be.false;
      expect(result.stdout).to.deep.equals(["stdout1", "stdout2"]);
      expect(result.stderr).to.deep.equals(["stderr1", "stderr2", "xxx started xxx"]);
      expect(result.exitCode).to.equals(null);
    });

    it("timeout", async () => {
      sandbox = sinon.createSandbox();
      const startCallback = sinon.stub().resolves();
      const stopCallback = sinon.stub().resolves(null);
      const spawnEvent = <cp.ChildProcess>new EventEmitter();
      spawnEvent.stdout = <Readable>new EventEmitter();
      spawnEvent.stderr = <Readable>new EventEmitter();
      sandbox.stub(cp, "spawn").callsFake(() => {
        return spawnEvent;
      });

      const task = new Task("taskTitle", true, "command");
      const promise = task.waitFor(new RegExp("started", "i"), startCallback, stopCallback, 30);

      await delay(5);

      spawnEvent.stdout.emit("data", "stdout1");
      spawnEvent.stdout.emit("data", "stdout2");
      spawnEvent.stderr.emit("data", "stderr1");
      spawnEvent.stderr.emit("data", "stderr2");

      const resultRes = await promise;

      expect(resultRes.isOk()).to.be.true;
      const result = (resultRes as any).value;
      expect(result.success).to.be.false;
      expect(result.stdout).to.deep.equals(["stdout1", "stdout2"]);
      expect(result.stderr).to.deep.equals(["stderr1", "stderr2"]);
      expect(result.exitCode).to.equals(null);
    });
  });
});
