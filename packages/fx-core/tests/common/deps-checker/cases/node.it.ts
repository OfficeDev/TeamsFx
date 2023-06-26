// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as chai from "chai";
import "mocha";
import semver from "semver";
import {
  LtsNodeChecker,
  NodeChecker,
  ProjectNodeChecker,
} from "../../../../src/common/deps-checker/internal/nodeChecker";
import * as uuid from "uuid";
import * as fs from "fs-extra";
import * as path from "path";

const ltsNodeRange = "16 || 18";

describe("NodeChecker E2E Test", async () => {
  let baseFolder: string | undefined = undefined;
  beforeEach(async function () {
    baseFolder = path.join(__dirname, "func-e2e-test-data", uuid.v4().substring(0, 6));
  });

  afterEach(async function () {
    if (baseFolder) {
      await fs.remove(baseFolder);
    }
  });

  [
    {
      nodeChecker: new ProjectNodeChecker(),
      type: "projectNodeChecker",
    },
    {
      nodeChecker: new LtsNodeChecker(),
      type: "ltsNodeChecker",
    },
  ].forEach((data) => {
    it(`${data.type}, Node supported version is installed`, async function () {
      const nodeVersion = await NodeChecker.getInstalledNodeVersion();
      if (!nodeVersion || !semver.satisfies(nodeVersion.version, ltsNodeRange)) {
        this.skip();
      }
      if (data.type === "projectNodeChecker") {
        await fs.ensureDir(baseFolder!);
        await fs.writeJson(path.join(baseFolder!, "package.json"), {
          engines: {
            node: ltsNodeRange,
          },
        });
      }
      const res = await data.nodeChecker.resolve({ projectPath: baseFolder! });

      chai.assert.isTrue(res.isInstalled);
      chai.assert.isTrue((await data.nodeChecker.getInstallationInfo()).isInstalled);
      chai.assert.isUndefined(res.error);
    });

    it(`${data.type}, Node not supported version is installed`, async function () {
      const nodeVersion = await NodeChecker.getInstalledNodeVersion();
      if (!nodeVersion || semver.satisfies(nodeVersion.version, ltsNodeRange)) {
        this.skip();
      }
      if (data.type === "projectNodeChecker") {
        await fs.ensureDir(baseFolder!);
        await fs.writeJson(path.join(baseFolder!, "package.json"), {
          engines: {
            node: ltsNodeRange,
          },
        });
      }
      const res = await data.nodeChecker.resolve({ projectPath: baseFolder! });

      chai.assert.isTrue(res.isInstalled);
      chai.assert.isTrue((await data.nodeChecker.getInstallationInfo()).isInstalled);
      chai.assert.isDefined(res.error);
    });

    it(`${data.type}, Node is not installed`, async function () {
      const nodeVersion = await NodeChecker.getInstalledNodeVersion();
      if (!!nodeVersion) {
        this.skip();
      }

      const res = await data.nodeChecker.resolve();
      chai.assert.isFalse(res.isInstalled);
      chai.assert.isFalse((await data.nodeChecker.getInstallationInfo()).isInstalled);
    });
  });

  it(`projectNodeChecker, Node is installed, no package.json`, async function () {
    const nodeVersion = await NodeChecker.getInstalledNodeVersion();
    if (!nodeVersion || semver.satisfies(nodeVersion.version, ltsNodeRange)) {
      this.skip();
    }
    const nodeChecker = new ProjectNodeChecker();
    const res = await nodeChecker.resolve({ projectPath: baseFolder! });

    chai.assert.isTrue(res.isInstalled);
    chai.assert.isTrue((await nodeChecker.getInstallationInfo()).isInstalled);
    chai.assert.isUndefined(res.error);
  });
});
