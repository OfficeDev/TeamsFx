// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// Use require so we can mock it
import * as os from "os";
import * as path from "path";
import fs from "fs-extra";
import * as tmp from "tmp";
import "mocha";
import * as sinon from "sinon";
import * as chai from "chai";
import mockFs from "mock-fs";
import { DepsType } from "../../../../src/common/deps-checker/depsChecker";
import { CheckerFactory } from "../../../../src/common/deps-checker/checkerFactory";
import { VxTestAppChecker } from "../../../../src/common/deps-checker/internal/vxTestAppChecker";
import { isMacOS, isWindows } from "../../../../src/common/deps-checker/util";

describe("VxTestAppChecker E2E Test", async () => {
  const fakeProjectPath = "fake project path";
  const vxTestAppExecutableName = isWindows()
    ? "video-extensibility-test-app.exe"
    : isMacOS()
    ? "video-extensibility-test-app.app"
    : "video-extensibility-test-app";
  let sandbox: sinon.SinonSandbox;

  beforeEach(() => {
    sandbox = sinon.createSandbox();
  });

  afterEach(() => {
    sandbox.restore();
    mockFs.restore();
  });

  it("VxTestApp already installed", async function () {
    mockFs({
      [path.join(
        fakeProjectPath,
        ".tools",
        "video-extensibility-test-app",
        vxTestAppExecutableName
      )]: "",
    });
    const checker = new VxTestAppChecker();
    const info = await checker.getInstallationInfo({
      projectPath: fakeProjectPath,
      version: "1.0.4",
    });

    chai.assert.isTrue(info.isInstalled);
  });

  it("VxTestApp not installed", async function () {
    // Currently vxTestApp only publishes Windows bits
    if (!isWindows()) {
      return;
    }

    const checker = CheckerFactory.createChecker(DepsType.VxTestApp);

    let tmpDir;
    try {
      tmpDir = tmp.dirSync({ unsafeCleanup: true });

      const res = await checker.resolve({ projectPath: tmpDir.name, version: "1.0.4" });
      chai.assert.isTrue(res.isInstalled);
      chai.assert.isTrue(
        fs.pathExistsSync(
          path.resolve(tmpDir.name, ".tools/video-extensibility-test-app", vxTestAppExecutableName)
        )
      );
      const stat = fs.lstatSync(path.resolve(tmpDir.name, ".tools/video-extensibility-test-app"));
      chai.assert.isTrue(stat.isSymbolicLink());
      chai.assert.isTrue(
        fs.pathExistsSync(
          path.resolve(
            os.homedir(),
            ".fx/bin/video-extensibility-test-app/1.0.4",
            vxTestAppExecutableName
          )
        )
      );
    } finally {
      tmpDir?.removeCallback();
    }
  });
});
