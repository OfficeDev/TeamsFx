// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ok } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import * as fs from "fs-extra";
import * as path from "path";
import * as sinon from "sinon";

import { envUtil, metadataUtil, pathUtils } from "@microsoft/teamsfx-core";
import { Uri } from "vscode";
import * as commonUtils from "../../src/debug/commonUtils";
import * as globalVariables from "../../src/globalVariables";

const testDataFolder = path.resolve(__dirname, "test-data");

describe("[debug > commonUtils]", () => {
  beforeEach(async () => {
    await fs.ensureDir(testDataFolder);
    await fs.emptyDir(testDataFolder);
  });

  describe("getV3TeamsAppId", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("returns teamsAppId successfully", async () => {
      sandbox.stub(globalVariables, "workspaceUri").value(Uri.file("test"));
      sandbox.stub(pathUtils, "getYmlFilePath");
      sandbox.stub(metadataUtil, "parse").resolves(
        ok({
          provision: {
            driverDefs: [
              {
                uses: "teamsApp/create",
                writeToEnvironmentFile: {
                  teamsAppId: "TeamsAppId",
                },
              },
            ],
          },
        } as any)
      );
      sandbox.stub(envUtil, "readEnv").resolves(
        ok({
          TeamsAppId: "testId",
        } as any)
      );

      const result = await commonUtils.getV3TeamsAppId("testProjectPath", "test");

      chai.expect(result).equals("testId");
    });
  });

  describe("isTestToolEnabledProject", () => {
    const sandbox = sinon.createSandbox();

    afterEach(async () => {
      sandbox.restore();
    });

    it("test tool yaml exist", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(true);
      const res = commonUtils.isTestToolEnabledProject("testPath");
      chai.assert.isTrue(res);
    });

    it("test tool yaml not exist", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(false);
      const res = commonUtils.isTestToolEnabledProject("testPath");
      chai.assert.isFalse(res);
    });
  });
});
