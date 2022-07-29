// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  InputsWithProjectPath,
  ok,
  Platform,
  ProjectSettingsV3,
  TeamsAppManifest,
} from "@microsoft/teamsfx-api";
import * as templatesAction from "../../../src/common/template-utils/templatesActions";
import * as manifestUtils from "../../../src/component/resource/appManifest/utils";
import * as projectSettingsLoader from "../../../src/core/middleware/projectSettingsLoader";
import { assert } from "chai";
import "mocha";
import * as os from "os";
import * as path from "path";
import fs from "fs-extra";
import { createSandbox } from "sinon";
import * as utils from "../../../src/component/utils";
import { getComponent } from "../../../src/component/workflow";
import { setTools } from "../../../src/core/globalVars";
import { MockTools, randomAppName } from "../../core/utils";
import "../../../src/component/core";
import { environmentManager } from "../../../src/core/environment";
import { ComponentNames } from "../../../src/component/constants";
import * as aadManifest from "../../../src/core/generateAadManifestTemplate";
import Container from "typedi";
describe("Tab Feature", () => {
  const sandbox = createSandbox();
  const tools = new MockTools();
  setTools(tools);
  const appName = `unittest${randomAppName()}`;
  const projectPath = path.join(os.homedir(), "TeamsApps", appName);
  const context = utils.createContextV3();
  const projectSetting: ProjectSettingsV3 = {
    appName: "",
    projectId: "",
    programmingLanguage: "typescript",
    components: [],
  };
  context.projectSetting = projectSetting;
  const manifest = {} as TeamsAppManifest;
  beforeEach(() => {
    sandbox.stub(tools.ui, "showMessage").resolves(ok("Confirm"));
    sandbox.stub(manifestUtils, "readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "writeAppManifest").resolves();
    sandbox.stub(projectSettingsLoader, "loadProjectSettings").resolves(ok(projectSetting));
    sandbox.stub(fs, "readJson").resolves({});
    sandbox.stub(fs, "writeJSON").resolves();
    sandbox.stub(fs, "writeJson").resolves();
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "copyFile").resolves();
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "appendFile").resolves();
    sandbox.stub(fs, "writeFile").resolves();
    sandbox.stub(fs, "ensureDirSync").returns();
    sandbox.stub(fs, "readdirSync").returns([]);
    sandbox.stub(fs, "appendFileSync").returns();
    sandbox.stub(fs, "writeFileSync").returns();
    sandbox.stub(environmentManager, "listRemoteEnvConfigs").resolves(ok(["dev"]));
    sandbox.stub(aadManifest, "generateAadManifestTemplate").resolves();
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("add react tab", async () => {
    sandbox.stub(templatesAction, "scaffoldFromTemplates").resolves();
    sandbox.stub(utils, "persistBicep").resolves(ok(undefined));

    const inputs: InputsWithProjectPath = {
      projectPath: projectPath,
      platform: Platform.VSCode,
      language: "typescript",
      "app-name": appName,
    };
    const component = Container.get("teams-tab") as any;
    const addTabRes = await component.add(context, inputs);
    if (addTabRes.isErr()) {
      console.log(addTabRes.error);
    }
    assert.isTrue(addTabRes.isOk());
    const teamsTab = getComponent(context.projectSetting, ComponentNames.TeamsTab);
    assert.exists(teamsTab);
    assert.equal(teamsTab?.hosting, ComponentNames.AzureStorage);
    assert.equal(teamsTab?.folder, "tabs");
    assert.isTrue(teamsTab?.build);
    const storage = getComponent(context.projectSetting, ComponentNames.AzureStorage);
    assert.exists(storage);
  });

  it("add react tab twice", async () => {
    sandbox.stub(templatesAction, "scaffoldFromTemplates").rejects();
    sandbox.stub(utils, "persistBicep").rejects();

    const inputs: InputsWithProjectPath = {
      projectPath: projectPath,
      platform: Platform.VSCode,
      language: "typescript",
      "app-name": appName,
    };
    const component = Container.get("teams-tab") as any;
    const addTabRes = await component.add(context, inputs);
    if (addTabRes.isErr()) {
      console.log(addTabRes.error);
    }
    assert.isTrue(addTabRes.isOk());
    const teamsTab = context.projectSetting.components.filter(
      (component) => component.name === ComponentNames.TeamsTab
    );
    assert.equal(teamsTab.length, 1);
    const storage = context.projectSetting.components.filter(
      (component) => component.name === ComponentNames.AzureStorage
    );
    assert.equal(storage.length, 1);
  });
});
