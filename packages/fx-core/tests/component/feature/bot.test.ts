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
import {
  AzureSolutionQuestionNames,
  NotificationOptionItem,
} from "../../../src/plugins/solution/fx-solution/question";
import { QuestionNames } from "../../../src/plugins/resource/bot/constants";
import { AppServiceOptionItem } from "../../../src/plugins/resource/bot/question";
import Container from "typedi";
describe("Bot Feature", () => {
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
  const manifest = {} as TeamsAppManifest;
  beforeEach(() => {
    sandbox.stub(tools.ui, "showMessage").resolves(ok("Confirm"));
    sandbox.stub(templatesAction, "scaffoldFromTemplates").resolves();
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
    sandbox.stub(utils, "persistBicep").resolves(ok(undefined));
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("add restify notification bot", async () => {
    sandbox.stub(utils.bicepUtils, "persistBiceps").resolves(ok(undefined));
    const inputs: InputsWithProjectPath = {
      projectPath: projectPath,
      platform: Platform.VSCode,
      [AzureSolutionQuestionNames.Features]: NotificationOptionItem.id,
      language: "typescript",
      "app-name": appName,
      [QuestionNames.BOT_HOST_TYPE_TRIGGER]: [AppServiceOptionItem.id],
    };
    const teamsBotComponent = Container.get("teams-bot") as any;
    const addBotRes = await teamsBotComponent.add(context, inputs);
    if (addBotRes.isErr()) {
      console.log(addBotRes.error);
    }
    assert.isTrue(addBotRes.isOk());
    const teamsBot = getComponent(context.projectSetting, ComponentNames.TeamsBot);
    assert.exists(teamsBot);
    assert.equal(teamsBot?.hosting, ComponentNames.AzureWebApp);
    assert.equal(teamsBot?.folder, "bot");
    assert.isTrue(teamsBot?.build);
    assert.deepEqual(teamsBot?.capabilities, ["notification"]);
    const webApp = getComponent(context.projectSetting, ComponentNames.AzureWebApp);
    assert.exists(webApp?.connections);
    if (webApp?.connections) {
      assert.include(webApp.connections, ComponentNames.TeamsBot);
      assert.include(webApp.connections, ComponentNames.Identity);
      assert.equal(webApp.connections.length, 2);
    }
    const botService = getComponent(context.projectSetting, ComponentNames.BotService);
    assert.exists(botService);
    assert.isTrue(botService?.provision);
  });
});
