// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  FxError,
  IQTreeNode,
  Inputs,
  LogProvider,
  Ok,
  OpenAIPluginManifest,
  Platform,
  Result,
  Stage,
  SystemError,
  TeamsAppManifest,
  UserError,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import jsyaml from "js-yaml";
import "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import * as os from "os";
import * as path from "path";
import sinon from "sinon";
import { FxCore, getUuid } from "../../src";
import { FeatureFlagName } from "../../src/common/constants";
import { LaunchHelper } from "../../src/common/m365/launchHelper";
import {
  TeamsfxConfigType,
  TeamsfxVersionState,
  projectTypeChecker,
} from "../../src/common/projectTypeChecker";
import {
  ErrorType,
  ListAPIResult,
  SpecParser,
  SpecParserError,
  ValidationStatus,
  WarningType,
} from "@microsoft/m365-spec-parser";
import {
  DriverDefinition,
  DriverInstance,
  ExecutionResult,
  ILifecycle,
  LifecycleName,
  Output,
  UnresolvedPlaceholders,
} from "../../src/component/configManager/interface";
import { YamlParser } from "../../src/component/configManager/parser";
import { coordinator } from "../../src/component/coordinator";
import { UpdateAadAppDriver } from "../../src/component/driver/aad/update";
import * as buildAadManifest from "../../src/component/driver/aad/utility/buildAadManifest";
import { AddWebPartDriver } from "../../src/component/driver/add/addWebPart";
import { DriverContext } from "../../src/component/driver/interface/commonArgs";
import { CreateAppPackageDriver } from "../../src/component/driver/teamsApp/createAppPackage";
import { teamsappMgr } from "../../src/component/driver/teamsApp/teamsappMgr";
import { manifestUtils } from "../../src/component/driver/teamsApp/utils/ManifestUtils";
import { ValidateManifestDriver } from "../../src/component/driver/teamsApp/validate";
import { ValidateAppPackageDriver } from "../../src/component/driver/teamsApp/validateAppPackage";
import "../../src/component/feature/sso";
import * as CopilotPluginHelper from "../../src/component/generator/copilotPlugin/helper";
import { OpenAIPluginManifestHelper } from "../../src/component/generator/copilotPlugin/helper";
import { createDriverContext } from "../../src/component/utils";
import { envUtil } from "../../src/component/utils/envUtil";
import { metadataUtil } from "../../src/component/utils/metadataUtil";
import { pathUtils } from "../../src/component/utils/pathUtils";
import * as collaborator from "../../src/core/collaborator";
import { environmentManager } from "../../src/core/environment";
import { setTools } from "../../src/core/globalVars";
import * as projectMigratorV3 from "../../src/core/middleware/projectMigratorV3";
import {
  FileNotFoundError,
  InvalidProjectError,
  MissingEnvironmentVariablesError,
  MissingRequiredInputError,
} from "../../src/error/common";
import { NoNeedUpgradeError } from "../../src/error/upgrade";
import {
  CapabilityOptions,
  QuestionNames,
  ScratchOptions,
  questionNodes,
} from "../../src/question";
import { HubOptions } from "../../src/question/other";
import { validationUtils } from "../../src/ui/validationUtils";
import { MockTools, randomAppName } from "./utils";
import { ValidateWithTestCasesDriver } from "../../src/component/driver/teamsApp/validateTestCases";
import { pluginManifestUtils } from "../../src/component/driver/teamsApp/utils/PluginManifestUtils";

const tools = new MockTools();

describe("Core basic APIs", () => {
  const sandbox = sinon.createSandbox();
  const appName = randomAppName();
  beforeEach(() => {
    setTools(tools);
  });
  afterEach(async () => {
    sandbox.restore();
  });

  it("deploy aad manifest happy path with param", async () => {
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    // sandbox.stub(UpdateAadAppDriver.prototype, "run").resolves(new Ok(new Map()));
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.AppName]: appName,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Capabilities]: ["Tab", "TabSSO"],
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AadAppManifestFilePath]: path.join(os.tmpdir(), appName, "aad.manifest.json"),
      [QuestionNames.TargetEnvName]: "dev",
      stage: Stage.deployAad,
      projectPath: path.join(os.tmpdir(), appName),
    };

    const runSpy = sandbox.spy(UpdateAadAppDriver.prototype, "execute");
    await core.deployAadManifest(inputs);
    sandbox.assert.calledOnce(runSpy);
    assert.isNotNull(runSpy.getCall(0).args[0]);
    assert.strictEqual(
      runSpy.getCall(0).args[0].manifestPath,
      path.join(os.tmpdir(), appName, "aad.manifest.json")
    );
    runSpy.restore();
  });

  it("add web part to SPFx", async () => {
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    const appPath = path.join(os.tmpdir(), appName);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      "spfx-folder": ".\\src",
      "manifest-path": path.join(appPath, "appPackage\\manifest.json"),
      "local-manifest-path": path.join(appPath, "appPackage\\manifest.local.json"),
      "spfx-webpart-name": "helloworld",
      "spfx-install-latest-package": "true",
      "spfx-load-package-version": "loaded",
      stage: Stage.addWebpart,
      projectPath: appPath,
    };

    const runSpy = sandbox.spy(AddWebPartDriver.prototype, "run");
    await core.addWebpart(inputs);
    sandbox.assert.calledOnce(runSpy);
    runSpy.restore();
  });

  it("add web part to SPFx - CLI help", async () => {
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    const appPath = path.join(os.tmpdir(), appName);
    const inputs: Inputs = {
      platform: Platform.CLI_HELP,
      [QuestionNames.Folder]: os.tmpdir(),
      "spfx-folder": ".\\src",
      "manifest-path": path.join(appPath, "appPackage\\manifest.json"),
      "local-manifest-path": path.join(appPath, "appPackage\\manifest.local.json"),
      "spfx-webpart-name": "helloworld",
      "spfx-install-latest-package": "true",
      "spfx-load-package-version": "loaded",
      stage: Stage.addWebpart,
      projectPath: appPath,
    };

    const runSpy = sandbox.spy(AddWebPartDriver.prototype, "run");
    await core.addWebpart(inputs);
    sandbox.assert.calledOnce(runSpy);
    runSpy.restore();
  });

  it("add web part to SPFx with empty .yo-rc.json", async () => {
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    const appPath = path.join(os.tmpdir(), appName);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      "spfx-folder": ".\\src",
      "manifest-path": path.join(appPath, "appPackage\\manifest.json"),
      "local-manifest-path": path.join(appPath, "appPackage\\manifest.local.json"),
      "spfx-webpart-name": "helloworld",
      "spfx-install-latest-package": "true",
      "spfx-load-package-version": "loaded",
      stage: Stage.addWebpart,
      projectPath: appPath,
    };

    sandbox.stub(fs, "pathExists").callsFake(async (directory: string) => {
      if (directory.includes(path.join("webparts", "helloworld"))) {
        return false;
      }
      return true;
    });
    sandbox.stub(fs, "readJson").resolves({});
    const runSpy = sandbox.stub(AddWebPartDriver.prototype, "run");
    await core.addWebpart(inputs);
    sandbox.assert.calledOnce(runSpy);
    runSpy.restore();
  });

  it("add web part to SPFx with framework", async () => {
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    const appPath = path.join(os.tmpdir(), appName);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      "spfx-folder": ".\\src",
      "manifest-path": path.join(appPath, "appPackage\\manifest.json"),
      "local-manifest-path": path.join(appPath, "appPackage\\manifest.local.json"),
      "spfx-webpart-name": "helloworld",
      "spfx-install-latest-package": "true",
      "spfx-load-package-version": "loaded",
      stage: Stage.addWebpart,
      projectPath: appPath,
    };

    sandbox.stub(fs, "pathExists").callsFake(async (directory: string) => {
      if (directory.includes(path.join("webparts", "helloworld"))) {
        return false;
      }
      return true;
    });
    sandbox.stub(fs, "readJson").resolves({
      "@microsoft/generator-sharepoint": {
        template: "react",
      },
    });
    const runSpy = sandbox.stub(AddWebPartDriver.prototype, "run");
    await core.addWebpart(inputs);
    sandbox.assert.calledOnce(runSpy);
    runSpy.restore();
  });

  it("deploy aad manifest happy path", async () => {
    const promtionOnVSC =
      'Your Microsoft Entra app has been deployed successfully. To view that, click "Learn more"';

    const core = new FxCore(tools);
    const showMessage = sandbox.spy(tools.ui, "showMessage") as unknown as sinon.SinonSpy<
      ["info" | "warn" | "error", string, boolean, ...string[]],
      Promise<Result<string | undefined, FxError>>
    >;
    const openUrl = sandbox.spy(tools.ui, "openUrl");
    const appName = await mockV3Project();
    sandbox
      .stub(UpdateAadAppDriver.prototype, "execute")
      .resolves({ result: new Ok(new Map()), summaries: [] });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.AppName]: appName,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Capabilities]: ["Tab", "TabSSO"],
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AadAppManifestFilePath]: path.join(os.tmpdir(), appName, "aad.manifest.json"),
      env: "dev",
      stage: Stage.deployAad,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const res = await core.deployAadManifest(inputs);
    assert.isTrue(await fs.pathExists(path.join(os.tmpdir(), appName, "build")));
    await deleteTestProject(appName);
    assert.isTrue(res.isOk());
    assert.isTrue(showMessage.called);
    assert.equal(showMessage.getCall(0).args[0], "info");
    assert.equal(showMessage.getCall(0).args[1], promtionOnVSC);
    assert.isFalse(showMessage.getCall(0).args[2]);
    assert.equal(showMessage.getCall(0).args[3], "Learn more");
    assert.isFalse(openUrl.called);
  });
  it("deploy aad manifest happy path with click learn more", async () => {
    const core = new FxCore(tools);
    sandbox.stub(tools.ui, "showMessage").resolves(ok("Learn more"));
    sandbox.stub(tools.ui, "openUrl").resolves(ok(true));
    const appName = await mockV3Project();
    sandbox
      .stub(UpdateAadAppDriver.prototype, "execute")
      .resolves({ result: new Ok(new Map()), summaries: [] });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.AppName]: appName,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Capabilities]: ["Tab", "TabSSO"],
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AadAppManifestFilePath]: path.join(os.tmpdir(), appName, "aad.manifest.json"),
      env: "dev",
      stage: Stage.deployAad,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const res = await core.deployAadManifest(inputs);
    assert.isTrue(res.isOk());
    if (res.isErr()) console.error(res.error);
    assert.isTrue(await fs.pathExists(path.join(os.tmpdir(), appName, "build")));
    await deleteTestProject(appName);
  });

  it("deploy aad manifest happy path without click learn more", async () => {
    const core = new FxCore(tools);
    sandbox.stub(tools.ui, "showMessage").resolves(err(new UserError("test", "test", "test")));
    sandbox.stub(tools.ui, "openUrl").resolves(ok(true));
    const appName = await mockV3Project();
    sandbox
      .stub(UpdateAadAppDriver.prototype, "execute")
      .resolves({ result: new Ok(new Map()), summaries: [] });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.AppName]: appName,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Capabilities]: ["Tab", "TabSSO"],
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AadAppManifestFilePath]: path.join(os.tmpdir(), appName, "aad.manifest.json"),
      env: "dev",
      stage: Stage.deployAad,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const res = await core.deployAadManifest(inputs);
    assert.isTrue(res.isOk());
    if (res.isErr()) console.error(res.error);
    assert.isTrue(await fs.pathExists(path.join(os.tmpdir(), appName, "build")));
    await deleteTestProject(appName);
  });
  it("deploy aad manifest happy path on cli", async () => {
    const core = new FxCore(tools);
    const showMessage = sandbox.spy(tools.ui, "showMessage") as unknown as sinon.SinonSpy<
      ["info" | "warn" | "error", string, boolean, ...string[]],
      Promise<Result<string | undefined, FxError>>
    >;
    const appName = await mockV3Project();
    sandbox
      .stub(UpdateAadAppDriver.prototype, "execute")
      .resolves({ result: new Ok(new Map()), summaries: [] });
    const inputs: Inputs = {
      platform: Platform.CLI,
      [QuestionNames.AppName]: appName,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Capabilities]: ["Tab", "TabSSO"],
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AadAppManifestFilePath]: path.join(os.tmpdir(), appName, "aad.manifest.json"),
      env: "dev",
      stage: Stage.deployAad,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const res = await core.deployAadManifest(inputs);
    await deleteTestProject(appName);
    assert.isTrue(showMessage.calledOnce);
    assert.equal(showMessage.getCall(0).args[0], "info");
    assert.equal(
      showMessage.getCall(0).args[1],
      "Your Microsoft Entra app has been updated successfully."
    );
    assert.isFalse(showMessage.getCall(0).args[2]);
    assert.isTrue(res.isOk());
  });

  it("deploy aad manifest return err", async () => {
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    const appManifestPath = path.join(os.tmpdir(), appName, "aad.manifest.json");
    sandbox.stub(environmentManager, "listAllEnvConfigs").resolves(ok(["dev", "local"]));
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.AppName]: appName,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Capabilities]: ["Tab", "TabSSO"],
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AadAppManifestFilePath]: appManifestPath,
      env: "dev",
      stage: Stage.deployAad,
      projectPath: path.join(os.tmpdir(), appName),
    };
    sandbox
      .stub(UpdateAadAppDriver.prototype, "execute")
      .throws(new UserError("error name", "fake_error", "fake_err_msg"));
    const errMsg = `AAD manifest doesn't exist in ${appManifestPath}, please use the CLI to specify an AAD manifest to deploy.`;
    const res = await core.deployAadManifest(inputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.strictEqual(res.error.message, "fake_err_msg");
    }
  });

  it("deploy aad manifest with missing env err", async () => {
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    const appManifestPath = path.join(os.tmpdir(), appName, "aad.manifest.json");
    sandbox.stub(environmentManager, "listAllEnvConfigs").resolves(ok([""]));
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.AppName]: appName,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Capabilities]: ["Tab", "TabSSO"],
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AadAppManifestFilePath]: appManifestPath,
      env: undefined,
      stage: Stage.deployAad,
      projectPath: path.join(os.tmpdir(), appName),
    };
    sandbox.stub(UpdateAadAppDriver.prototype, "execute").resolves({
      result: err(
        new MissingEnvironmentVariablesError(
          "aadApp/update",
          "AAD_APP_OBJECT_ID",
          "fake path",
          "https://fake-help-link"
        )
      ),
      summaries: [],
    });
    const res = await core.deployAadManifest(inputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      // Cannot assert the full message because the mocked code can't get correct env file path
      assert.include(
        res.error.message,
        "The program cannot proceed as the following environment variables are missing: 'AAD_APP_OBJECT_ID', which are required for file: fake path. Make sure the required variables are set either by editing the .env file"
      );
      assert.include(
        res.error.message,
        "If you are developing with a new project created with Teams Toolkit, running provision or debug will register correct values for these environment variables"
      );
    }
  });

  it("deploy aad manifest not exist", async () => {
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    const appManifestPath = path.join(os.tmpdir(), appName, "aad.manifest.json");
    await fs.remove(appManifestPath);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.AppName]: appName,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Capabilities]: ["Tab", "TabSSO"],
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AadAppManifestFilePath]: path.join(os.tmpdir(), appName, "aad.manifest.json"),
      env: "dev",
      stage: Stage.deployAad,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const res = await core.deployAadManifest(inputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof FileNotFoundError);
    }
    await deleteTestProject(appName);
  });

  it("phantomMigrationV3 happy path", async () => {
    const core = new FxCore(tools);
    const appName = await mockV2Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir(), appName),
      skipUserConfirm: true,
    };
    const res = await core.phantomMigrationV3(inputs);
    assert.isTrue(res.isOk());
    await deleteTestProject(appName);
  });

  it("phantomMigrationV3 return error for invalid V2 project", async () => {
    sandbox.stub(projectMigratorV3, "checkActiveResourcePlugins").resolves(false);

    const core = new FxCore(tools);
    const appName = await mockV2Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir(), appName),
      skipUserConfirm: true,
    };
    const res = await core.phantomMigrationV3(inputs);
    assert.isTrue(res.isErr());
    assert.isTrue(res._unsafeUnwrapErr().message.includes(new InvalidProjectError().message));
    await deleteTestProject(appName);
  });

  it("phantomMigrationV3 return error for non-project", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir()),
      skipUserConfirm: true,
    };
    const res = await core.phantomMigrationV3(inputs);
    assert.isTrue(res.isErr());
    assert.isTrue(res._unsafeUnwrapErr().message.includes(new InvalidProjectError().message));
  });

  it("phantomMigrationV3 return error for V5 project", async () => {
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const res = await core.phantomMigrationV3(inputs);
    assert.isTrue(res.isErr());
    assert.isTrue(res._unsafeUnwrapErr().message.includes(new NoNeedUpgradeError().message));
    await deleteTestProject(appName);
  });

  it("permission v3", async () => {
    let res;
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir(), appName),
    };
    sandbox.stub(questionNodes, "grantPermission").returns({ data: { type: "group" } });
    sandbox.stub(questionNodes, "listCollaborator").returns({ data: { type: "group" } });
    sandbox.stub(collaborator, "listCollaborator").resolves(ok(undefined as any));
    sandbox.stub(collaborator, "checkPermission").resolves(ok(undefined as any));
    sandbox.stub(collaborator, "grantPermission").resolves(ok(undefined as any));

    res = await core.listCollaborator(inputs);
    assert.isTrue(res.isOk());
    res = await core.checkPermission(inputs);
    assert.isTrue(res.isOk());
    res = await core.grantPermission(inputs);
    assert.isTrue(res.isOk());
  });

  it("buildAadManifest method should exist", async () => {
    const restore = mockedEnv({
      TEAMSFX_DEBUG_TEMPLATE: "true", // workaround test failure that when local template not released to GitHub
      NODE_ENV: "development", // workaround test failure that when local template not released to GitHub
      AAD_APP_OBJECT_ID: getUuid(),
      AAD_APP_CLIENT_ID: getUuid(),
      TAB_DOMAIN: "fake",
      TAB_ENDPOINT: "fake",
    });

    const originFunc = envUtil.readEnv;
    try {
      envUtil.readEnv = async () => {
        return ok({
          AAD_APP_OBJECT_ID: getUuid(),
          AAD_APP_CLIENT_ID: getUuid(),
          TAB_DOMAIN: "fake",
          TAB_ENDPOINT: "fake",
        });
      };
      const core = new FxCore(tools);
      const appName = await mockV3Project();
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: path.join(os.tmpdir(), appName),
      };
      sandbox.stub(buildAadManifest, "buildAadManifest").resolves({} as any);
      const result = await core.buildAadManifest(inputs);
      assert.isTrue(result.isOk());
    } finally {
      envUtil.readEnv = originFunc;
      restore();
    }
  });

  it("addSso method should exist", async () => {
    const restore = mockedEnv({
      TEAMSFX_DEBUG_TEMPLATE: "true", // workaround test failures when template changed but not release to GitHub alpha template
      NODE_ENV: "development", // workaround test failures when template changed but not release to GitHub alpha template
    });
    try {
      const appName = randomAppName();
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.AppName]: appName,
        [QuestionNames.Scratch]: ScratchOptions.yes().id,
        [QuestionNames.ProgrammingLanguage]: "javascript",
        [QuestionNames.Capabilities]: CapabilityOptions.nonSsoTab().id,
        [QuestionNames.Folder]: os.tmpdir(),
        stage: Stage.create,
        projectPath: path.join(os.tmpdir(), appName, "samples-v3"),
      };
      const res = await core.createProject(inputs);
      const projectPath = inputs.projectPath!;
      assert.isTrue(res.isOk() && res.value.projectPath === projectPath);

      const implement = new FxCore(tools);

      const mockFunc = {
        namespace: "mock namespace",
        method: "addSso",
      };

      const result = await implement.executeUserTask(mockFunc, inputs);
      assert.isTrue(result.isOk());
    } finally {
      restore();
    }
  });
});

describe("apply yaml template", async () => {
  const tools = new MockTools();
  beforeEach(() => {
    setTools(tools);
  });
  describe("when run with missing input", async () => {
    it("should return error when projectPath is undefined", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: undefined,
      };
      const res = await core.apply(inputs, "", "provision");
      assert.isTrue(
        res.isErr() &&
          res.error.name === "InvalidInput" &&
          res.error.message.includes("projectPath")
      );
    });

    it("should return error when env is undefined", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: undefined,
      };
      const res = await core.apply(inputs, "", "provision");
      assert.isTrue(
        res.isErr() && res.error.name === "InvalidInput" && res.error.message.includes("env")
      );
    });
  });

  describe("when readEnv returns error", async () => {
    const sandbox = sinon.createSandbox();

    const mockedError = new SystemError("mockedSource", "mockedError", "mockedMessage");

    before(() => {
      sandbox.stub(envUtil, "readEnv").resolves(err(mockedError));
    });

    after(() => {
      sandbox.restore();
    });

    it("should return error too", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: "dev",
      };
      const res = await core.apply(inputs, "./", "provision");
      assert.isTrue(res.isErr() && res.error.name === "mockedError");
    });
  });

  describe("when YamlParser returns error", async () => {
    const sandbox = sinon.createSandbox();

    const mockedError = new SystemError("mockedSource", "mockedError", "mockedMessage");

    before(() => {
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      sandbox.stub(YamlParser.prototype, "parse").resolves(err(mockedError));
    });

    after(() => {
      sandbox.restore();
    });

    it("should return error too", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: "dev",
      };
      const res = await core.apply(inputs, "./", "provision");
      assert.isTrue(res.isErr() && res.error.name === "mockedError");
    });
  });

  describe("when running against an empty yaml file", async () => {
    const sandbox = sinon.createSandbox();

    before(() => {
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      sandbox.stub(YamlParser.prototype, "parse").resolves(ok({ version: "1.0.0" }));
    });

    after(() => {
      sandbox.restore();
    });

    it("should return ok", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: "dev",
      };
      const res = await core.apply(inputs, "./", "provision");
      assert.isTrue(res.isOk());
    });
  });

  describe("when lifecycle returns error", async () => {
    const sandbox = sinon.createSandbox();
    const mockedError = new SystemError("mockedSource", "mockedError", "mockedMessage");

    class MockedProvision implements ILifecycle {
      name: LifecycleName = "provision";
      driverDefs: DriverDefinition[] = [];
      public async run(ctx: DriverContext): Promise<Result<Output, FxError>> {
        return err(mockedError);
      }

      public resolvePlaceholders(): UnresolvedPlaceholders {
        return [];
      }

      public async execute(ctx: DriverContext): Promise<ExecutionResult> {
        return {
          result: err({
            kind: "Failure",
            error: mockedError,
          }),
          summaries: [],
        };
      }

      public resolveDriverInstances(log: LogProvider): Result<DriverInstance[], FxError> {
        return ok([]);
      }
    }

    before(() => {
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      sandbox.stub(YamlParser.prototype, "parse").resolves(
        ok({
          version: "1.0.0",
          provision: new MockedProvision(),
        })
      );
    });

    after(() => {
      sandbox.restore();
    });

    it("should return error", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: "dev",
      };
      const res = await core.apply(inputs, "./", "provision");
      assert.isTrue(res.isErr() && res.error.name === "mockedError");
    });
  });
  describe("runLifecycle", async () => {
    const sandbox = sinon.createSandbox();

    const mockedError = new SystemError("mockedSource", "mockedError", "mockedMessage");
    class MockedProvision implements ILifecycle {
      name: LifecycleName = "provision";
      driverDefs: DriverDefinition[] = [];
      public async run(ctx: DriverContext): Promise<Result<Output, FxError>> {
        return err(mockedError);
      }

      public resolvePlaceholders(): UnresolvedPlaceholders {
        return [];
      }

      public async execute(ctx: DriverContext): Promise<ExecutionResult> {
        return {
          result: ok(new Map()),
          summaries: [],
        };
      }

      public resolveDriverInstances(log: LogProvider): Result<DriverInstance[], FxError> {
        return ok([]);
      }
    }

    afterEach(() => {
      sandbox.restore();
    });

    it("happy", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: "dev",
      };
      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      const context = createDriverContext(inputs);
      const lifecycle = new MockedProvision();
      const res = await core.runLifecycle(lifecycle, context, "dev");
      assert.isTrue(res.isOk());
    });

    it("partial success", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: "dev",
      };
      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      const lifecycle = new MockedProvision();
      sandbox.stub(lifecycle, "execute").resolves({
        result: err({
          kind: "PartialSuccess",
          env: new Map(),
          reason: {
            kind: "UnresolvedPlaceholders",
            failedDriver: { uses: "t", with: {} },
            unresolvedPlaceHolders: ["TEST_VAR"],
          },
        }),
        summaries: [],
      });
      const context = createDriverContext(inputs);
      const res = await core.runLifecycle(lifecycle, context, "dev");
      assert.isTrue(res.isOk());
    });

    it("DriverError", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: "dev",
      };
      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      const lifecycle = new MockedProvision();
      sandbox.stub(lifecycle, "execute").resolves({
        result: err({
          kind: "PartialSuccess",
          env: new Map(),
          reason: {
            kind: "DriverError",
            failedDriver: { uses: "t", with: {} },
            error: mockedError,
          },
        }),
        summaries: [],
      });
      const context = createDriverContext(inputs);
      const res = await core.runLifecycle(lifecycle, context, "dev");
      assert.isTrue(res.isErr());
    });
  });
});

async function mockV3Project(): Promise<string> {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  // await fs.move(path.join(__dirname, "../sampleV3"), path.join(os.tmpdir(), appName));
  await fs.copy(path.join(__dirname, "../samples/sampleV3/"), path.join(projectPath));
  return appName;
}

async function mockV2Project(): Promise<string> {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  await fs.copy(
    path.join(__dirname, "../core/middleware/testAssets/v3Migration/happyPath"),
    path.join(projectPath)
  );
  return appName;
}

async function deleteTestProject(appName: string) {
  await fs.remove(path.join(os.tmpdir(), appName));
}

describe("createEnvCopyV3", async () => {
  const tools = new MockTools();
  const sandbox = sinon.createSandbox();
  const sourceEnvContent = [
    "# this is a comment",
    "TEAMSFX_ENV=dev",
    "APP_NAME_SUFFIX=dev",
    "",
    "_KEY1=value1",
    "KEY2=value2",
    "SECRET_KEY3=xxxx",
  ];
  const sourceEnvStr = sourceEnvContent.join(os.EOL);

  const writeStreamContent: string[] = [];
  // fs.WriteStream's full interface is too hard to mock. We only use write() and end() so we just mock them here.
  class MockedWriteStream {
    write(chunk: any, callback?: ((error: Error | null | undefined) => void) | undefined): boolean {
      writeStreamContent.push(chunk);
      return true;
    }
    end(): boolean {
      return true;
    }
  }

  beforeEach(() => {
    sandbox.stub(fs, "readFile").resolves(Buffer.from(sourceEnvStr, "utf8"));
    sandbox.stub<any, any>(fs, "createWriteStream").returns(new MockedWriteStream());
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("should create new .env file with desired content", async () => {
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("./env/.env.dev"));
    sandbox.stub(fs, "pathExists").resolves(true);
    const core = new FxCore(tools);
    const res = await core.createEnvCopyV3("newEnv", "dev", "./");
    assert(res.isOk());
    assert(
      writeStreamContent[0] === `${sourceEnvContent[0]}${os.EOL}`,
      "comments should be copied"
    );
    assert(
      writeStreamContent[1] === `TEAMSFX_ENV=newEnv${os.EOL}`,
      "TEAMSFX_ENV's value should be new env name"
    );
    assert(
      writeStreamContent[2] === `APP_NAME_SUFFIX=newEnv${os.EOL}`,
      "APP_NAME_SUFFIX's value should be new env name"
    );
    assert(writeStreamContent[3] === `${os.EOL}`, "empty line should be coped");
    assert(
      writeStreamContent[4] === `_KEY1=${os.EOL}`,
      "key starts with _ should be copied with empty value"
    );
    assert(
      writeStreamContent[5] === `KEY2=${os.EOL}`,
      "key not starts with _ should be copied with empty value"
    );
    assert(
      writeStreamContent[6] === `SECRET_KEY3=${os.EOL}`,
      "key not starts with SECRET_ should be copied with empty value"
    );
  });

  it("should failed case 1", async () => {
    sandbox
      .stub(pathUtils, "getEnvFilePath")
      .onFirstCall()
      .resolves(err(new UserError({})));
    const core = new FxCore(tools);
    const res = await core.createEnvCopyV3("newEnv", "dev", "./");
    assert(res.isErr());
  });

  it("should failed case 2", async () => {
    sandbox
      .stub(pathUtils, "getEnvFilePath")
      .onFirstCall()
      .resolves(ok("./env"))
      .onSecondCall()
      .resolves(err(new UserError({})));
    const core = new FxCore(tools);
    const res = await core.createEnvCopyV3("newEnv", "dev", "./");
    assert(res.isErr());
  });
});

describe("publishInDeveloperPortal", () => {
  const tools = new MockTools();
  const sandbox = sinon.createSandbox();

  before(() => {
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
  });
  afterEach(() => {
    sandbox.restore();
  });

  it("success", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      env: "local",
      projectPath: "project-path",
      platform: Platform.VSCode,
      [QuestionNames.AppPackagePath]: "path",
      ignoreLockByUT: true,
    };
    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(coordinator, "publishInDeveloperPortal").resolves(ok(undefined));
    const res = await core.publishInDeveloperPortal(inputs);

    if (res.isErr()) {
      console.log(res.error);
    }
    assert.isTrue(res.isOk());
  });
});

describe("Teams app APIs", async () => {
  const tools = new MockTools();
  const core = new FxCore(tools);

  afterEach(() => {
    sinon.restore();
  });

  it("validate app package", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppPackageFilePath]: ".\\build\\appPackage\\appPackage.dev.zip",
      [QuestionNames.ValidateMethod]: "validateAgainstAppPackage",
      projectPath: path.join(os.tmpdir(), appName),
    };

    const runSpy = sinon.spy(ValidateAppPackageDriver.prototype, "execute");
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    await core.validateApplication(inputs);
    sinon.assert.calledOnce(runSpy);
  });

  it("validate manifest", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: ".\\appPackage\\manifest.json",
      [QuestionNames.ValidateMethod]: "validateAgainstSchema",
      projectPath: path.join(os.tmpdir(), appName),
    };

    const runSpy = sinon.spy(ValidateManifestDriver.prototype, "execute");
    await core.validateApplication(inputs);
    sinon.assert.calledOnce(runSpy);
  });

  it("validate with test cases", async () => {
    const appName = await mockV3Project();

    const mockedEnvRestore = mockedEnv({
      [FeatureFlagName.AsyncAppValidation]: "true",
    });

    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppPackageFilePath]: ".\\build\\appPackage\\appPackage.dev.zip",
      [QuestionNames.ValidateMethod]: "validateWithTestCases",
      projectPath: path.join(os.tmpdir(), appName),
    };

    const runSpy = sinon.spy(ValidateWithTestCasesDriver.prototype, "execute");
    await core.validateApplication(inputs);
    sinon.assert.calledOnce(runSpy);

    mockedEnvRestore();
  });

  it("create app package", async () => {
    setTools(tools);
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: ".\\appPackage\\manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
      [QuestionNames.OutputZipPathParamName]: ".\\build\\appPackage\\appPackage.dev.zip",
    };

    sinon.stub(process, "platform").value("win32");
    const runStub = sinon
      .stub(CreateAppPackageDriver.prototype, "execute")
      .resolves({ result: ok(new Map()), summaries: [] });
    const showMessageStub = sinon.stub(tools.ui, "showMessage");
    await core.createAppPackage(inputs);
    sinon.assert.calledOnce(runStub);
    sinon.assert.calledOnce(showMessageStub);
  });

  it("publish application", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      projectPath: path.join(os.tmpdir(), appName),
    };

    sinon
      .stub(coordinator, "publish")
      .resolves(err(new SystemError("mockedSource", "mockedError", "mockedMessage")));
    await core.publishApplication(inputs);
  });
});

describe("previewWithManifest", () => {
  const tools = new MockTools();
  const core = new FxCore(tools);

  afterEach(() => {
    sinon.restore();
  });

  it("getManifestV3 error", async () => {
    sinon.stub(manifestUtils, "getManifestV3").resolves(err({ foo: "bar" } as any));
    const appName = await mockV3Project();
    const inputs: Inputs = {
      [QuestionNames.M365Host]: HubOptions.teams().id,
      [QuestionNames.TeamsAppManifestFilePath]: path.join(
        os.tmpdir(),
        appName,
        "appPackage",
        "manifest.template.json"
      ),
      env: "dev",
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const result = await core.previewWithManifest(inputs);
    assert.isTrue(result.isErr());
    assert.deepEqual((result as any).error, { foo: "bar" });
  });

  it("getLaunchUrl error", async () => {
    const appName = await mockV3Project();
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(new TeamsAppManifest()));
    sinon.stub(LaunchHelper.prototype, "getLaunchUrl").resolves(err({ foo: "bar" } as any));
    const inputs: Inputs = {
      [QuestionNames.M365Host]: HubOptions.teams().id,
      [QuestionNames.TeamsAppManifestFilePath]: path.join(
        os.tmpdir(),
        appName,
        "appPackage",
        "manifest.template.json"
      ),
      env: "dev",
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const result = await core.previewWithManifest(inputs);
    assert.isTrue(result.isErr());
    assert.deepEqual((result as any).error, { foo: "bar" });
  });

  it("happy path", async () => {
    const appName = await mockV3Project();
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(new TeamsAppManifest()));
    sinon.stub(LaunchHelper.prototype, "getLaunchUrl").resolves(ok("test-url"));
    const inputs: Inputs = {
      [QuestionNames.M365Host]: HubOptions.teams().id,
      [QuestionNames.TeamsAppManifestFilePath]: path.join(
        os.tmpdir(),
        appName,
        "appPackage",
        "manifest.template.json"
      ),
      env: "dev",
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const result = await core.previewWithManifest(inputs);
    assert.isTrue(result.isOk());
    assert.deepEqual((result as any).value, "test-url");
  });
});

describe("getProjectId", async () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("happy path", async () => {
    const core = new FxCore(tools);
    sandbox.stub(core, "getProjectMetadata").resolves(
      ok({
        projectId: "12345",
        version: "1.1.1",
      })
    );
    const res = await core.getProjectId(".");
    assert.isTrue(res.isOk() && res.value === "12345");
  });
  it("return empty value", async () => {
    const core = new FxCore(tools);
    sandbox.stub(core, "getProjectMetadata").resolves(ok({}));
    const res = await core.getProjectId(".");
    assert.isTrue(res.isOk() && res.value === "");
  });
});
describe("getProjectMetadata", async () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("happy path", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").resolves("version: 1.1.1\nprojectId: 12345" as any);
    const core = new FxCore(tools);
    const res = await core.getProjectMetadata(".");
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(res.value, {
        projectId: "12345",
        version: "1.1.1",
      });
    }
  });
  it("yml not exist", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    sandbox.stub(fs, "pathExists").resolves(false);
    const core = new FxCore(tools);
    const res = await core.getProjectMetadata(".");
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(res.value, {});
    }
  });
  it("throw error", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    sandbox.stub(fs, "pathExists").rejects(new Error("mocked error"));
    const core = new FxCore(tools);
    const res = await core.getProjectMetadata(".");
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(res.value, {});
    }
  });
});
describe("getTeamsAppName", async () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("happy path", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    const mockProjectModel: any = {
      projectId: "12345",
      provision: {
        name: "provision",
        driverDefs: [
          {
            uses: "teamsApp/create",
            with: {
              name: "testappname-${{TEAMSFX_ENV}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
        ],
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    const core = new FxCore(tools);
    const res = await core.getTeamsAppName(".");
    assert.isTrue(res.isOk() && res.value === "testappname-");
  });
  it("happy path", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    const mockProjectModel: any = {
      projectId: "12345",
      provision: {
        name: "provision",
        driverDefs: [
          {
            uses: "teamsApp/create",
            with: {
              name: "testappname${{APP_NAME_SUFFIX}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
        ],
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    const core = new FxCore(tools);
    const res = await core.getTeamsAppName(".");
    assert.isTrue(res.isOk() && res.value === "testappname");
  });
  it("return empty value", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    const mockProjectModel: any = {};
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    const core = new FxCore(tools);
    const res = await core.getTeamsAppName(".");
    assert.isTrue(res.isOk() && res.value === "");
  });
  it("parse yml error", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    sandbox.stub(metadataUtil, "parse").resolves(err(new UserError({})));
    const core = new FxCore(tools);
    const res = await core.getTeamsAppName(".");
    assert.isTrue(res.isErr());
  });
});

describe("getProjectInfo", async () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("happy path", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    const mockProjectModel: any = {
      projectId: "mock-project-id",
      provision: {
        name: "provision",
        driverDefs: [
          {
            uses: "teamsApp/create",
            with: {
              name: "testappname-${{TEAMSFX_ENV}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
        ],
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "readEnv").resolves(
      ok({
        TEAMS_APP_ID: "mock-team-app-id",
        TEAMS_APP_TENANT_ID: "mock-tenant-id",
      })
    );
    const core = new FxCore(tools);
    const res = await core.getProjectInfo(".", "dev");
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(res.value, {
        projectId: "mock-project-id",
        teamsAppId: "mock-team-app-id",
        m365TenantId: "mock-tenant-id",
        teamsAppName: "testappname",
      });
    }
  });
  it("parse yml error", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    sandbox.stub(metadataUtil, "parse").resolves(err(new UserError({})));
    const core = new FxCore(tools);
    const res = await core.getProjectInfo(".", "dev");
    assert.isTrue(res.isErr());
  });
  it("read env error", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    sandbox.stub(metadataUtil, "parse").resolves(ok({} as any));
    sandbox.stub(envUtil, "readEnv").resolves(err(new UserError({})));
    const core = new FxCore(tools);
    const res = await core.getProjectInfo(".", "dev");
    assert.isTrue(res.isErr());
  });
});

describe("checkProjectType", async () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("happy 1", async () => {
    sandbox.stub(projectTypeChecker, "checkProjectType").resolves({
      isTeamsFx: false,
      lauguages: [],
      hasTeamsManifest: false,
      dependsOnTeamsJs: false,
    });
    const core = new FxCore(tools);
    const res = await core.checkProjectType("");
    assert.isTrue(res.isOk());
  });

  it("happy 2", async () => {
    sandbox.stub(projectTypeChecker, "checkProjectType").resolves({
      isTeamsFx: true,
      teamsfxConfigType: TeamsfxConfigType.teamsappYml,
      teamsfxConfigVersion: "1.0.0",
      teamsfxVersionState: TeamsfxVersionState.Compatible,
      teamsfxProjectId: "xxxx-xxxx-xxxx",
      lauguages: [],
      hasTeamsManifest: true,
      manifestCapabilities: ["bot"],
      manifestAppId: "xxx",
      manifestVersion: "1.16",
      dependsOnTeamsJs: true,
    });
    const core = new FxCore(tools);
    const res = await core.checkProjectType("");
    assert.isTrue(res.isOk());
  });
});

describe("isEnvFile", async () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("file patten not match", async () => {
    const core = new FxCore(tools);
    const res = await core.isEnvFile(".", ".abc.dev");
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.isFalse(res.value);
    }
  });
  it("getEnvFolderPath return error", async () => {
    sandbox.stub(pathUtils, "getEnvFolderPath").resolves(err(new UserError({})));
    const core = new FxCore(tools);
    const res = await core.isEnvFile(".", ".env.dev");
    assert.isTrue(res.isErr());
  });
  it("getEnvFolderPath return undefined", async () => {
    sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok(undefined));
    const core = new FxCore(tools);
    const res = await core.isEnvFile(".", ".env.dev");
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.isFalse(res.value);
    }
  });
  it("folder not match", async () => {
    sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("/tmp"));
    const core = new FxCore(tools);
    const res = await core.isEnvFile("/tmp", "/tmp1/.env.dev");
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.isFalse(res.value);
    }
  });
  it("match", async () => {
    sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("/tmp"));
    const core = new FxCore(tools);
    const res = await core.isEnvFile("/tmp", "/tmp/.env.dev");
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.isTrue(res.value);
    }
  });
});
describe("getQuestions", async () => {
  const sandbox = sinon.createSandbox();
  let mockedEnvRestore: RestoreFn = () => {};
  afterEach(() => {
    sandbox.restore();
    mockedEnvRestore();
  });
  it("happy path", async () => {
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
      [FeatureFlagName.CopilotPlugin]: "false",
    });
    const core = new FxCore(tools);
    const res = await core.getQuestions(Stage.create, { platform: Platform.CLI_HELP });
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      const node = res.value;
      const names: string[] = [];
      collectNodeNames(node!, names);
      assert.deepEqual(names, [
        "capabilities",
        "bot-host-type-trigger",
        "spfx-solution",
        "spfx-install-latest-package",
        "spfx-framework-type",
        "spfx-webpart-name",
        "spfx-folder",
        "me-architecture",
        "openapi-spec-location",
        "api-operation",
        "api-me-auth",
        "custom-copilot-rag",
        "openapi-spec-location",
        "api-operation",
        "custom-copilot-agent",
        "programming-language",
        "llm-service",
        "azure-openai-key",
        "azure-openai-endpoint",
        "openai-key",
        "office-addin-framework-type",
        "folder",
        "app-name",
      ]);
    }
  });
  it("happy path with runtime", async () => {
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "true",
      [FeatureFlagName.CopilotPlugin]: "false",
    });
    const core = new FxCore(tools);
    const res = await core.getQuestions(Stage.create, { platform: Platform.CLI_HELP });
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      const node = res.value;
      const names: string[] = [];
      collectNodeNames(node!, names);
      assert.deepEqual(names, [
        "runtime",
        "capabilities",
        "bot-host-type-trigger",
        "spfx-solution",
        "spfx-install-latest-package",
        "spfx-framework-type",
        "spfx-webpart-name",
        "spfx-folder",
        "me-architecture",
        "openapi-spec-location",
        "api-operation",
        "api-me-auth",
        "custom-copilot-rag",
        "openapi-spec-location",
        "api-operation",
        "custom-copilot-agent",
        "programming-language",
        "llm-service",
        "azure-openai-key",
        "azure-openai-endpoint",
        "openai-key",
        "office-addin-framework-type",
        "folder",
        "app-name",
      ]);
    }
  });

  it("happy path: API Copilot plugin enabled", async () => {
    const restore = mockedEnv({
      [FeatureFlagName.CopilotPlugin]: "true",
      [FeatureFlagName.ApiCopilotPlugin]: "true",
    });
    const core = new FxCore(tools);
    const res = await core.getQuestions(Stage.create, { platform: Platform.CLI_HELP });
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      const node = res.value;
      const names: string[] = [];
      collectNodeNames(node!, names);
      assert.deepEqual(names, [
        "capabilities",
        "bot-host-type-trigger",
        "spfx-solution",
        "spfx-install-latest-package",
        "spfx-framework-type",
        "spfx-webpart-name",
        "spfx-folder",
        "me-architecture",
        "openapi-spec-location",
        "api-operation",
        "api-me-auth",
        "custom-copilot-rag",
        "openapi-spec-location",
        "api-operation",
        "custom-copilot-agent",
        "programming-language",
        "llm-service",
        "azure-openai-key",
        "azure-openai-endpoint",
        "openai-key",
        "office-addin-framework-type",
        "folder",
        "app-name",
      ]);
    }
    restore();
  });

  it("happy path: copilot feature enabled but not API Copilot plugin", async () => {
    const restore = mockedEnv({
      [FeatureFlagName.CopilotPlugin]: "true",
      [FeatureFlagName.ApiCopilotPlugin]: "false",
    });
    const core = new FxCore(tools);
    const res = await core.getQuestions(Stage.create, { platform: Platform.CLI_HELP });
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      const node = res.value;
      const names: string[] = [];
      collectNodeNames(node!, names);
      assert.deepEqual(names, [
        "capabilities",
        "bot-host-type-trigger",
        "spfx-solution",
        "spfx-install-latest-package",
        "spfx-framework-type",
        "spfx-webpart-name",
        "spfx-folder",
        "me-architecture",
        "openapi-spec-location",
        "api-operation",
        "api-me-auth",
        "custom-copilot-rag",
        "openapi-spec-location",
        "api-operation",
        "custom-copilot-agent",
        "programming-language",
        "llm-service",
        "azure-openai-key",
        "azure-openai-endpoint",
        "openai-key",
        "office-addin-framework-type",
        "folder",
        "app-name",
      ]);
    }
    restore();
  });

  function collectNodeNames(node: IQTreeNode, names: string[]) {
    if (node.data.type !== "group") {
      names.push(node.data.name);
    }
    if (node.children) {
      for (const child of node.children) {
        collectNodeNames(child, names);
      }
    }
  }
});
describe("copilotPlugin", async () => {
  let mockedEnvRestore: RestoreFn = () => {};

  afterEach(() => {
    sinon.restore();
    mockedEnvRestore();
  });

  it("add API - json", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];
    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server",
          api: "GET /user/{userId}",
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server",
          api: "GET /store/order",
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };
    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
  });

  it("add API - Copilot plugin", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ManifestPath]: "manifest.json",
      [QuestionNames.Capabilities]: CapabilityOptions.copilotPluginApiSpec().id,
      [QuestionNames.DestinationApiSpecFilePath]: "destination.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.plugins = [
      {
        pluginFile: "ai-plugin.json",
      },
    ];
    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server",
          api: "GET /user/{userId}",
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server",
          api: "GET /store/order",
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generateForCopilot").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(manifestUtils, "getPluginFilePath").resolves(ok("ai-plugin.json"));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(CopilotPluginHelper, "listPluginExistingOperations").resolves([]);
    const result = await core.copilotPluginAddAPI(inputs);
    if (result.isErr()) {
      console.log(result.error);
    }
    assert.isTrue(result.isOk());
  });

  it("add API missing required input - Copilot plugin", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ManifestPath]: "manifest.json",
      [QuestionNames.Capabilities]: CapabilityOptions.copilotPluginApiSpec().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.plugins = [
      {
        pluginFile: "ai-plugin.json",
      },
    ];
    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server",
          api: "GET /user/{userId}",
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server",
          api: "GET /store/order",
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generateForCopilot").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(CopilotPluginHelper, "listPluginExistingOperations").resolves([]);
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.isTrue(result.error instanceof MissingRequiredInputError);
    }
  });

  it("add API error when getting plugin path - Copilot plugin", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ManifestPath]: "manifest.json",
      [QuestionNames.Capabilities]: CapabilityOptions.copilotPluginApiSpec().id,
      [QuestionNames.DestinationApiSpecFilePath]: "destination.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.plugins = [
      {
        pluginFile: "ai-plugin.json",
      },
    ];
    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server",
          api: "GET /user/{userId}",
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server",
          api: "GET /store/order",
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generateForCopilot").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon
      .stub(manifestUtils, "getPluginFilePath")
      .resolves(err(new SystemError("testError", "testError", "", "")));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(CopilotPluginHelper, "listPluginExistingOperations").resolves([]);
    const result = await core.copilotPluginAddAPI(inputs);

    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "testError");
    }
  });

  it("add API - return multiple auth error", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
      [FeatureFlagName.ApiKey]: "true",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth2",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal((result.error as FxError).name, "MultipleAuthError");
    }
  });

  it("add API - return multiple server error", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
      [FeatureFlagName.ApiKey]: "true",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server2",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal((result.error as FxError).name, "MultipleServerError");
    }
  });

  it("add API - no provision section in teamsapp yaml file", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
      [FeatureFlagName.ApiKey]: "true",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    const teamsappObject = {
      version: "1.0.0",
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").resolves(yamlString as any);
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal((result.error as FxError).name, "InjectAPIKeyActionFailedError");
    }
  });

  it("add API - no teamsApp/create action in teamsapp yaml file", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
      [FeatureFlagName.ApiKey]: "true",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").resolves(yamlString as any);
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal((result.error as FxError).name, "InjectAPIKeyActionFailedError");
    }
  });

  it("add API - no teams app id in teamsapp yaml file", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
      [FeatureFlagName.ApiKey]: "true",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/create",
          with: {
            name: "dfefeef-${{TEAMSFX_ENV}}",
          },
          writeToEnvironmentFile: {
            otherEnv: "OtherEnv",
          },
        },
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").resolves(yamlString as any);
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal((result.error as FxError).name, "InjectAPIKeyActionFailedError");
    }
  });

  it("add API - should inject api key action to teamsapp yaml file", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
      [FeatureFlagName.ApiKey]: "true",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: path.join(os.tmpdir(), appName, "appPackage/manifest.json"),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/create",
          with: {
            name: "dfefeef-${{TEAMSFX_ENV}}",
          },
          writeToEnvironmentFile: {
            teamsAppId: "TEAMS_APP_ID",
          },
        },
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").callsFake(async (path: string) => {
      return !path.endsWith("yml");
    });
    sinon.stub(fs, "readFile").resolves(yamlString as any);

    let writeYamlObjectTriggeredTimes = 0;
    sinon.stub(fs, "writeFile").callsFake((_, yamlString) => {
      writeYamlObjectTriggeredTimes++;
      const yamlObject = jsyaml.load(yamlString);

      assert.deepEqual(yamlObject, {
        provision: [
          {
            uses: "teamsApp/create",
            with: {
              name: "dfefeef-${{TEAMSFX_ENV}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
          {
            uses: "apiKey/register",
            with: {
              name: "bearerAuth1",
              appId: "${{TEAMS_APP_ID}}",
              apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
            },
            writeToEnvironmentFile: {
              registrationId: "BEARERAUTH1_REGISTRATION_ID",
            },
          },
          {
            uses: "teamsApp/zipAppPackage",
            with: {
              manifestPath: "./appPackage/manifest.json",
              outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
              outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
            },
          },
        ],
      });
    });

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
    assert.isTrue(writeYamlObjectTriggeredTimes === 1);
  });

  it("add API - should not inject api key action to teamsapp yaml file when already exists", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
      [FeatureFlagName.ApiKey]: "true",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: path.join(os.tmpdir(), appName, "appPackage/manifest.json"),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/create",
          with: {
            name: "dfefeef-${{TEAMSFX_ENV}}",
          },
          writeToEnvironmentFile: {
            teamsAppId: "TEAMS_APP_ID",
          },
        },
        {
          uses: "apiKey/register",
          with: {
            name: "bearerAuth1",
            appId: "${{TEAMS_APP_ID}}",
            apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
          },
          writeToEnvironmentFile: {
            registrationId: "BEARERAUTH1_REGISTRATION_ID",
          },
        },
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").callsFake(async (path: string) => {
      return !path.endsWith("yml");
    });
    sinon.stub(fs, "readFile").resolves(yamlString as any);

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
  });

  it("add API - should inject api key action to teamsapp yaml file when name does not match", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
      [FeatureFlagName.ApiKey]: "true",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: path.join(os.tmpdir(), appName, "appPackage/manifest.json"),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/create",
          with: {
            name: "dfefeef-${{TEAMSFX_ENV}}",
          },
          writeToEnvironmentFile: {
            teamsAppId: "TEAMS_APP_ID",
          },
        },
        {
          uses: "apiKey/register",
          with: {
            name: "api_key2",
            appId: "${{TEAMS_APP_ID}}",
            apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
          },
          writeToEnvironmentFile: {
            registrationId: "API_KEY1_REGISTRATION_ID",
          },
        },
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").callsFake(async (path: string) => {
      return !path.endsWith("yml");
    });
    sinon.stub(fs, "readFile").resolves(yamlString as any);

    let writeYamlObjectTriggeredTimes = 0;
    sinon.stub(fs, "writeFile").callsFake((_, yamlString) => {
      writeYamlObjectTriggeredTimes++;
      const yamlObject = jsyaml.load(yamlString);

      assert.deepEqual(yamlObject, {
        provision: [
          {
            uses: "teamsApp/create",
            with: {
              name: "dfefeef-${{TEAMSFX_ENV}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
          {
            uses: "apiKey/register",
            with: {
              name: "bearerAuth1",
              appId: "${{TEAMS_APP_ID}}",
              apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
            },
            writeToEnvironmentFile: {
              registrationId: "BEARERAUTH1_REGISTRATION_ID",
            },
          },
          {
            uses: "teamsApp/zipAppPackage",
            with: {
              manifestPath: "./appPackage/manifest.json",
              outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
              outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
            },
          },
        ],
      });
    });

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
    assert.isTrue(writeYamlObjectTriggeredTimes === 1);
  });

  it("add API - should inject api key action to teamsapp yaml file when missing with in yaml", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
      [FeatureFlagName.ApiKey]: "true",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: path.join(os.tmpdir(), appName, "appPackage/manifest.json"),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/create",
          with: {
            name: "dfefeef-${{TEAMSFX_ENV}}",
          },
          writeToEnvironmentFile: {
            teamsAppId: "TEAMS_APP_ID",
          },
        },
        {
          uses: "apiKey/register",
          writeToEnvironmentFile: {
            registrationId: "BEARERAUTH1_REGISTRATION_ID",
          },
        },
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").callsFake(async (path: string) => {
      return !path.endsWith("yml");
    });
    sinon.stub(fs, "readFile").resolves(yamlString as any);

    let writeYamlObjectTriggeredTimes = 0;
    sinon.stub(fs, "writeFile").callsFake((_, yamlString) => {
      writeYamlObjectTriggeredTimes++;
      const yamlObject = jsyaml.load(yamlString);

      assert.deepEqual(yamlObject, {
        provision: [
          {
            uses: "teamsApp/create",
            with: {
              name: "dfefeef-${{TEAMSFX_ENV}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
          {
            uses: "apiKey/register",
            with: {
              name: "bearerAuth1",
              appId: "${{TEAMS_APP_ID}}",
              apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
            },
            writeToEnvironmentFile: {
              registrationId: "BEARERAUTH1_REGISTRATION_ID",
            },
          },
          {
            uses: "teamsApp/zipAppPackage",
            with: {
              manifestPath: "./appPackage/manifest.json",
              outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
              outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
            },
          },
        ],
      });
    });

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
    assert.isTrue(writeYamlObjectTriggeredTimes === 1);
  });

  it("add API - should inject api key action to teamsapp yaml file when missing name in yaml", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
      [FeatureFlagName.ApiKey]: "true",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: path.join(os.tmpdir(), appName, "appPackage/manifest.json"),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };
    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/create",
          with: {
            name: "dfefeef-${{TEAMSFX_ENV}}",
          },
          writeToEnvironmentFile: {
            teamsAppId: "TEAMS_APP_ID",
          },
        },
        {
          uses: "apiKey/register",
          with: {
            appId: "${{TEAMS_APP_ID}}",
            apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
          },
          writeToEnvironmentFile: {
            registrationId: "API_KEY1_REGISTRATION_ID",
          },
        },
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").callsFake(async (path: string) => {
      return !path.endsWith("yml");
    });
    sinon.stub(fs, "readFile").resolves(yamlString as any);

    let writeYamlObjectTriggeredTimes = 0;
    sinon.stub(fs, "writeFile").callsFake((_, yamlString) => {
      writeYamlObjectTriggeredTimes++;
      const yamlObject = jsyaml.load(yamlString);

      assert.deepEqual(yamlObject, {
        provision: [
          {
            uses: "teamsApp/create",
            with: {
              name: "dfefeef-${{TEAMSFX_ENV}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
          {
            uses: "apiKey/register",
            with: {
              name: "bearerAuth1",
              appId: "${{TEAMS_APP_ID}}",
              apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
            },
            writeToEnvironmentFile: {
              registrationId: "BEARERAUTH1_REGISTRATION_ID",
            },
          },
          {
            uses: "teamsApp/zipAppPackage",
            with: {
              manifestPath: "./appPackage/manifest.json",
              outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
              outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
            },
          },
        ],
      });
    });

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
    assert.isTrue(writeYamlObjectTriggeredTimes === 1);
  });

  it("add API - should inject api key action to teamsapp yaml file with local teamsapp file", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
      [FeatureFlagName.ApiKey]: "true",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: path.join(os.tmpdir(), appName, "appPackage/manifest.json"),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/create",
          with: {
            name: "dfefeef-${{TEAMSFX_ENV}}",
          },
          writeToEnvironmentFile: {
            teamsAppId: "TEAMS_APP_ID",
          },
        },
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").resolves(yamlString as any);

    let writeYamlObjectTriggeredTimes = 0;
    sinon.stub(fs, "writeFile").callsFake((_, yamlString) => {
      writeYamlObjectTriggeredTimes++;
      const yamlObject = jsyaml.load(yamlString);
      assert.deepEqual(yamlObject, {
        provision: [
          {
            uses: "teamsApp/create",
            with: {
              name: "dfefeef-${{TEAMSFX_ENV}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
          {
            uses: "apiKey/register",
            with: {
              name: "bearerAuth1",
              appId: "${{TEAMS_APP_ID}}",
              apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
            },
            writeToEnvironmentFile: {
              registrationId: "BEARERAUTH1_REGISTRATION_ID",
            },
          },
          {
            uses: "teamsApp/zipAppPackage",
            with: {
              manifestPath: "./appPackage/manifest.json",
              outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
              outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
            },
          },
        ],
      });
    });

    const result = await core.copilotPluginAddAPI(inputs);

    assert.isTrue(result.isOk());
    assert.isTrue(writeYamlObjectTriggeredTimes === 2);
  });

  it("add API - should filter unknown api key action", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
      [FeatureFlagName.ApiKey]: "true",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: path.join(os.tmpdir(), appName, "appPackage/manifest.json"),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/create",
          with: {
            name: "dfefeef-${{TEAMSFX_ENV}}",
          },
          writeToEnvironmentFile: {
            teamsAppId: "TEAMS_APP_ID",
          },
        },
        {
          uses: "apiKey/register",
          with: {
            name: "api_key_unknown",
            appId: "${{TEAMS_APP_ID}}",
            apiSpecPath: ".appPackage/apiSpecificationFiles/openapi.json",
          },
          writeToEnvironmentFile: {
            registrationId: "API_KEY_UNKNOWN_REGISTRATION_ID",
          },
        },
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").resolves(yamlString as any);

    let writeYamlObjectTriggeredTimes = 0;
    sinon.stub(fs, "writeFile").callsFake((_, yamlString) => {
      writeYamlObjectTriggeredTimes++;
      const yamlObject = jsyaml.load(yamlString);
      assert.deepEqual(yamlObject, {
        provision: [
          {
            uses: "teamsApp/create",
            with: {
              name: "dfefeef-${{TEAMSFX_ENV}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
          {
            uses: "apiKey/register",
            with: {
              name: "bearerAuth1",
              appId: "${{TEAMS_APP_ID}}",
              apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
            },
            writeToEnvironmentFile: {
              registrationId: "BEARERAUTH1_REGISTRATION_ID",
            },
          },
          {
            uses: "teamsApp/zipAppPackage",
            with: {
              manifestPath: "./appPackage/manifest.json",
              outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
              outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
            },
          },
        ],
      });
    });

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
    assert.isTrue(writeYamlObjectTriggeredTimes === 2);
  });

  it("add API - warnings", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ManifestPath]: path.join(os.tmpdir(), appName, "appPackage/manifest.json"),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [
          {
            id: "getUserById",
            title: "Get User By Id",
          },
          {
            id: "notexist",
            title: "Get User By Id",
          },
        ],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server",
          api: "GET /user/{userId}",
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server",
          api: "GET /store/order",
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [{ type: WarningType.OperationOnlyContainsOptionalParam, content: "fakeMessage" }],
      allSuccess: false,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
  });

  it("add API - readManifestFailed", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["testOperation"],
      projectPath: path.join(os.tmpdir(), appName),
    };
    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").throws(new Error("fakeError"));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
  });

  it("add API - assembleError", async () => {
    const appName = await mockV3Project();
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["testOperation"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").throws(new Error("fakeError"));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
  });

  it("add API - SpecParserError", async () => {
    const appName = await mockV3Project();
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["testOperation"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const core = new FxCore(tools);
    sinon
      .stub(SpecParser.prototype, "generate")
      .throws(new SpecParserError("fakeMessage", ErrorType.SpecNotValid));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
  });

  describe("listPluginApiSpecs", async () => {
    it("success", async () => {
      const inputs = {
        [QuestionNames.ManifestPath]: "manifest.json",
        platform: Platform.VS,
      };
      const manifest = new TeamsAppManifest();
      manifest.plugins = [
        {
          pluginFile: "ai-plugin.json",
        },
      ];
      sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
      sinon
        .stub(pluginManifestUtils, "getApiSpecFilePathFromTeamsManifest")
        .resolves(ok(["apispec.json"]));

      const core = new FxCore(tools);
      const res = await core.listPluginApiSpecs(inputs);

      assert.isTrue(res.isOk());
    });

    it("read manifest error", async () => {
      const inputs = {
        [QuestionNames.ManifestPath]: "manifest.json",
        platform: Platform.VS,
      };
      sinon
        .stub(manifestUtils, "_readAppManifest")
        .resolves(err(new SystemError("read manifest error", "read manifest error", "", "")));

      const core = new FxCore(tools);
      const res = await core.listPluginApiSpecs(inputs);

      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, "read manifest error");
      }
    });

    it("get api spec error", async () => {
      const inputs = {
        [QuestionNames.ManifestPath]: "manifest.json",
        platform: Platform.VS,
      };
      const manifest = new TeamsAppManifest();
      manifest.plugins = [
        {
          pluginFile: "ai-plugin.json",
        },
      ];
      sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
      sinon
        .stub(pluginManifestUtils, "getApiSpecFilePathFromTeamsManifest")
        .resolves(err(new SystemError("get plugin error", "get plugin error", "", "")));

      const core = new FxCore(tools);
      const res = await core.listPluginApiSpecs(inputs);

      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, "get plugin error");
      }
    });
  });

  it("load OpenAI manifest - should run successful", async () => {
    const core = new FxCore(tools);
    const inputs = { domain: "mydomain.com" };
    sinon
      .stub(OpenAIPluginManifestHelper, "loadOpenAIPluginManifest")
      .returns(Promise.resolve<OpenAIPluginManifest>(undefined as any));
    const result = await core.copilotPluginLoadOpenAIManifest(inputs as any);
    assert.isTrue(result.isOk());
  });

  it("load OpenAI manifest - should return an error when an exception is thrown", async () => {
    const core = new FxCore(tools);
    const inputs = { domain: "mydomain.com" };
    sinon
      .stub(OpenAIPluginManifestHelper, "loadOpenAIPluginManifest")
      .throws(new Error("Test error"));
    const result = await core.copilotPluginLoadOpenAIManifest(inputs as any);
    assert.isTrue(result.isErr());
  });

  it("load operations - should return a list of operations when given valid inputs", async () => {
    const core = new FxCore(tools);
    const inputs = {
      manifest: {},
      apiSpecUrl: "https://example.com/api-spec",
      shouldLogWarning: true,
    };
    const expectedResult = [
      {
        id: "operation1",
        label: "operation1",
        groupName: "1",
        data: { serverUrl: "https://server1" },
      },
      {
        id: "operation2",
        label: "operation2",
        groupName: "2",
        data: { serverUrl: "https://server2" },
      },
    ];
    sinon.stub(CopilotPluginHelper, "listOperations").returns(Promise.resolve(ok(expectedResult)));
    const result = await core.copilotPluginListOperations(inputs as any);
    assert.isTrue(result.isOk());
    if (result.isOk()) {
      assert.deepEqual(result.value, expectedResult);
    }
  });

  it("load operations - should return an error when an exception is thrown", async () => {
    const core = new FxCore(tools);
    const inputs = {
      manifest: {},
      apiSpecUrl: "https://example.com/api-spec",
      shouldLogWarning: true,
    };
    sinon.stub(CopilotPluginHelper, "listOperations").returns(Promise.resolve(err([])));
    const result = await core.copilotPluginListOperations(inputs as any);
    assert.isTrue(result.isErr());
  });

  it("load operations - no manifest in inputs", async () => {
    const core = new FxCore(tools);
    const inputs = {
      apiSpecUrl: "https://example.com/api-spec",
      shouldLogWarning: true,
      includeExistingAPIs: false,
    };

    sinon
      .stub(SpecParser.prototype, "validate")
      .resolves({ status: ValidationStatus.Valid, warnings: [], errors: [] });
    sinon
      .stub(SpecParser.prototype, "list")
      .resolves({ APIs: [], allAPICount: 0, validAPICount: 0 });

    try {
      await core.copilotPluginListOperations(inputs as any);
    } catch (e: any) {
      assert.equal(e.name, MissingRequiredInputError.name);
    }
  });

  it("load operations - invalid manifest", async () => {
    const core = new FxCore(tools);
    const inputs = {
      apiSpecUrl: "https://example.com/api-spec",
      shouldLogWarning: true,
      includeExistingAPIs: false,
      "manifest-path": "fakePath",
    };

    sinon
      .stub(manifestUtils, "_readAppManifest")
      .returns(Promise.resolve(err(new FileNotFoundError("file", "fakePath"))));
    sinon
      .stub(SpecParser.prototype, "validate")
      .resolves({ status: ValidationStatus.Valid, warnings: [], errors: [] });
    sinon
      .stub(SpecParser.prototype, "list")
      .resolves({ APIs: [], allAPICount: 0, validAPICount: 0 });

    try {
      await core.copilotPluginListOperations(inputs as any);
    } catch (e: any) {
      assert.equal(e.name, FileNotFoundError.name);
    }
  });

  it("teamsapp management APIs", async () => {
    const core = new FxCore(tools);
    const inputs = {
      platform: Platform.CLI,
    };
    sinon.stub(teamsappMgr, "updateTeamsApp").resolves(ok(undefined));
    sinon
      .stub(teamsappMgr, "packageTeamsApp")
      .resolves(ok({ manifestPath: "", outputJsonPath: "", outputZipPath: "" }));
    sinon.stub(teamsappMgr, "validateTeamsApp").resolves(ok(undefined));
    sinon.stub(teamsappMgr, "publishTeamsApp").resolves(ok(undefined));
    const res1 = await core.updateTeamsAppCLIV3(inputs as any);
    const res2 = await core.packageTeamsAppCLIV3(inputs as any);
    const res3 = await core.validateTeamsAppCLIV3(inputs as any);
    const res4 = await core.publishTeamsAppCLIV3(inputs as any);
    assert.isTrue(res1.isOk());
    assert.isTrue(res2.isOk());
    assert.isTrue(res3.isOk());
    assert.isTrue(res4.isOk());
  });
});
