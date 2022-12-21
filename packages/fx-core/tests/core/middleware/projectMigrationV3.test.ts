// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  err,
  FxError,
  Inputs,
  ok,
  Platform,
  Result,
  SettingsFileName,
  SettingsFolderName,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import mockedEnv from "mocked-env";
import * as os from "os";
import * as path from "path";
import * as sinon from "sinon";
import { MockTools, MockUserInteraction, randomAppName } from "../utils";
import { CoreHookContext } from "../../../src/core/types";
import { setTools } from "../../../src/core/globalVars";
import { MigrationContext } from "../../../src/core/middleware/utils/migrationContext";
import {
  generateAppYml,
  generateSettingsJson,
  manifestsMigration,
  statesMigration,
  updateLaunchJson,
  migrate,
  wrapRunMigration,
  checkVersionForMigration,
  configsMigration,
  generateApimPluginEnvContent,
  userdataMigration,
  debugMigration,
  azureParameterMigration,
  generateLocalConfig,
  checkapimPluginExists,
  ProjectMigratorMWV3,
} from "../../../src/core/middleware/projectMigratorV3";
import * as MigratorV3 from "../../../src/core/middleware/projectMigratorV3";
import { UpgradeCanceledError } from "../../../src/core/error";
import {
  Metadata,
  MetadataV2,
  MetadataV3,
  VersionState,
} from "../../../src/common/versionMetadata";
import {
  getDownloadLinkByVersionAndPlatform,
  getTrackingIdFromPath,
  getVersionState,
  migrationNotificationMessage,
  outputCancelMessage,
} from "../../../src/core/middleware/utils/v3MigrationUtils";
import { getProjectSettingPathV3 } from "../../../src/core/middleware/projectSettingsLoader";
import * as debugV3MigrationUtils from "../../../src/core/middleware/utils/debug/debugV3MigrationUtils";
import { VersionForMigration } from "../../../src/core/middleware/types";
import { isMigrationV3Enabled } from "../../../src/common/tools";

let mockedEnvRestore: () => void;

describe("ProjectMigratorMW", () => {
  const sandbox = sinon.createSandbox();
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
    await fs.ensureDir(path.join(projectPath, ".fx"));
    mockedEnvRestore = mockedEnv({
      TEAMSFX_V3_MIGRATION: "true",
    });
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
    mockedEnvRestore();
  });

  it("happy path", async () => {
    sandbox.stub(MockUserInteraction.prototype, "showMessage").resolves(ok("Upgrade"));
    const tools = new MockTools();
    setTools(tools);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    class MyClass {
      tools = tools;
      async other(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
        return ok("");
      }
    }
    hooks(MyClass, {
      other: [ProjectMigratorMWV3],
    });

    const inputs: Inputs = { platform: Platform.VSCode, ignoreEnvInfo: true };
    inputs.projectPath = projectPath;
    const my = new MyClass();
    try {
      const res = await my.other(inputs);
      assert.isTrue(res.isOk());
    } finally {
      await fs.rmdir(inputs.projectPath!, { recursive: true });
    }
  });

  it("user cancel", async () => {
    sandbox
      .stub(MockUserInteraction.prototype, "showMessage")
      .resolves(err(new Error("user cancel") as FxError));
    const tools = new MockTools();
    setTools(tools);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    class MyClass {
      tools = tools;
      async other(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
        return ok("");
      }
    }
    hooks(MyClass, {
      other: [ProjectMigratorMWV3],
    });

    const inputs: Inputs = { platform: Platform.VSCode, ignoreEnvInfo: true };
    inputs.projectPath = projectPath;
    const my = new MyClass();
    try {
      const res = await my.other(inputs);
      assert.isTrue(res.isErr());
    } finally {
      await fs.rmdir(inputs.projectPath!, { recursive: true });
    }
  });

  it("wrap run error ", async () => {
    const tools = new MockTools();
    setTools(tools);
    sandbox.stub(MigratorV3, "migrate").throws(new Error("mocker error"));
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    const inputs: Inputs = { platform: Platform.VSCode, ignoreEnvInfo: true };
    inputs.projectPath = projectPath;
    const ctx = {
      arguments: [inputs],
    };
    const context = await MigrationContext.create(ctx);
    const res = wrapRunMigration(context, migrate);
  });
});

describe("MigrationContext", () => {
  const sandbox = sinon.createSandbox();
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
    await fs.ensureDir(path.join(projectPath, ".fx"));
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
    mockedEnvRestore();
  });

  it("happy path", async () => {
    const tools = new MockTools();
    setTools(tools);

    const inputs: Inputs = { platform: Platform.VSCode, ignoreEnvInfo: true };
    inputs.projectPath = projectPath;
    const ctx = {
      arguments: [inputs],
    };
    const context = await MigrationContext.create(ctx);
    let res = await context.backup(".fx");
    assert.isTrue(res);
    res = await context.backup("no-exist");
    assert.isFalse(res);
    await context.fsWriteFile("a", "test-data");
    await context.fsCopy("a", "a-copy");
    assert.isTrue(await fs.pathExists(path.join(context.projectPath, "a-copy")));
    await context.fsEnsureDir("b/c");
    assert.isTrue(await fs.pathExists(path.join(context.projectPath, "b/c")));
    await context.fsCreateFile("d");
    assert.isTrue(await fs.pathExists(path.join(context.projectPath, "d")));
    const modifiedPaths = context.getModifiedPaths();
    assert.isTrue(modifiedPaths.includes("a"));
    assert.isTrue(modifiedPaths.includes("a-copy"));
    assert.isTrue(modifiedPaths.includes("b"));
    assert.isTrue(modifiedPaths.includes("b/c"));
    assert.isTrue(modifiedPaths.includes("d"));
    await context.fsRemove("d");
    await context.cleanModifiedPaths();
    assert.isEmpty(context.getModifiedPaths());

    context.addReport("test report");
    context.addTelemetryProperties({ testProperrty: "test property" });
    await context.restoreBackup();
    await context.cleanTeamsfx();
  });
});

describe("generateSettingsJson", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("happy path", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    await copyTestProject(Constants.happyPathTestProject, projectPath);
    const oldProjectSettings = await readOldProjectSettings(projectPath);

    await generateSettingsJson(migrationContext);

    assert.isTrue(
      await fs.pathExists(path.join(projectPath, SettingsFolderName, SettingsFileName))
    );
    const newSettings = await readSettingJson(projectPath);
    assert.equal(newSettings.trackingId, oldProjectSettings.projectId);
    assert.equal(newSettings.version, "3.0.0");
  });

  it("no project id", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    await copyTestProject(Constants.happyPathTestProject, projectPath);
    const projectSetting = await readOldProjectSettings(projectPath);
    delete projectSetting.projectId;
    await fs.writeJson(
      path.join(projectPath, Constants.oldProjectSettingsFilePath),
      projectSetting
    );

    await generateSettingsJson(migrationContext);

    const newSettings = await readSettingJson(projectPath);
    assert.isTrue(newSettings.hasOwnProperty("trackingId")); // will auto generate a new trackingId if old project does not have project id
  });
});

describe("generateAppYml-js/ts", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  let migrationContext: MigrationContext;

  beforeEach(async () => {
    migrationContext = await mockMigrationContext(projectPath);
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("should success for js SSO tab", async () => {
    await copyTestProject("jsSsoTab", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, "teamsfx/app.yml", "js.app.yml");
  });

  it("should success for ts SSO tab", async () => {
    await copyTestProject("jsSsoTab", projectPath);
    const projectSetting = await readOldProjectSettings(projectPath);
    projectSetting.programmingLanguage = "typescript";
    await fs.writeJson(
      path.join(projectPath, Constants.oldProjectSettingsFilePath),
      projectSetting
    );

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, "teamsfx/app.yml", "ts.app.yml");
  });

  it("should success for js non SSO tab", async () => {
    await copyTestProject("jsNonSsoTab", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, "teamsfx/app.yml", "js.app.yml");
  });

  it("should success for ts non SSO tab", async () => {
    await copyTestProject("jsNonSsoTab", projectPath);
    const projectSetting = await readOldProjectSettings(projectPath);
    projectSetting.programmingLanguage = "typescript";
    await fs.writeJson(
      path.join(projectPath, Constants.oldProjectSettingsFilePath),
      projectSetting
    );

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, "teamsfx/app.yml", "ts.app.yml");
  });

  it("should success for js tab with api", async () => {
    await copyTestProject("jsTabWithApi", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, "teamsfx/app.yml", "js.app.yml");
  });

  it("should success for ts tab with api", async () => {
    await copyTestProject("jsTabWithApi", projectPath);
    const projectSetting = await readOldProjectSettings(projectPath);
    projectSetting.programmingLanguage = "typescript";
    await fs.writeJson(
      path.join(projectPath, Constants.oldProjectSettingsFilePath),
      projectSetting
    );

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, "teamsfx/app.yml", "ts.app.yml");
  });

  it("should success for js function bot", async () => {
    await copyTestProject("jsFunctionBot", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, "teamsfx/app.yml", "js.app.yml");
  });

  it("should success for ts function bot", async () => {
    await copyTestProject("jsFunctionBot", projectPath);
    const projectSetting = await readOldProjectSettings(projectPath);
    projectSetting.programmingLanguage = "typescript";
    await fs.writeJson(
      path.join(projectPath, Constants.oldProjectSettingsFilePath),
      projectSetting
    );

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, "teamsfx/app.yml", "ts.app.yml");
  });

  it("should success for js webapp bot", async () => {
    await copyTestProject("jsWebappBot", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, "teamsfx/app.yml", "js.app.yml");
  });

  it("should success for ts webapp bot", async () => {
    await copyTestProject("jsWebappBot", projectPath);
    const projectSetting = await readOldProjectSettings(projectPath);
    projectSetting.programmingLanguage = "typescript";
    await fs.writeJson(
      path.join(projectPath, Constants.oldProjectSettingsFilePath),
      projectSetting
    );

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, "teamsfx/app.yml", "ts.app.yml");
  });
});

describe("generateAppYml-csharp", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  let migrationContext: MigrationContext;

  beforeEach(async () => {
    migrationContext = await mockMigrationContext(projectPath);
    migrationContext.arguments.push({
      platform: "vs",
    });
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("should success for sso tab project", async () => {
    await copyTestProject("csharpSsoTab", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, "teamsfx/app.yml", "app.yml");
  });

  it("should success for non-sso tab project", async () => {
    await copyTestProject("csharpNonSsoTab", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, "teamsfx/app.yml", "app.yml");
  });

  it("should success for web app bot project", async () => {
    await copyTestProject("csharpWebappBot", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, "teamsfx/app.yml", "app.yml");
  });

  it("should success for function bot project", async () => {
    await copyTestProject("csharpFunctionBot", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, "teamsfx/app.yml", "app.yml");
  });
});

describe("generateAppYml-csharp", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  let migrationContext: MigrationContext;

  beforeEach(async () => {
    migrationContext = await mockMigrationContext(projectPath);
    migrationContext.arguments.push({
      platform: "vs",
    });
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("should success for local sso tab project", async () => {
    await copyTestProject("csharpSsoTab", projectPath);

    await generateLocalConfig(migrationContext);
  });
});

describe("generateAppYml-spfx", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  let migrationContext: MigrationContext;

  beforeEach(async () => {
    migrationContext = await mockMigrationContext(projectPath);
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("should success for spfx project", async () => {
    await copyTestProject("spfxTab", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, "teamsfx/app.yml", "app.yml");
  });
});

describe("manifestsMigration", () => {
  const sandbox = sinon.createSandbox();
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
  });

  it("happy path: aad manifest exists", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    await copyTestProject(Constants.manifestsMigrationHappyPath, projectPath);

    // Action
    await manifestsMigration(migrationContext);

    // Assert
    const oldAppPackageFolderPath = path.join(projectPath, "templates", "appPackage");
    assert.isFalse(await fs.pathExists(oldAppPackageFolderPath));

    const appPackageFolderPath = path.join(projectPath, "appPackage");
    assert.isTrue(await fs.pathExists(appPackageFolderPath));

    const resourcesPath = path.join(appPackageFolderPath, "resources", "test.png");
    assert.isTrue(await fs.pathExists(resourcesPath));

    const manifestPath = path.join(appPackageFolderPath, "manifest.template.json");
    assert.isTrue(await fs.pathExists(manifestPath));
    const manifest = (await fs.readFile(manifestPath, "utf-8"))
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    const manifestExpeceted = (
      await fs.readFile(path.join(projectPath, "expected", "manifest.template.json"), "utf-8")
    )
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    assert.equal(manifest, manifestExpeceted);

    const aadManifestPath = path.join(projectPath, "aad.manifest.template.json");
    assert.isTrue(await fs.pathExists(aadManifestPath));
    const aadManifest = (await fs.readFile(aadManifestPath, "utf-8"))
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    const aadManifestExpected = (
      await fs.readFile(path.join(projectPath, "expected", "aad.manifest.template.json"), "utf-8")
    )
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    assert.equal(aadManifest, aadManifestExpected);
  });

  it("happy path: spfx", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    await copyTestProject(Constants.manifestsMigrationHappyPathSpfx, projectPath);

    // Action
    await manifestsMigration(migrationContext);

    // Assert
    const oldAppPackageFolderPath = path.join(projectPath, "templates", "appPackage");
    assert.isFalse(await fs.pathExists(oldAppPackageFolderPath));

    const appPackageFolderPath = path.join(projectPath, "appPackage");
    assert.isTrue(await fs.pathExists(appPackageFolderPath));

    const resourcesPath = path.join(appPackageFolderPath, "resources", "test.png");
    assert.isTrue(await fs.pathExists(resourcesPath));

    const remoteManifestPath = path.join(appPackageFolderPath, "manifest.template.json");
    assert.isTrue(await fs.pathExists(remoteManifestPath));
    const remoteManifest = (await fs.readFile(remoteManifestPath, "utf-8"))
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    const remoteManifestExpeceted = (
      await fs.readFile(path.join(projectPath, "expected", "manifest.template.json"), "utf-8")
    )
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    assert.equal(remoteManifest, remoteManifestExpeceted);

    const localManifestPath = path.join(appPackageFolderPath, "manifest.template.local.json");
    assert.isTrue(await fs.pathExists(localManifestPath));
    const localManifest = (await fs.readFile(localManifestPath, "utf-8"))
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    const localManifestExpeceted = (
      await fs.readFile(path.join(projectPath, "expected", "manifest.template.local.json"), "utf-8")
    )
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    assert.equal(localManifest, localManifestExpeceted);
  });

  it("happy path: aad manifest does not exist", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    await copyTestProject(Constants.manifestsMigrationHappyPath, projectPath);
    await fs.remove(path.join(projectPath, "templates/appPackage/aad.template.json"));

    // Action
    await manifestsMigration(migrationContext);

    // Assert
    const appPackageFolderPath = path.join(projectPath, "appPackage");
    assert.isTrue(await fs.pathExists(appPackageFolderPath));

    const resourcesPath = path.join(appPackageFolderPath, "resources", "test.png");
    assert.isTrue(await fs.pathExists(resourcesPath));

    const manifestPath = path.join(appPackageFolderPath, "manifest.template.json");
    assert.isTrue(await fs.pathExists(manifestPath));
    const manifest = (await fs.readFile(manifestPath, "utf-8"))
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    const manifestExpeceted = (
      await fs.readFile(path.join(projectPath, "expected", "manifest.template.json"), "utf-8")
    )
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    assert.equal(manifest, manifestExpeceted);

    const aadManifestPath = path.join(projectPath, "aad.manifest.template.json");
    assert.isFalse(await fs.pathExists(aadManifestPath));
  });

  it("migrate manifests failed: appPackage does not exist", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(false);

    try {
      await manifestsMigration(migrationContext);
    } catch (error) {
      assert.equal(error.name, "MigrationReadFileError");
      assert.equal(error.innerError.message, "templates/appPackage does not exist");
    }
  });

  it("migrate manifests success: provision.bicep does not exist", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    await copyTestProject(Constants.manifestsMigrationHappyPath, projectPath);
    await fs.remove(path.join(projectPath, "templates", "azure", "provision.bicep"));

    // Action
    await manifestsMigration(migrationContext);

    // Assert
    const appPackageFolderPath = path.join(projectPath, "appPackage");
    assert.isTrue(await fs.pathExists(appPackageFolderPath));

    const resourcesPath = path.join(appPackageFolderPath, "resources", "test.png");
    assert.isTrue(await fs.pathExists(resourcesPath));
  });

  it("migrate manifests failed: teams app manifest does not exist", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    await copyTestProject(Constants.manifestsMigrationHappyPath, projectPath);
    await fs.remove(path.join(projectPath, "templates/appPackage/manifest.template.json"));

    try {
      await manifestsMigration(migrationContext);
    } catch (error) {
      assert.equal(error.name, "MigrationReadFileError");
      assert.equal(
        error.innerError.message,
        "templates/appPackage/manifest.template.json does not exist"
      );
    }
  });
});

describe("azureParameterMigration", () => {
  const sandbox = sinon.createSandbox();
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
  });

  it("Happy Path", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    await copyTestProject(Constants.manifestsMigrationHappyPath, projectPath);

    // Action
    await azureParameterMigration(migrationContext);

    // Assert
    const azureParameterDevFilePath = path.join(
      projectPath,
      "templates",
      "azure",
      "azure.parameters.dev.json"
    );
    const azureParameterTestFilePath = path.join(
      projectPath,
      "templates",
      "azure",
      "azure.parameters.test.json"
    );
    assert.isTrue(await fs.pathExists(azureParameterDevFilePath));
    assert.isTrue(await fs.pathExists(azureParameterTestFilePath));
    const azureParameterExpected = await fs.readFile(
      path.join(projectPath, "expected", "azure.parameters.json"),
      "utf-8"
    );
    const azureParameterDev = await fs.readFile(azureParameterDevFilePath, "utf-8");
    const azureParameterTest = await fs.readFile(azureParameterTestFilePath, "utf-8");
    assert.equal(azureParameterDev, azureParameterExpected);
    assert.equal(azureParameterTest, azureParameterExpected);
  });

  it("migrate azure.parameter failed: .fx/config does not exist", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Action
    await azureParameterMigration(migrationContext);

    // Assert
    const azureParameterDevFilePath = path.join(
      projectPath,
      "templates",
      "azure",
      "azure.parameters.dev.json"
    );
    assert.isFalse(await fs.pathExists(azureParameterDevFilePath));
  });

  it("migrate azure.parameter failed: provision.bicep does not exist", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    await fs.ensureDir(path.join(projectPath, ".fx", "config"));

    try {
      await azureParameterMigration(migrationContext);
    } catch (error) {
      assert.equal(error.name, "MigrationReadFileError");
      assert.equal(error.innerError.message, "templates/azure/provision.bicep does not exist");
    }
  });
});

describe("updateLaunchJson", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("should success in happy path", async () => {
    const migrationContext = await mockMigrationContext(projectPath);
    await copyTestProject(Constants.happyPathTestProject, projectPath);

    await updateLaunchJson(migrationContext);

    assert.isTrue(
      await fs.pathExists(path.join(projectPath, "teamsfx/backup/.vscode/launch.json"))
    );
    const updatedLaunchJson = await fs.readJson(path.join(projectPath, Constants.launchJsonPath));
    assert.equal(
      updatedLaunchJson.configurations[0].url,
      "https://teams.microsoft.com/l/app/${dev:teamsAppId}?installAppPackage=true&webjoin=true&${account-hint}"
    );
    assert.equal(
      updatedLaunchJson.configurations[1].url,
      "https://teams.microsoft.com/l/app/${dev:teamsAppId}?installAppPackage=true&webjoin=true&${account-hint}"
    );
    assert.equal(
      updatedLaunchJson.configurations[2].url,
      "https://teams.microsoft.com/l/app/${local:teamsAppId}?installAppPackage=true&webjoin=true&${account-hint}"
    );
    assert.equal(
      updatedLaunchJson.configurations[3].url,
      "https://teams.microsoft.com/l/app/${local:teamsAppId}?installAppPackage=true&webjoin=true&${account-hint}"
    );
    assert.equal(
      updatedLaunchJson.configurations[4].url,
      "https://outlook.office.com/host/${local:teamsAppInternalId}?${account-hint}" // for M365 app
    );
    assert.equal(
      updatedLaunchJson.configurations[5].url,
      "https://outlook.office.com/host/${local:teamsAppInternalId}?${account-hint}" // for M365 app
    );
  });
});

describe("stateMigration", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("happy path", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    await copyTestProject(Constants.happyPathTestProject, projectPath);
    await statesMigration(migrationContext);

    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx")));

    const trueEnvContent_dev = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "state.dev"
    );
    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx", ".env.dev")));
    const testEnvContent_dev = await readEnvFile(path.join(projectPath, "teamsfx"), "dev");
    assert.equal(testEnvContent_dev, trueEnvContent_dev);

    const trueEnvContent_local = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "state.local"
    );
    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx", ".env.local")));
    const testEnvContent_local = await readEnvFile(path.join(projectPath, "teamsfx"), "local");
    assert.equal(testEnvContent_local, trueEnvContent_local);
  });
});

describe("configMigration", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("happy path", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    await copyTestProject(Constants.happyPathTestProject, projectPath);
    await configsMigration(migrationContext);

    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx")));

    const trueEnvContent_dev = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "config.dev"
    );
    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx", ".env.dev")));
    const testEnvContent_dev = await readEnvFile(path.join(projectPath, "teamsfx"), "dev");
    assert.equal(testEnvContent_dev, trueEnvContent_dev);

    const trueEnvContent_local = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "config.local"
    );
    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx", ".env.local")));
    const testEnvContent_local = await readEnvFile(path.join(projectPath, "teamsfx"), "local");
    assert.equal(testEnvContent_local, trueEnvContent_local);
  });
});

describe("userdataMigration", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("happy path for userdata migration", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    await copyTestProject(Constants.happyPathTestProject, projectPath);
    await userdataMigration(migrationContext);

    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx")));

    const trueEnvContent_dev = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "userdata.dev"
    );
    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx", ".env.dev")));
    const testEnvContent_dev = await readEnvFile(path.join(projectPath, "teamsfx"), "dev");
    assert.equal(testEnvContent_dev, trueEnvContent_dev);

    const trueEnvContent_local = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "userdata.local"
    );
    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx", ".env.local")));
    const testEnvContent_local = await readEnvFile(path.join(projectPath, "teamsfx"), "local");
    assert.equal(testEnvContent_local, trueEnvContent_local);
  });
});

describe("generateApimPluginEnvContent", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  const sandbox = sinon.createSandbox();

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
  });

  it("happy path", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    await copyTestProject(Constants.happyPathTestProject, projectPath);
    await generateApimPluginEnvContent(migrationContext);

    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx")));

    const trueEnvContent_dev = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "apimPlugin.dev"
    );
    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx", ".env.dev")));
    const testEnvContent_dev = await readEnvFile(path.join(projectPath, "teamsfx"), "dev");
    assert.equal(testEnvContent_dev, trueEnvContent_dev);
  });

  it("checkapimPluginExists: apim exists", () => {
    const pjSettings_1 = {
      appName: "testapp",
      components: [
        {
          name: "teams-tab",
        },
        {
          name: "apim",
        },
      ],
    };
    assert.isTrue(checkapimPluginExists(pjSettings_1));
  });

  it("checkapimPluginExists: apim not exists", () => {
    const pjSettings_2 = {
      appName: "testapp",
      components: [
        {
          name: "teams-tab",
        },
      ],
    };
    assert.isFalse(checkapimPluginExists(pjSettings_2));
  });

  it("checkapimPluginExists: components not exists", () => {
    const pjSettings_3 = {
      appName: "testapp",
    };
    assert.isFalse(checkapimPluginExists(pjSettings_3));
  });

  it("checkapimPluginExists: obj null", () => {
    const pjSettings_4 = null;
    assert.isFalse(checkapimPluginExists(pjSettings_4));
  });
});

describe("allEnvMigration", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("happy path for all env migration", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    await copyTestProject(Constants.happyPathTestProject, projectPath);
    await configsMigration(migrationContext);
    await statesMigration(migrationContext);
    await userdataMigration(migrationContext);
    await generateApimPluginEnvContent(migrationContext);

    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx")));

    const trueEnvContent_dev = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "all.dev"
    );
    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx", ".env.dev")));
    const testEnvContent_dev = await readEnvFile(path.join(projectPath, "teamsfx"), "dev");
    assert.equal(testEnvContent_dev, trueEnvContent_dev);

    const trueEnvContent_local = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "all.local"
    );
    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx", ".env.local")));
    const testEnvContent_local = await readEnvFile(path.join(projectPath, "teamsfx"), "local");
    assert.equal(testEnvContent_local, trueEnvContent_local);
  });
});

describe("Migration utils", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  const sandbox = sinon.createSandbox();

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
  });

  it("checkVersionForMigration V2", async () => {
    const migrationContext = await mockMigrationContext(projectPath);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    const state = await checkVersionForMigration(migrationContext);
    assert.equal(state.state, VersionState.upgradeable);
  });

  it("checkVersionForMigration V3", async () => {
    const migrationContext = await mockMigrationContext(projectPath);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readJson").resolves("3.0.0");
    const state = await checkVersionForMigration(migrationContext);
    assert.equal(state.state, VersionState.compatible);
  });

  it("checkVersionForMigration empty", async () => {
    const migrationContext = await mockMigrationContext(projectPath);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    sandbox.stub(fs, "pathExists").resolves(false);
    const state = await checkVersionForMigration(migrationContext);
    assert.equal(state.state, VersionState.unsupported);
  });

  it("UpgradeCanceledError", () => {
    const err = UpgradeCanceledError();
    assert.isNotNull(err);
  });

  it("getTrackingIdFromPath: V2 ", async () => {
    sandbox.stub(fs, "pathExists").callsFake(async (path: string) => {
      if (path === getProjectSettingPathV3(projectPath)) {
        return false;
      }
      return true;
    });
    sandbox.stub(fs, "readJson").resolves({ projectId: MetadataV2.projectMaxVersion });
    const trackingId = await getTrackingIdFromPath(projectPath);
    assert.equal(trackingId, MetadataV2.projectMaxVersion);
  });

  it("getTrackingIdFromPath: V3 ", async () => {
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readJson").resolves({ trackingId: MetadataV3.projectVersion });
    const trackingId = await getTrackingIdFromPath(projectPath);
    assert.equal(trackingId, MetadataV3.projectVersion);
  });

  it("getTrackingIdFromPath: empty", async () => {
    sandbox.stub(fs, "pathExists").resolves(false);
    const trackingId = await getTrackingIdFromPath(projectPath);
    assert.equal(trackingId, "");
  });

  it("getTrackingIdFromPath: empty", async () => {
    sandbox.stub(fs, "pathExists").resolves(false);
    const trackingId = await getTrackingIdFromPath(projectPath);
    assert.equal(trackingId, "");
  });

  it("getVersionState", () => {
    assert.equal(getVersionState("2.0.0"), VersionState.upgradeable);
    assert.equal(getVersionState("3.0.0"), VersionState.compatible);
    assert.equal(getVersionState("4.0.0"), VersionState.unsupported);
  });

  it("getDownloadLinkByVersionAndPlatform", () => {
    assert.equal(
      getDownloadLinkByVersionAndPlatform("2.0.0", Platform.VS),
      `${Metadata.versionMatchLink}#visual-studio`
    );
    assert.equal(
      getDownloadLinkByVersionAndPlatform("2.0.0", Platform.CLI),
      `${Metadata.versionMatchLink}#cli`
    );
    assert.equal(
      getDownloadLinkByVersionAndPlatform("2.0.0", Platform.VSCode),
      `${Metadata.versionMatchLink}#vscode`
    );
  });

  it("outputCancelMessage", () => {
    outputCancelMessage("2.0.0", Platform.VS);
    outputCancelMessage("2.0.0", Platform.CLI);
    outputCancelMessage("2.0.0", Platform.VSCode);
  });

  it("migrationNotificationMessage", () => {
    const tools = new MockTools();
    setTools(tools);

    const version: VersionForMigration = {
      currentVersion: "2.0.0",
      state: VersionState.upgradeable,
      platform: Platform.VS,
    };

    migrationNotificationMessage(version);
    version.platform = Platform.VSCode;
    migrationNotificationMessage(version);
    version.platform = Platform.CLI;
    migrationNotificationMessage(version);
  });

  it("isMigrationV3Enabled", () => {
    const enabled = isMigrationV3Enabled();
    assert.isFalse(enabled);
  });
});

describe("debugMigration", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  let runTabScript = "";
  let runAuthScript = "";
  let runBotScript = "";
  let runFunctionScript = "";

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
    sinon.stub(debugV3MigrationUtils, "updateLocalEnv").callsFake(async () => {});
    sinon
      .stub(debugV3MigrationUtils, "saveRunScript")
      .callsFake(async (context, filename, script) => {
        if (filename === "run.tab.js") {
          runTabScript = script;
        } else if (filename === "run.auth.js") {
          runAuthScript = script;
        } else if (filename === "run.bot.js") {
          runBotScript = script;
        } else if (filename === "run.api.js") {
          runFunctionScript = script;
        }
      });
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sinon.restore();
    runTabScript = "";
    runBotScript = "";
    runFunctionScript = "";
  });

  const testCases = [
    "transparent-tab",
    "transparent-sso-tab",
    "transparent-bot",
    "transparent-sso-bot",
    "transparent-notification",
    "transparent-tab-bot-func",
    "beforeV3.4.0-tab",
    "beforeV3.4.0-bot",
    "beforeV3.4.0-tab-bot-func",
    "V3.5.0-V4.0.6-tab",
    "V3.5.0-V4.0.6-tab-bot-func",
    "V3.5.0-V4.0.6-notification-trigger",
    "V3.5.0-V4.0.6-command",
  ];

  testCases.forEach((testCase) => {
    it(testCase, async () => {
      const migrationContext = await mockMigrationContext(projectPath);

      await copyTestProject(path.join("debug", testCase), projectPath);

      await debugMigration(migrationContext);

      assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx")));
      assert.equal(
        await fs.readFile(path.join(projectPath, "teamsfx", "app.local.yml"), "utf-8"),
        await fs.readFile(path.join(projectPath, "expected", "app.local.yml"), "utf-8")
      );
      assert.equal(
        await fs.readFile(path.join(projectPath, ".vscode", "tasks.json"), "utf-8"),
        await fs.readFile(path.join(projectPath, "expected", "tasks.json"), "utf-8")
      );

      const runTabScriptPath = path.join(projectPath, "expected", "run.tab.js");
      if (await fs.pathExists(runTabScriptPath)) {
        assert.equal(runTabScript, await fs.readFile(runTabScriptPath, "utf-8"));
      }
      const runBotScriptPath = path.join(projectPath, "expected", "run.bot.js");
      if (await fs.pathExists(runBotScriptPath)) {
        assert.equal(runBotScript, await fs.readFile(runBotScriptPath, "utf-8"));
      }
      const runFunctionScriptPath = path.join(projectPath, "expected", "run.api.js");
      if (await fs.pathExists(runFunctionScriptPath)) {
        assert.equal(runFunctionScript, await fs.readFile(runFunctionScriptPath, "utf-8"));
      }
    });
  });
});

export async function mockMigrationContext(projectPath: string): Promise<MigrationContext> {
  const inputs: Inputs = { platform: Platform.VSCode, ignoreEnvInfo: true };
  inputs.projectPath = projectPath;
  const ctx = {
    arguments: [inputs],
  };
  return await MigrationContext.create(ctx);
}

function getTestAssetsPath(projectName: string): string {
  return path.join("tests/core/middleware/testAssets/v3Migration", projectName.toString());
}

// Change CRLF to LF to avoid test failures in different OS
function normalizeLineBreaks(content: string): string {
  return content.replace(/\r\n/g, "\n");
}

async function assertFileContent(
  projectPath: string,
  actualFilePath: string,
  expectedFileName: string
): Promise<void> {
  const actualFileFullPath = path.join(projectPath, actualFilePath);
  const expectedFileFulePath = path.join(projectPath, "expectedResult", expectedFileName);
  assert.isTrue(await fs.pathExists(actualFileFullPath));
  const actualFileContent = normalizeLineBreaks(await fs.readFile(actualFileFullPath, "utf8"));
  const expectedFileContent = normalizeLineBreaks(await fs.readFile(expectedFileFulePath, "utf8"));
  assert.equal(actualFileContent, expectedFileContent);
}

async function copyTestProject(projectName: string, targetPath: string): Promise<void> {
  await fs.copy(getTestAssetsPath(projectName), targetPath);
}

async function readOldProjectSettings(projectPath: string): Promise<any> {
  return await fs.readJson(path.join(projectPath, Constants.oldProjectSettingsFilePath));
}

async function readSettingJson(projectPath: string): Promise<any> {
  return await fs.readJson(path.join(projectPath, Constants.settingsFilePath));
}

async function readEnvFile(projectPath: string, env: string): Promise<any> {
  return await fs.readFileSync(path.join(projectPath, ".env." + env)).toString();
}

function getAction(lifecycleDefinition: Array<any>, actionName: string): any[] {
  if (lifecycleDefinition) {
    return lifecycleDefinition.filter((item) => item.uses === actionName);
  }
  return [];
}

const Constants = {
  happyPathTestProject: "happyPath",
  settingsFilePath: "teamsfx/settings.json",
  oldProjectSettingsFilePath: ".fx/configs/projectSettings.json",
  appYmlPath: "teamsfx/app.yml",
  manifestsMigrationHappyPath: "manifestsHappyPath",
  manifestsMigrationHappyPathSpfx: "manifestsHappyPathSpfx",
  launchJsonPath: ".vscode/launch.json",
  happyPathWithoutFx: "happyPath_for_needMigrateToAadManifest/happyPath_no_fx",
  happyPathAadManifestTemplateExist:
    "happyPath_for_needMigrateToAadManifest/happyPath_aadManifestTemplateExist",
  happyPathWithoutPermission: "happyPath_for_needMigrateToAadManifest/happyPath_no_permissionFile",
  happyPathAadPluginNotActive:
    "happyPath_for_needMigrateToAadManifest/happyPath_aadPluginNotActive",
};
