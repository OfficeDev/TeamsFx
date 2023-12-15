/**
 * @author HuihuiWu-Microsoft <73154171+HuihuiWu-Microsoft@users.noreply.github.com>
 */
import * as chai from "chai";
import * as fs from "fs-extra";
import * as path from "path";
import * as sinon from "sinon";
import * as uuid from "uuid";
import * as vscode from "vscode";
import * as mockfs from "mock-fs";

import {
  ConfigFolderName,
  FxError,
  Inputs,
  ManifestUtil,
  OptionItem,
  Platform,
  Result,
  Stage,
  SystemError,
  UserError,
  Void,
  VsCodeEnv,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import * as commonTools from "@microsoft/teamsfx-core/build/common/tools";
import * as globalState from "@microsoft/teamsfx-core/build/common/globalState";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";
import {
  AppDefinition,
  AppStudioClient,
  CollaborationState,
  DepsManager,
  DepsType,
  FxCore,
  UnhandledError,
  UserCancelError,
  environmentManager,
  manifestUtils,
  pathUtils,
} from "@microsoft/teamsfx-core";
import commandController from "../../src/commandController";
import { AzureAccountManager } from "../../src/commonlib/azureLogin";
import { signedIn, signedOut } from "../../src/commonlib/common/constant";
import { VsCodeLogProvider } from "../../src/commonlib/log";
import M365TokenInstance from "../../src/commonlib/m365Login";
import { DeveloperPortalHomeLink, GlobalKey } from "../../src/constants";
import { PanelType } from "../../src/controls/PanelType";
import { WebviewPanel } from "../../src/controls/webviewPanel";
import * as debugCommonUtils from "../../src/debug/commonUtils";
import * as debugConstants from "../../src/debug/constants";
import * as launch from "../../src/debug/launch";
import { ExtensionErrors } from "../../src/error";
import { TreatmentVariableValue } from "../../src/exp/treatmentVariables";
import * as extension from "../../src/extension";
import * as globalVariables from "../../src/globalVariables";
import * as handlers from "../../src/handlers";
import { ProgressHandler } from "../../src/progressHandler";
import { VsCodeUI } from "../../src/qm/vsc_ui";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import * as extTelemetryEvents from "../../src/telemetry/extTelemetryEvents";
import accountTreeViewProviderInstance from "../../src/treeview/account/accountTreeViewProvider";
import envTreeProviderInstance from "../../src/treeview/environmentTreeViewProvider";
import TreeViewManagerInstance from "../../src/treeview/treeViewManager";
import * as commonUtils from "../../src/utils/commonUtils";
import * as localizeUtils from "../../src/utils/localizeUtils";
import { ExtensionSurvey } from "../../src/utils/survey";
import { MockCore } from "../mocks/mockCore";
import VsCodeLogInstance from "../../src/commonlib/log";
import * as localPrerequisites from "../../src/debug/prerequisitesHandler";
import { TeamsAppMigrationHandler } from "../../src/migration/migrationHandler";

describe("handlers", () => {
  describe("activate()", function () {
    const sandbox = sinon.createSandbox();

    beforeEach(() => {
      sandbox.stub(accountTreeViewProviderInstance, "subscribeToStatusChanges");
      sandbox.stub(vscode.extensions, "getExtension").returns(undefined);
      sandbox.stub(TreeViewManagerInstance, "getTreeView").returns(undefined);
      sandbox.stub(ExtTelemetry, "dispose");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("No globalState error", async () => {
      const result = await handlers.activate();
      chai.assert.deepEqual(result.isOk() ? result.value : result.error.name, {});
    });

    it("Valid project", async () => {
      sandbox.stub(projectSettingsHelper, "isValidProject").returns(true);
      const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const addSharedPropertyStub = sandbox.stub(ExtTelemetry, "addSharedProperty");
      const setCommandIsRunningStub = sandbox.stub(globalVariables, "setCommandIsRunning");
      const lockedByOperationStub = sandbox.stub(commandController, "lockedByOperation");
      const unlockedByOperationStub = sandbox.stub(commandController, "unlockedByOperation");
      const azureAccountSetStatusChangeMapStub = sandbox.stub(
        AzureAccountManager.prototype,
        "setStatusChangeMap"
      );
      const m365AccountSetStatusChangeMapStub = sandbox.stub(
        M365TokenInstance,
        "setStatusChangeMap"
      );
      const showMessageStub = sandbox.stub(vscode.window, "showInformationMessage");
      let lockCallback: any;
      let unlockCallback: any;

      sandbox.stub(FxCore.prototype, "on").callsFake((event: string, callback: any) => {
        if (event === "lock") {
          lockCallback = callback;
        } else {
          unlockCallback = callback;
        }
      });
      azureAccountSetStatusChangeMapStub.callsFake(
        (
          name: string,
          statusChange: (
            status: string,
            token?: string,
            accountInfo?: Record<string, unknown>
          ) => Promise<void>,
          immediateCall?: boolean
        ) => {
          statusChange(signedIn).then(() => {});
          statusChange(signedOut).then(() => {});
          return Promise.resolve(true);
        }
      );
      m365AccountSetStatusChangeMapStub.callsFake(
        (
          name: string,
          tokenRequest: unknown,
          statusChange: (
            status: string,
            token?: string,
            accountInfo?: Record<string, unknown>
          ) => Promise<void>,
          immediateCall?: boolean
        ) => {
          statusChange(signedIn).then(() => {});
          statusChange(signedOut).then(() => {});
          return Promise.resolve(ok(true));
        }
      );
      const result = await handlers.activate();

      chai.assert.isTrue(addSharedPropertyStub.called);
      chai.assert.isTrue(sendTelemetryStub.calledOnceWith("open-teams-app"));
      chai.assert.deepEqual(result.isOk() ? result.value : result.error.name, {});

      lockCallback("test");
      setCommandIsRunningStub.calledOnceWith(true);
      lockedByOperationStub.calledOnceWith("test");

      unlockCallback("test");
      unlockedByOperationStub.calledOnceWith("test");

      chai.assert.isTrue(showMessageStub.called);
    });

    it("throws error", async () => {
      sandbox.stub(projectSettingsHelper, "isValidProject").returns(false);
      sandbox.stub(M365TokenInstance, "setStatusChangeMap");
      sandbox.stub(FxCore.prototype, "on").throws(new Error("test"));
      const showErrorMessageStub = sinon.stub(vscode.window, "showErrorMessage");

      const result = await handlers.activate();

      chai.assert.isTrue(result.isErr());
      chai.assert.isTrue(showErrorMessageStub.called);
    });
  });
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });

  it("getSystemInputs()", () => {
    const input: Inputs = handlers.getSystemInputs();

    chai.expect(input.platform).equals(Platform.VSCode);
  });

  it("getSettingsVersion", async () => {
    sandbox.stub(handlers, "core").value(new MockCore());
    sandbox.stub(handlers, "getSystemInputs").returns({} as Inputs);
    sandbox
      .stub(MockCore.prototype, "projectVersionCheck")
      .resolves(ok({ currentVersion: "3.0.0" }));
    const res = await handlers.getSettingsVersion();
    chai.assert.equal(res, "3.0.0");
  });

  it("addFileSystemWatcher detect SPFx project", async () => {
    const workspacePath = "test";
    const isValidProject = sandbox.stub(projectSettingsHelper, "isValidProject").returns(true);
    const initGlobalVariables = sandbox.stub(globalVariables, "initializeGlobalVariables");
    const updateTreeViewsOnSPFxChanged = sandbox.stub(
      TreeViewManagerInstance,
      "updateTreeViewsOnSPFxChanged"
    );

    const watcher = {
      onDidCreate: () => ({ dispose: () => undefined }),
      onDidChange: () => ({ dispose: () => undefined }),
      onDidDelete: () => ({ dispose: () => undefined }),
    } as any;
    const createWatcher = sandbox
      .stub(vscode.workspace, "createFileSystemWatcher")
      .returns(watcher);
    const createListener = sandbox.stub(watcher, "onDidCreate").callsFake((...args: unknown[]) => {
      (args as any)[0]();
    });
    const changeListener = sandbox.stub(watcher, "onDidChange").callsFake((...args: unknown[]) => {
      (args as any)[0]();
    });
    const deleteListener = sandbox.stub(watcher, "onDidDelete").callsFake((...args: unknown[]) => {
      (args as any)[0]();
    });
    const sendTelemetryEventFunc = sandbox
      .stub(ExtTelemetry, "sendTelemetryEvent")
      .callsFake(() => {});

    handlers.addFileSystemWatcher(workspacePath);

    chai.assert.equal(createWatcher.callCount, 2);
    chai.assert.equal(createListener.callCount, 2);
    chai.assert.isTrue(changeListener.calledTwice);
  });

  it("addFileSystemWatcher in invalid project", async () => {
    const workspacePath = "test";
    const isValidProject = sandbox.stub(projectSettingsHelper, "isValidProject").returns(false);

    const watcher = {
      onDidCreate: () => ({ dispose: () => undefined }),
      onDidChange: () => ({ dispose: () => undefined }),
    } as any;
    const createWatcher = sandbox
      .stub(vscode.workspace, "createFileSystemWatcher")
      .returns(watcher);
    const createListener = sandbox.stub(watcher, "onDidCreate").resolves();
    const changeListener = sandbox.stub(watcher, "onDidChange").resolves();

    handlers.addFileSystemWatcher(workspacePath);

    chai.assert.isTrue(createWatcher.notCalled);
    chai.assert.isTrue(createListener.notCalled);
    chai.assert.isTrue(changeListener.notCalled);
  });

  it("sendSDKVersionTelemetry", async () => {
    const filePath = "test/package-lock.json";

    const readJsonFunc = sandbox.stub(fs, "readJson").resolves();
    const sendTelemetryEventFunc = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

    handlers.sendSDKVersionTelemetry(filePath);

    chai.assert.isTrue(readJsonFunc.calledOnce);
  });

  it("updateAutoOpenGlobalKey", async () => {
    sandbox.stub(commonUtils, "isTriggerFromWalkThrough").returns(true);
    sandbox.stub(globalVariables, "checkIsSPFx").returns(true);
    const globalStateUpdateStub = sinon.stub(globalState, "globalStateUpdate");

    await handlers.updateAutoOpenGlobalKey(false, vscode.Uri.file("test"), [
      { type: "type", content: "content" },
    ]);

    chai.assert.isTrue(globalStateUpdateStub.callCount === 4);
  });

  describe("command handlers", function () {
    this.afterEach(() => {
      sinon.restore();
    });

    it("createNewProjectHandler()", async () => {
      const clock = sinon.useFakeTimers();

      sinon.stub(handlers, "core").value(new MockCore());
      const sendTelemetryEventFunc = sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(globalVariables, "checkIsSPFx").returns(false);
      const createProject = sinon.spy(handlers.core, "createProject");
      const executeCommandFunc = sinon.stub(vscode.commands, "executeCommand");

      await handlers.createNewProjectHandler();

      chai.assert.isTrue(
        sendTelemetryEventFunc.calledWith(extTelemetryEvents.TelemetryEvent.CreateProjectStart)
      );
      chai.assert.isTrue(
        sendTelemetryEventFunc.calledWith(extTelemetryEvents.TelemetryEvent.CreateProject)
      );
      sinon.assert.calledOnce(createProject);
      chai.assert.isTrue(executeCommandFunc.calledOnceWith("vscode.openFolder"));
      sinon.restore();
      clock.restore();
    });

    it("provisionHandler()", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const provisionResources = sinon.spy(handlers.core, "provisionResources");
      sinon.stub(envTreeProviderInstance, "reloadEnvironments");

      await handlers.provisionHandler();

      sinon.assert.calledOnce(provisionResources);
      sinon.restore();
    });

    it("deployHandler()", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const deployArtifacts = sinon.spy(handlers.core, "deployArtifacts");

      await handlers.deployHandler();

      sinon.assert.calledOnce(deployArtifacts);
      sinon.restore();
    });

    it("publishHandler()", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const publishApplication = sinon.spy(handlers.core, "publishApplication");

      await handlers.publishHandler();

      sinon.assert.calledOnce(publishApplication);
      sinon.restore();
    });

    it("buildPackageHandler()", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      const sendTelemetryErrorEvent = sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");

      await handlers.buildPackageHandler();

      // should show error for invalid project
      sinon.assert.calledOnce(sendTelemetryErrorEvent);
      sinon.restore();
    });

    it("validateManifestHandler() - app package", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(localizeUtils, "localize").returns("");
      sinon.stub(projectSettingsHelper, "isValidProject").returns(true);
      sinon.stub(handlers, "getSystemInputs").returns({} as Inputs);
      const validateApplication = sinon.spy(handlers.core, "validateApplication");

      sinon.stub(extension, "VS_CODE_UI").value({
        selectOption: () => {
          return Promise.resolve(ok({ type: "success", result: "validateAgainstPackage" }));
        },
      });

      await handlers.validateManifestHandler();
      sinon.assert.calledOnce(validateApplication);
    });

    it("copilotPluginAddAPIHandler()", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      const addAPIHanlder = sinon.spy(handlers.core, "copilotPluginAddAPI");
      const args = [
        {
          fsPath: "manifest.json",
        },
      ];

      await handlers.copilotPluginAddAPIHandler(args);

      sinon.assert.calledOnce(addAPIHanlder);
    });

    it("treeViewPreviewHandler() - previewWithManifest error", async () => {
      sinon.stub(localizeUtils, "localize").returns("");
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sandbox.stub(handlers, "getSystemInputs").returns({} as Inputs);
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(handlers.core, "previewWithManifest").resolves(err({ foo: "bar" } as any));

      const result = await handlers.treeViewPreviewHandler("dev");

      chai.assert.isTrue(result.isErr());
    });

    it("treeViewPreviewHandler() - happy path", async () => {
      sinon.stub(localizeUtils, "localize").returns("");
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sandbox.stub(handlers, "getSystemInputs").returns({} as Inputs);
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(handlers.core, "previewWithManifest").resolves(ok("test-url"));
      sandbox.stub(launch, "openHubWebClient").resolves();

      const result = await handlers.treeViewPreviewHandler("dev");

      chai.assert.isTrue(result.isOk());
    });

    it("selectTutorialsHandler()", async () => {
      sinon.stub(localizeUtils, "localize").returns("");
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(TreatmentVariableValue, "inProductDoc").value(true);
      sinon.stub(globalVariables, "isSPFxProject").value(false);
      let tutorialOptions: OptionItem[] = [];
      sinon.stub(extension, "VS_CODE_UI").value({
        selectOption: (options: any) => {
          tutorialOptions = options.options;
          return Promise.resolve(ok({ type: "success", result: { id: "test", data: "data" } }));
        },
        openUrl: () => Promise.resolve(ok(true)),
      });

      const result = await handlers.selectTutorialsHandler();

      chai.assert.equal(tutorialOptions.length, 17);
      chai.assert.isTrue(result.isOk());
      chai.assert.equal(tutorialOptions[1].data, "https://aka.ms/teamsfx-notification-new");
    });

    it("selectTutorialsHandler() for SPFx projects - v3", async () => {
      sinon.stub(localizeUtils, "localize").returns("");
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(TreatmentVariableValue, "inProductDoc").value(true);
      sinon.stub(globalVariables, "isSPFxProject").value(true);
      let tutorialOptions: OptionItem[] = [];
      sinon.stub(extension, "VS_CODE_UI").value({
        selectOption: (options: any) => {
          tutorialOptions = options.options;
          return Promise.resolve(ok({ type: "success", result: { id: "test", data: "data" } }));
        },
        openUrl: () => Promise.resolve(ok(true)),
      });

      const result = await handlers.selectTutorialsHandler();

      chai.assert.equal(tutorialOptions.length, 1);
      chai.assert.isTrue(result.isOk());
      chai.assert.equal(tutorialOptions[0].data, "https://aka.ms/teamsfx-add-cicd-new");
    });
  });

  it("openAccountHelpHandler()", async () => {
    const createOrShow = sandbox.stub(WebviewPanel, "createOrShow");
    handlers.openAccountHelpHandler();
    sandbox.assert.calledOnceWithExactly(createOrShow, PanelType.AccountHelp);
  });

  describe("runCommand()", function () {
    this.afterEach(() => {
      sinon.restore();
    });
    it("openConfigStateFile() - InvalidArgs", async () => {
      const env = "local";
      const tmpDir = fs.mkdtempSync(path.resolve("./tmp"));

      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");

      sinon.stub(globalVariables, "workspaceUri").value(vscode.Uri.file(tmpDir));
      const projectSettings: any = {
        appName: "myapp",
        version: "1.0.0",
        projectId: "123",
      };
      const configFolder = path.resolve(tmpDir, `.${ConfigFolderName}`, "configs");
      await fs.mkdir(configFolder, { recursive: true });
      const settingsFile = path.resolve(configFolder, "projectSettings.json");
      await fs.writeJSON(settingsFile, JSON.stringify(projectSettings, null, 4));

      sinon.stub(globalVariables, "context").value({ extensionPath: path.resolve("../../") });
      sinon.stub(extension, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: env })),
      });

      const res = await handlers.openConfigStateFile([]);
      await fs.remove(tmpDir);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.InvalidArgs);
      }
    });

    it("openConfigStateFile() - noOpenWorkspace", async () => {
      const env = "local";

      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");

      sinon.stub(globalVariables, "workspaceUri").value({ fsPath: undefined });

      sinon.stub(globalVariables, "context").value({ extensionPath: path.resolve("../../") });
      sinon.stub(extension, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: env })),
      });

      const res = await handlers.openConfigStateFile([]);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.NoWorkspaceError);
      }
    });

    it("openConfigStateFile() - invalidProject", async () => {
      const env = "local";
      const tmpDir = fs.mkdtempSync(path.resolve("./tmp"));

      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(projectSettingsHelper, "isValidProject").returns(false);

      sinon.stub(globalVariables, "workspaceUri").value(vscode.Uri.file(tmpDir));
      sinon.stub(globalVariables, "context").value({ extensionPath: path.resolve("../../") });
      sinon.stub(extension, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: env })),
      });

      const res = await handlers.openConfigStateFile([]);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.InvalidProject);
      }
    });

    it("openConfigStateFile() - invalid target environment", async () => {
      const env = "local";
      const tmpDir = fs.mkdtempSync(path.resolve("./tmp"));

      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");

      sinon.stub(globalVariables, "workspaceUri").value(vscode.Uri.file(tmpDir));
      const projectSettings: any = {
        appName: "myapp",
        version: "1.0.0",
        projectId: "123",
      };
      const configFolder = path.resolve(tmpDir, `.${ConfigFolderName}`, "configs");
      await fs.mkdir(configFolder, { recursive: true });
      const settingsFile = path.resolve(configFolder, "projectSettings.json");
      await fs.writeJSON(settingsFile, JSON.stringify(projectSettings, null, 4));

      sinon.stub(globalVariables, "context").value({ extensionPath: path.resolve("../../") });
      sinon.stub(extension, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(err({ error: "invalid target env" })),
      });
      sinon.stub(environmentManager, "listAllEnvConfigs").resolves(ok([]));
      sinon.stub(fs, "pathExists").resolves(false);
      sinon.stub(pathUtils, "getEnvFolderPath").resolves(ok(env));

      const res = await handlers.openConfigStateFile([{ env: undefined, type: "env" }]);
      await fs.remove(tmpDir);

      if (res) {
        chai.assert.isTrue(res.isErr());
      }
    });

    it("openConfigStateFile() - valid args", async () => {
      const env = "local";
      const tmpDir = fs.mkdtempSync(path.resolve("./tmp"));

      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");

      sinon.stub(globalVariables, "workspaceUri").value(vscode.Uri.file(tmpDir));
      const projectSettings: any = {
        appName: "myapp",
        version: "1.0.0",
        projectId: "123",
      };
      const configFolder = path.resolve(tmpDir, `.${ConfigFolderName}`, "configs");
      await fs.mkdir(configFolder, { recursive: true });
      const settingsFile = path.resolve(configFolder, "projectSettings.json");
      await fs.writeJSON(settingsFile, JSON.stringify(projectSettings, null, 4));

      sinon.stub(globalVariables, "context").value({ extensionPath: path.resolve("../../") });
      sinon.stub(extension, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: env })),
      });
      sinon.stub(pathUtils, "getEnvFolderPath").resolves(ok(env));
      sinon.stub(fs, "pathExists").resolves(false);
      sinon.stub(environmentManager, "listAllEnvConfigs").resolves(ok([]));

      const res = await handlers.openConfigStateFile([{ env: undefined, type: "env" }]);
      await fs.remove(tmpDir);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.EnvFileNotFoundError);
      }
    });

    it("openConfigStateFile() - invalid env folder", async () => {
      const env = "local";
      const tmpDir = fs.mkdtempSync(path.resolve("./tmp"));

      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");

      sinon.stub(globalVariables, "workspaceUri").value(vscode.Uri.file(tmpDir));
      const projectSettings: any = {
        appName: "myapp",
        version: "1.0.0",
        projectId: "123",
      };
      const configFolder = path.resolve(tmpDir, `.${ConfigFolderName}`, "configs");
      await fs.mkdir(configFolder, { recursive: true });
      const settingsFile = path.resolve(configFolder, "projectSettings.json");
      await fs.writeJSON(settingsFile, JSON.stringify(projectSettings, null, 4));

      sinon.stub(globalVariables, "context").value({ extensionPath: path.resolve("../../") });
      sinon.stub(extension, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: env })),
      });
      sinon.stub(pathUtils, "getEnvFolderPath").resolves(err({ error: "unknown" } as any));
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(vscode.workspace, "openTextDocument").resolves("" as any);

      const res = await handlers.openConfigStateFile([{ env: env, type: "env" }]);
      await fs.remove(tmpDir);

      if (res) {
        chai.assert.isTrue(res.isErr());
      }
    });

    it("openConfigStateFile() - success", async () => {
      const env = "local";
      const tmpDir = fs.mkdtempSync(path.resolve("./tmp"));

      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");

      sinon.stub(globalVariables, "workspaceUri").value(vscode.Uri.file(tmpDir));
      const projectSettings: any = {
        appName: "myapp",
        version: "1.0.0",
        projectId: "123",
      };
      const configFolder = path.resolve(tmpDir, `.${ConfigFolderName}`, "configs");
      await fs.mkdir(configFolder, { recursive: true });
      const settingsFile = path.resolve(configFolder, "projectSettings.json");
      await fs.writeJSON(settingsFile, JSON.stringify(projectSettings, null, 4));

      sinon.stub(globalVariables, "context").value({ extensionPath: path.resolve("../../") });
      sinon.stub(extension, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: env })),
      });
      sinon.stub(pathUtils, "getEnvFolderPath").resolves(ok(env));
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(vscode.workspace, "openTextDocument").returns(Promise.resolve("" as any));

      const res = await handlers.openConfigStateFile([{ env: env, type: "env" }]);
      await fs.remove(tmpDir);

      if (res) {
        chai.assert.isTrue(res.isOk());
      }
    });

    it("create sample with projectid", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      const sendTelemetryEvent = sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const createProject = sinon.spy(handlers.core, "createProject");
      sinon.stub(vscode.commands, "executeCommand");
      const inputs = { projectId: uuid.v4(), platform: Platform.VSCode };

      await handlers.runCommand(Stage.create, inputs);

      sinon.assert.calledOnce(createProject);
      chai.assert.isTrue(createProject.args[0][0].projectId != undefined);
      chai.assert.isTrue(sendTelemetryEvent.args[0][1]!["new-project-id"] != undefined);
    });

    it("create from scratch without projectid", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      const sendTelemetryEvent = sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const createProject = sinon.spy(handlers.core, "createProject");
      sinon.stub(vscode.commands, "executeCommand");

      await handlers.runCommand(Stage.create);

      sinon.restore();
      sinon.assert.calledOnce(createProject);
      chai.assert.isTrue(createProject.args[0][0].projectId != undefined);
      chai.assert.isTrue(sendTelemetryEvent.args[0][1]!["new-project-id"] != undefined);
    });

    it("provisionResources", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const provisionResources = sinon.spy(handlers.core, "provisionResources");

      await handlers.runCommand(Stage.provision);

      sinon.restore();
      sinon.assert.calledOnce(provisionResources);
    });

    it("provisionResources - local", async () => {
      const mockCore = new MockCore();
      const mockCoreStub = sinon
        .stub(mockCore, "provisionResources")
        .resolves(err(new UserError("test", "test", "test")));
      sinon.stub(handlers, "core").value(mockCore);
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");

      const res = await handlers.runCommand(Stage.provision, {
        platform: Platform.VSCode,
        env: "local",
      } as Inputs);
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(
          res.error.recommendedOperation,
          debugConstants.RecommendedOperations.DebugInTestTool
        );
      }
      sinon.restore();
      sinon.assert.calledOnce(mockCoreStub);
    });

    it("deployArtifacts", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const deployArtifacts = sinon.spy(handlers.core, "deployArtifacts");

      await handlers.runCommand(Stage.deploy);

      sinon.restore();
      sinon.assert.calledOnce(deployArtifacts);
    });

    it("deployArtifacts - local", async () => {
      const mockCore = new MockCore();
      const mockCoreStub = sinon
        .stub(mockCore, "deployArtifacts")
        .resolves(err(new UserError("test", "test", "test")));
      sinon.stub(handlers, "core").value(mockCore);
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");

      await handlers.runCommand(Stage.deploy, {
        platform: Platform.VSCode,
        env: "local",
      } as Inputs);

      sinon.restore();
      sinon.assert.calledOnce(mockCoreStub);
    });

    it("deployAadManifest", async () => {
      const sandbox = sinon.createSandbox();
      sandbox.stub(handlers, "core").value(new MockCore());
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const deployAadManifest = sandbox.spy(handlers.core, "deployAadManifest");
      const input: Inputs = handlers.getSystemInputs();
      await handlers.runCommand(Stage.deployAad, input);

      sandbox.assert.calledOnce(deployAadManifest);
      sandbox.restore();
    });

    it("deployAadManifest happy path", async () => {
      const sandbox = sinon.createSandbox();
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sandbox.stub(handlers.core, "deployAadManifest").resolves(ok(undefined));
      const input: Inputs = handlers.getSystemInputs();
      const res = await handlers.runCommand(Stage.deployAad, input);
      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.strictEqual(res.value, undefined);
      }
      sandbox.restore();
    });

    it("localDebug", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");

      let ignoreEnvInfo: boolean | undefined = undefined;
      let localDebugCalled = 0;
      sinon
        .stub(handlers.core, "localDebug")
        .callsFake(async (inputs: Inputs): Promise<Result<undefined, FxError>> => {
          ignoreEnvInfo = inputs.ignoreEnvInfo;
          localDebugCalled += 1;
          return ok(undefined);
        });

      await handlers.runCommand(Stage.debug);

      sinon.restore();
      chai.expect(ignoreEnvInfo).to.equal(false);
      chai.expect(localDebugCalled).equals(1);
    });

    it("publishApplication", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const publishApplication = sinon.spy(handlers.core, "publishApplication");

      await handlers.runCommand(Stage.publish);

      sinon.restore();
      sinon.assert.calledOnce(publishApplication);
    });

    it("createEnv", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const createEnv = sinon.spy(handlers.core, "createEnv");
      sinon.stub(vscode.commands, "executeCommand");

      await handlers.runCommand(Stage.createEnv);

      sinon.restore();
      sinon.assert.calledOnce(createEnv);
    });
  });

  describe("detectVsCodeEnv()", function () {
    this.afterEach(() => {
      sinon.restore();
    });

    it("locally run", () => {
      const expectedResult = {
        extensionKind: vscode.ExtensionKind.UI,
        id: "",
        extensionUri: vscode.Uri.file(""),
        extensionPath: "",
        isActive: true,
        packageJSON: {},
        exports: undefined,
        activate: sinon.spy(),
      };
      const getExtension = sinon
        .stub(vscode.extensions, "getExtension")
        .callsFake((name: string) => {
          return expectedResult;
        });

      chai.expect(handlers.detectVsCodeEnv()).equals(VsCodeEnv.local);
      getExtension.restore();
    });

    it("Remotely run", () => {
      const expectedResult = {
        extensionKind: vscode.ExtensionKind.Workspace,
        id: "",
        extensionUri: vscode.Uri.file(""),
        extensionPath: "",
        isActive: true,
        packageJSON: {},
        exports: undefined,
        activate: sinon.spy(),
      };
      const getExtension = sinon
        .stub(vscode.extensions, "getExtension")
        .callsFake((name: string) => {
          return expectedResult;
        });

      chai
        .expect(handlers.detectVsCodeEnv())
        .oneOf([VsCodeEnv.remote, VsCodeEnv.codespaceVsCode, VsCodeEnv.codespaceBrowser]);
      getExtension.restore();
    });
  });

  it("openWelcomeHandler", async () => {
    const executeCommands = sandbox.stub(vscode.commands, "executeCommand");
    const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

    await handlers.openWelcomeHandler();

    sandbox.assert.calledOnceWithExactly(
      executeCommands,
      "workbench.action.openWalkthrough",
      "TeamsDevApp.ms-teams-vscode-extension#teamsToolkitGetStarted"
    );
  });

  it("openSurveyHandler", async () => {
    const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const openLink = sandbox.stub(ExtensionSurvey.getInstance(), "openSurveyLink");
    sandbox.stub(localizeUtils, "getDefaultString").returns("test");

    await handlers.openSurveyHandler([extTelemetryEvents.TelemetryTriggerFrom.TreeView]);
    chai.assert.isTrue(sendTelemetryEvent.calledOnce);
    chai.assert.isTrue(openLink.calledOnce);
  });

  it("openSamplesHandler", async () => {
    const createOrShow = sandbox.stub(WebviewPanel, "createOrShow");
    const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

    await handlers.openSamplesHandler();

    sandbox.assert.calledOnceWithExactly(createOrShow, PanelType.SampleGallery, false);
  });

  it("openReadMeHandler", async () => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "isTeamsFxProject").value(true);
    const executeCommands = sandbox.stub(vscode.commands, "executeCommand");
    sandbox
      .stub(vscode.workspace, "workspaceFolders")
      .value([{ uri: { fsPath: "readmeTestFolder" } }]);
    sandbox.stub(fs, "pathExists").resolves(true);
    const openTextDocumentStub = sandbox
      .stub(vscode.workspace, "openTextDocument")
      .resolves({} as any as vscode.TextDocument);

    await handlers.openReadMeHandler([extTelemetryEvents.TelemetryTriggerFrom.Auto]);

    chai.assert.isTrue(openTextDocumentStub.calledOnce);
    chai.assert.isTrue(executeCommands.calledOnce);
  });

  it("openReadMeHandler - create project", async () => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "isTeamsFxProject").value(false);
    sandbox.stub(handlers, "core").value(undefined);
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .callsFake(
        (title: string, options: vscode.MessageOptions, ...items: vscode.MessageItem[]) => {
          return Promise.resolve({
            title: "Yes",
            run: (options as any).run,
          } as vscode.MessageItem);
        }
      );
    await handlers.openReadMeHandler([extTelemetryEvents.TelemetryTriggerFrom.Auto]);

    chai.assert.isTrue(showMessageStub.calledOnce);
  });

  it("openReadMeHandler - open folder", async () => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "isTeamsFxProject").value(false);
    sandbox.stub(handlers, "core").value(undefined);
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .callsFake(
        (title: string, options: vscode.MessageOptions, ...items: vscode.MessageItem[]) => {
          return Promise.resolve({
            title: "Yes",
            run: (items[0] as any).run,
          } as vscode.MessageItem);
        }
      );
    await handlers.openReadMeHandler([extTelemetryEvents.TelemetryTriggerFrom.Auto]);

    chai.assert.isTrue(executeCommandStub.calledOnce);
  });

  it("openReadMeHandler - function notification bot template", async () => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "isTeamsFxProject").value(true);
    sandbox
      .stub(vscode.workspace, "workspaceFolders")
      .value([{ uri: { fsPath: "readmeTestFolder" } }]);
    sandbox.stub(TreatmentVariableValue, "inProductDoc").value(true);
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").resolves(Buffer.from("## Get Started with the Notification bot"));
    const createOrShow = sandbox.stub(WebviewPanel, "createOrShow");

    await handlers.openReadMeHandler([extTelemetryEvents.TelemetryTriggerFrom.Auto]);

    sandbox.assert.calledOnceWithExactly(
      createOrShow,
      PanelType.FunctionBasedNotificationBotReadme
    );
  });

  it("openReadMeHandler - restify notification bot template", async () => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "isTeamsFxProject").value(true);
    sandbox
      .stub(vscode.workspace, "workspaceFolders")
      .value([{ uri: { fsPath: "readmeTestFolder" } }]);
    sandbox.stub(TreatmentVariableValue, "inProductDoc").value(true);
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox
      .stub(fs, "readFile")
      .resolves(Buffer.from("## Get Started with the Notification bot restify"));
    const createOrShow = sandbox.stub(WebviewPanel, "createOrShow");

    await handlers.openReadMeHandler([extTelemetryEvents.TelemetryTriggerFrom.Auto]);

    sandbox.assert.calledOnceWithExactly(
      createOrShow,
      PanelType.RestifyServerNotificationBotReadme
    );
  });

  it("signOutM365", async () => {
    const signOut = sandbox.stub(M365TokenInstance, "signout").resolves(true);
    const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(envTreeProviderInstance, "reloadEnvironments");

    await handlers.signOutM365(false);

    sandbox.assert.calledOnce(signOut);
  });

  it("signOutAzure", async () => {
    Object.setPrototypeOf(AzureAccountManager, sandbox.stub());
    const signOut = sandbox.stub(AzureAccountManager.getInstance(), "signout");
    const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

    await handlers.signOutAzure(false);

    sandbox.assert.calledOnce(signOut);
  });

  describe("decryptSecret", function () {
    this.afterEach(() => {
      sinon.restore();
    });
    it("successfully update secret", async () => {
      sinon.stub(globalVariables, "context").value({ extensionPath: "" });
      sinon.stub(handlers, "core").value(new MockCore());
      const sendTelemetryEvent = sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      const sendTelemetryErrorEvent = sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const decrypt = sinon.spy(handlers.core, "decrypt");
      const encrypt = sinon.spy(handlers.core, "encrypt");
      sinon.stub(vscode.commands, "executeCommand");
      const editBuilder = sinon.spy();
      sinon.stub(vscode.window, "activeTextEditor").value({
        edit: function (callback: (eb: any) => void) {
          callback({
            replace: editBuilder,
          });
        },
      });
      sinon.stub(extension, "VS_CODE_UI").value({
        inputText: () => Promise.resolve(ok({ type: "success", result: "inputValue" })),
      });
      const range = new vscode.Range(new vscode.Position(0, 10), new vscode.Position(0, 15));

      await handlers.decryptSecret("test", range);

      sinon.assert.calledOnce(decrypt);
      sinon.assert.calledOnce(encrypt);
      sinon.assert.calledOnce(editBuilder);
      sinon.assert.calledTwice(sendTelemetryEvent);
      sinon.assert.notCalled(sendTelemetryErrorEvent);
      sinon.restore();
    });

    it("failed to update due to corrupted secret", async () => {
      sinon.stub(globalVariables, "context").value({ extensionPath: "" });
      sinon.stub(handlers, "core").value(new MockCore());
      const sendTelemetryEvent = sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      const sendTelemetryErrorEvent = sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const decrypt = sinon.stub(handlers.core, "decrypt");
      decrypt.returns(Promise.resolve(err(new UserError("", "fake error", ""))));
      const encrypt = sinon.spy(handlers.core, "encrypt");
      sinon.stub(vscode.commands, "executeCommand");
      const editBuilder = sinon.spy();
      sinon.stub(vscode.window, "activeTextEditor").value({
        edit: function (callback: (eb: any) => void) {
          callback({
            replace: editBuilder,
          });
        },
      });
      const showMessage = sinon.stub(vscode.window, "showErrorMessage");
      const range = new vscode.Range(new vscode.Position(0, 10), new vscode.Position(0, 15));

      await handlers.decryptSecret("test", range);

      sinon.assert.calledOnce(decrypt);
      sinon.assert.notCalled(encrypt);
      sinon.assert.notCalled(editBuilder);
      sinon.assert.calledOnce(showMessage);
      sinon.assert.calledOnce(sendTelemetryEvent);
      sinon.assert.calledOnce(sendTelemetryErrorEvent);
      sinon.restore();
    });
  });

  describe("permission v3", function () {
    const sandbox = sinon.createSandbox();

    this.afterEach(() => {
      sandbox.restore();
    });

    it("happy path: grant permission", async () => {
      sandbox.stub(handlers, "core").value(new MockCore());
      sandbox.stub(extension, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: "grantPermission" })),
      });
      sandbox.stub(MockCore.prototype, "grantPermission").returns(
        Promise.resolve(
          ok({
            state: CollaborationState.OK,
            userInfo: {
              userObjectId: "fake-user-object-id",
              userPrincipalName: "fake-user-principle-name",
            },
            permissions: [
              {
                name: "name",
                type: "type",
                resourceId: "id",
                roles: ["Owner"],
              },
            ],
          })
        )
      );

      const result = await handlers.manageCollaboratorHandler("env");
      chai.expect(result.isOk()).equals(true);
    });

    it("happy path: list collaborator", async () => {
      sandbox.stub(handlers, "core").value(new MockCore());
      sandbox.stub(extension, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: "listCollaborator" })),
      });
      sandbox.stub(MockCore.prototype, "listCollaborator").returns(
        Promise.resolve(
          ok({
            state: CollaborationState.OK,
            collaborators: [
              {
                userPrincipalName: "userPrincipalName",
                userObjectId: "userObjectId",
                isAadOwner: true,
                teamsAppResourceId: "teamsAppResourceId",
              },
            ],
          })
        )
      );
      const vscodeLogProviderInstance = VsCodeLogProvider.getInstance();
      sandbox.stub(vscodeLogProviderInstance, "outputChannel").value({
        name: "name",
        append: (value: string) => {},
        appendLine: (value: string) => {},
        replace: (value: string) => {},
        clear: () => {},
        show: (...params: any[]) => {},
        hide: () => {},
        dispose: () => {},
      });

      const result = await handlers.manageCollaboratorHandler("env");
      chai.expect(result.isOk()).equals(true);
    });

    it("happy path: list collaborator throws error", async () => {
      sandbox.stub(handlers, "core").value(new MockCore());
      sandbox.stub(extension, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: "listCollaborator" })),
      });
      sandbox.stub(MockCore.prototype, "listCollaborator").throws(new Error("Error"));
      const vscodeLogProviderInstance = VsCodeLogProvider.getInstance();
      sandbox.stub(vscodeLogProviderInstance, "outputChannel").value({
        name: "name",
        append: (value: string) => {},
        appendLine: (value: string) => {},
        replace: (value: string) => {},
        clear: () => {},
        show: (...params: any[]) => {},
        hide: () => {},
        dispose: () => {},
      });

      const result = await handlers.manageCollaboratorHandler("env");
      chai.expect(result.isErr()).equals(true);
    });

    it("happy path: list collaborator throws login error", async () => {
      sandbox.stub(handlers, "core").value(new MockCore());
      sandbox.stub(extension, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: "listCollaborator" })),
      });
      const showErrorMessageStub = sandbox.stub(vscode.window, "showErrorMessage");
      sandbox
        .stub(MockCore.prototype, "listCollaborator")
        .throws(new Error("Cannot get user login information"));
      const vscodeLogProviderInstance = VsCodeLogProvider.getInstance();
      sandbox.stub(vscodeLogProviderInstance, "outputChannel").value({
        name: "name",
        append: (value: string) => {},
        appendLine: (value: string) => {},
        replace: (value: string) => {},
        clear: () => {},
        show: (...params: any[]) => {},
        hide: () => {},
        dispose: () => {},
      });

      const result = await handlers.manageCollaboratorHandler("env");
      chai.expect(result.isErr()).equals(true);
      chai.assert.isTrue(showErrorMessageStub.called);
    });

    it("User Cancel", async () => {
      sandbox.stub(handlers, "core").value(new MockCore());
      sandbox.stub(extension, "VS_CODE_UI").value({
        selectOption: () =>
          Promise.resolve(err(new UserError("source", "errorName", "errorMessage"))),
      });

      const result = await handlers.manageCollaboratorHandler();
      chai.expect(result.isErr()).equals(true);
    });
  });

  describe("checkUpgrade", function () {
    const sandbox = sinon.createSandbox();
    const mockCore = new MockCore();

    beforeEach(() => {
      sandbox.stub(handlers, "getSystemInputs").returns({} as Inputs);
      sandbox.stub(handlers, "core").value(mockCore);
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("calls phantomMigrationV3 with isNonmodalMessage when auto triggered", async () => {
      const phantomMigrationV3Stub = sandbox
        .stub(mockCore, "phantomMigrationV3")
        .resolves(ok(undefined));
      await handlers.checkUpgrade([extTelemetryEvents.TelemetryTriggerFrom.Auto]);
      chai.assert.isTrue(
        phantomMigrationV3Stub.calledOnceWith({
          locale: "en-us",
          platform: "vsc",
          projectPath: undefined,
          vscodeEnv: "local",
          isNonmodalMessage: true,
        } as Inputs)
      );
    });

    it("calls phantomMigrationV3 with skipUserConfirm trigger from sideBar and command palette", async () => {
      const phantomMigrationV3Stub = sandbox
        .stub(mockCore, "phantomMigrationV3")
        .resolves(ok(undefined));
      await handlers.checkUpgrade([extTelemetryEvents.TelemetryTriggerFrom.SideBar]);
      chai.assert.isTrue(
        phantomMigrationV3Stub.calledOnceWith({
          locale: "en-us",
          platform: "vsc",
          projectPath: undefined,
          vscodeEnv: "local",
          skipUserConfirm: true,
        } as Inputs)
      );
      await handlers.checkUpgrade([extTelemetryEvents.TelemetryTriggerFrom.CommandPalette]);
      chai.assert.isTrue(
        phantomMigrationV3Stub.calledWith({
          locale: "en-us",
          platform: "vsc",
          projectPath: undefined,
          vscodeEnv: "local",
          skipUserConfirm: true,
        } as Inputs)
      );
    });

    it("shows error message when phantomMigrationV3 fails", async () => {
      const error = new UserError(
        "test source",
        "test name",
        "test message",
        "test displayMessage"
      );
      error.helpLink = "test helpLink";
      const phantomMigrationV3Stub = sandbox
        .stub(mockCore, "phantomMigrationV3")
        .resolves(err(error));
      sandbox.stub(localizeUtils, "localize").returns("");
      const showErrorMessageStub = sandbox.stub(vscode.window, "showErrorMessage");
      sandbox.stub(vscode.commands, "executeCommand");

      await handlers.checkUpgrade([extTelemetryEvents.TelemetryTriggerFrom.SideBar]);
      chai.assert.isTrue(
        phantomMigrationV3Stub.calledOnceWith({
          locale: "en-us",
          platform: "vsc",
          projectPath: undefined,
          vscodeEnv: "local",
          skipUserConfirm: true,
        } as Inputs)
      );
      chai.assert.isTrue(showErrorMessageStub.calledOnce);
    });
  });

  it("downloadSample", async () => {
    const inputs: Inputs = {
      scratch: "no",
      platform: Platform.VSCode,
    };
    sandbox.stub(handlers, "core").value(new MockCore());
    const createProject = sandbox.spy(handlers.core, "createSampleProject");

    await handlers.downloadSample(inputs);

    inputs.stage = Stage.create;
    chai.assert.isTrue(createProject.calledOnceWith(inputs));
  });

  it("downloadSample - error", async () => {
    const inputs: Inputs = {
      scratch: "no",
      platform: Platform.VSCode,
    };
    sandbox.stub(handlers, "core").value(new MockCore());
    const showErrorMessageStub = sandbox.stub(vscode.window, "showErrorMessage");
    const createProject = sandbox
      .stub(handlers.core, "createSampleProject")
      .rejects(err(new Error("Cannot get user login information")));

    await handlers.downloadSample(inputs);

    inputs.stage = Stage.create;
    chai.assert.isTrue(createProject.calledOnceWith(inputs));
    chai.assert.isTrue(showErrorMessageStub.calledOnce);
  });

  it("downloadSample - LoginFailureError", async () => {
    const inputs: Inputs = {
      scratch: "no",
      platform: Platform.VSCode,
    };
    sandbox.stub(handlers, "core").value(new MockCore());
    const showErrorMessageStub = sandbox.stub(vscode.window, "showErrorMessage");
    const createProject = sandbox
      .stub(handlers.core, "createProject")
      .resolves(err(new SystemError("test", "test", "Cannot get user login information")));

    await handlers.downloadSample(inputs);
  });

  it("deployAadAppmanifest", async () => {
    sandbox.stub(handlers, "core").value(new MockCore());
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    const deployAadManifest = sandbox.spy(handlers.core, "deployAadManifest");
    await handlers.updateAadAppManifest([{ fsPath: "path/aad.dev.template" }]);
    sandbox.assert.calledOnce(deployAadManifest);
    deployAadManifest.restore();
  });

  it("showError", async () => {
    sandbox.stub(localizeUtils, "localize").returns("");
    const showErrorMessageStub = sandbox
      .stub(vscode.window, "showErrorMessage")
      .callsFake((title: string, button: any) => {
        return Promise.resolve(button);
      });
    const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(vscode.commands, "executeCommand");
    const error = new UserError("test source", "test name", "test message", "test displayMessage");
    error.helpLink = "test helpLink";

    await handlers.showError(error);

    chai.assert.isTrue(
      sendTelemetryEventStub.calledWith(extTelemetryEvents.TelemetryEvent.ClickGetHelp, {
        "error-code": "test source.test name",
        "error-message": "test displayMessage",
        "help-link": "test helpLink",
      })
    );
  });

  it("showError - similar issues", async () => {
    sandbox
      .stub(vscode.window, "showErrorMessage")
      .callsFake((title: string, button: unknown, ...items: vscode.MessageItem[]) => {
        return Promise.resolve(items[0]);
      });
    const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");
    const error = new SystemError("Core", "DecryptionError", "test");

    await handlers.showError(error);

    chai.assert.isTrue(sendTelemetryEventStub.called);
    chai.assert.isTrue(executeCommandStub.called);
  });

  [
    {
      type: "user error",
      buildError: () => {
        const error = new UserError(
          "test source",
          "test name",
          "test message",
          "test displayMessage"
        );
        error.helpLink = "test helpLink";
        error.recommendedOperation = debugConstants.RecommendedOperations.DebugInTestTool;

        return error;
      },
      buttonNum: 2,
    },
    {
      type: "system error",
      buildError: () => {
        const error = new SystemError(
          "test source",
          "test name",
          "test message",
          "test displayMessage"
        );
        error.recommendedOperation = debugConstants.RecommendedOperations.DebugInTestTool;
        return error;
      },
      buttonNum: 3,
    },
  ].forEach(({ type, buildError, buttonNum }) => {
    it(`showError - ${type} - recommend test tool`, async () => {
      sandbox.stub(localizeUtils, "localize").returns("");
      const showErrorMessageStub = sandbox
        .stub<any, any>(vscode.window, "showErrorMessage")
        .callsFake((title: string, button: any) => {
          return Promise.resolve(button);
        });
      sandbox.stub(debugCommonUtils, "isTestToolEnabledProject").returns(true);
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("path"));
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(vscode.commands, "executeCommand");
      const error = buildError();
      await handlers.showError(error);

      chai.assert.isTrue(
        sendTelemetryEventStub.calledWith(extTelemetryEvents.TelemetryEvent.MessageDebugInTestTool)
      );
      chai.assert.equal(showErrorMessageStub.firstCall.args.length, buttonNum + 1);
    });
  });

  describe("getDotnetPathHandler", async () => {
    afterEach(() => {
      sinon.restore();
    });
    it("dotnet is installed", async () => {
      sinon.stub(DepsManager.prototype, "getStatus").resolves([
        {
          name: ".NET Core SDK",
          type: DepsType.Dotnet,
          isInstalled: true,
          command: "",
          details: {
            isLinuxSupported: false,
            installVersion: "",
            supportedVersions: [],
            binFolders: ["dotnet-bin-folder/dotnet"],
          },
        },
      ]);

      const dotnetPath = await handlers.getDotnetPathHandler();
      chai.assert.equal(dotnetPath, `${path.delimiter}dotnet-bin-folder${path.delimiter}`);
    });

    it("dotnet is not installed", async () => {
      sinon.stub(DepsManager.prototype, "getStatus").resolves([
        {
          name: ".NET Core SDK",
          type: DepsType.Dotnet,
          isInstalled: false,
          command: "",
          details: {
            isLinuxSupported: false,
            installVersion: "",
            supportedVersions: [],
            binFolders: undefined,
          },
        },
      ]);

      const dotnetPath = await handlers.getDotnetPathHandler();
      chai.assert.equal(dotnetPath, `${path.delimiter}`);
    });

    it("failed to get dotnet path", async () => {
      sinon.stub(DepsManager.prototype, "getStatus").rejects(new Error("failed to get status"));
      const dotnetPath = await handlers.getDotnetPathHandler();
      chai.assert.equal(dotnetPath, `${path.delimiter}`);
    });
  });

  describe("scaffoldFromDeveloperPortalHandler", async () => {
    beforeEach(() => {
      sinon.stub(ExtTelemetry, "sendTelemetryEvent").resolves();
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent").resolves();
      sinon.stub(globalVariables, "checkIsSPFx").returns(false);
    });
    afterEach(() => {
      sinon.restore();
    });
    it("missing args", async () => {
      const progressHandler = new ProgressHandler("title", 1);
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      const createProgressBar = sinon
        .stub(extension.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);

      const res = await handlers.scaffoldFromDeveloperPortalHandler();

      chai.assert.equal(res.isOk(), true);
      chai.assert.equal(createProgressBar.notCalled, true);
    });

    it("incorrect number of args", async () => {
      const progressHandler = new ProgressHandler("title", 1);
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      const createProgressBar = sinon
        .stub(extension.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);

      const res = await handlers.scaffoldFromDeveloperPortalHandler([]);

      chai.assert.equal(res.isOk(), true);
      chai.assert.equal(createProgressBar.notCalled, true);
    });

    it("general error when signing in M365", async () => {
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      const progressHandler = new ProgressHandler("title", 1);
      const startProgress = sinon.stub(progressHandler, "start").resolves();
      const endProgress = sinon.stub(progressHandler, "end").resolves();
      sinon.stub(M365TokenInstance, "signInWhenInitiatedFromTdp").throws("error1");
      const createProgressBar = sinon
        .stub(extension.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);
      const showErrorMessage = sinon.stub(vscode.window, "showErrorMessage");

      const res = await handlers.scaffoldFromDeveloperPortalHandler(["appId"]);
      chai.assert.isTrue(res.isErr());
      chai.assert.isTrue(createProgressBar.calledOnce);
      chai.assert.isTrue(startProgress.calledOnce);
      chai.assert.isTrue(endProgress.calledOnceWithExactly(false));
      chai.assert.isTrue(showErrorMessage.calledOnce);
      if (res.isErr()) {
        chai.assert.isTrue(res.error instanceof UnhandledError);
      }
    });

    it("error when signing M365", async () => {
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      const progressHandler = new ProgressHandler("title", 1);
      const startProgress = sinon.stub(progressHandler, "start").resolves();
      const endProgress = sinon.stub(progressHandler, "end").resolves();
      sinon
        .stub(M365TokenInstance, "signInWhenInitiatedFromTdp")
        .resolves(err(new UserError("source", "name", "message", "displayMessage")));
      const createProgressBar = sinon
        .stub(extension.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);
      const showErrorMessage = sinon.stub(vscode.window, "showErrorMessage");

      const res = await handlers.scaffoldFromDeveloperPortalHandler(["appId"]);

      chai.assert.equal(res.isErr(), true);
      chai.assert.equal(createProgressBar.calledOnce, true);
      chai.assert.equal(startProgress.calledOnce, true);
      chai.assert.equal(endProgress.calledOnceWithExactly(false), true);
      chai.assert.equal(showErrorMessage.calledOnce, true);
    });

    it("error when signing in M365 but missing display message", async () => {
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      const progressHandler = new ProgressHandler("title", 1);
      const startProgress = sinon.stub(progressHandler, "start").resolves();
      const endProgress = sinon.stub(progressHandler, "end").resolves();
      sinon
        .stub(M365TokenInstance, "signInWhenInitiatedFromTdp")
        .resolves(err(new UserError("source", "name", "", "")));
      const createProgressBar = sinon
        .stub(extension.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);
      const showErrorMessage = sinon.stub(vscode.window, "showErrorMessage");

      const res = await handlers.scaffoldFromDeveloperPortalHandler(["appId"]);

      chai.assert.equal(res.isErr(), true);
      chai.assert.equal(createProgressBar.calledOnce, true);
      chai.assert.equal(startProgress.calledOnce, true);
      chai.assert.equal(endProgress.calledOnceWithExactly(false), true);
      chai.assert.equal(showErrorMessage.calledOnce, true);
    });

    it("failed to get teams app", async () => {
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      const progressHandler = new ProgressHandler("title", 1);
      const startProgress = sinon.stub(progressHandler, "start").resolves();
      const endProgress = sinon.stub(progressHandler, "end").resolves();
      sinon.stub(M365TokenInstance, "signInWhenInitiatedFromTdp").resolves(ok("token"));
      sinon
        .stub(M365TokenInstance, "getAccessToken")
        .resolves(err(new SystemError("source", "name", "", "")));
      const createProgressBar = sinon
        .stub(extension.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(vscode.commands, "executeCommand");
      sinon.stub(globalState, "globalStateUpdate");
      const getApp = sinon.stub(AppStudioClient, "getApp").throws("error");

      const res = await handlers.scaffoldFromDeveloperPortalHandler(["appId"]);

      chai.assert.isTrue(res.isErr());
      chai.assert.isTrue(getApp.calledOnce);
      chai.assert.isTrue(createProgressBar.calledOnce);
      chai.assert.isTrue(startProgress.calledOnce);
      chai.assert.isTrue(endProgress.calledOnceWithExactly(true));
    });

    it("happy path", async () => {
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      const progressHandler = new ProgressHandler("title", 1);
      const startProgress = sinon.stub(progressHandler, "start").resolves();
      const endProgress = sinon.stub(progressHandler, "end").resolves();
      sinon.stub(M365TokenInstance, "signInWhenInitiatedFromTdp").resolves(ok("token"));
      sinon.stub(M365TokenInstance, "getAccessToken").resolves(ok("authSvcToken"));
      sinon.stub(commonTools, "setRegion").resolves();
      const createProgressBar = sinon
        .stub(extension.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);
      sinon.stub(handlers, "core").value(new MockCore());
      const createProject = sinon.spy(handlers.core, "createProject");
      sinon.stub(vscode.commands, "executeCommand");
      sinon.stub(globalState, "globalStateUpdate");
      const appDefinition: AppDefinition = {
        teamsAppId: "mock-id",
      };
      sinon.stub(AppStudioClient, "getApp").resolves(appDefinition);

      const res = await handlers.scaffoldFromDeveloperPortalHandler(["appId", "testuser"]);

      chai.assert.equal(createProject.args[0][0].teamsAppFromTdp.teamsAppId, "mock-id");
      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(createProgressBar.calledOnce);
      chai.assert.isTrue(startProgress.calledOnce);
      chai.assert.isTrue(endProgress.calledOnceWithExactly(true));
    });
  });

  describe("publishInDeveloperPortalHandler", async () => {
    beforeEach(() => {
      sinon.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("path"));
    });

    afterEach(() => {
      sinon.restore();
    });

    it("publish in developer portal - success", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      sinon
        .stub(extension.VS_CODE_UI, "selectFile")
        .resolves(ok({ type: "success", result: "test.zip" }));
      const publish = sinon.spy(handlers.core, "publishInDeveloperPortal");
      sinon
        .stub(extension.VS_CODE_UI, "selectOption")
        .resolves(ok({ type: "success", result: "test.zip" }));
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(vscode.commands, "executeCommand");
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readdir").resolves(["test.zip", "test.json"] as any);

      const res = await handlers.publishInDeveloperPortalHandler();
      if (res.isErr()) {
        console.log(res.error);
      }
      chai.assert.isTrue(publish.calledOnce);
      chai.assert.isTrue(res.isOk());
    });

    it("publish in developer portal - cancelled", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      sinon
        .stub(extension.VS_CODE_UI, "selectFile")
        .resolves(ok({ type: "success", result: "test2.zip" }));
      const publish = sinon.spy(handlers.core, "publishInDeveloperPortal");
      sinon.stub(extension.VS_CODE_UI, "selectOption").resolves(err(new UserCancelError("VSC")));
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(vscode.commands, "executeCommand");
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readdir").resolves(["test.zip", "test.json"] as any);

      const res = await handlers.publishInDeveloperPortalHandler();
      if (res.isErr()) {
        console.log(res.error);
      }
      chai.assert.isTrue(publish.notCalled);
      chai.assert.isTrue(res.isOk());
    });

    it("select file error", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      sinon.stub(extension.VS_CODE_UI, "selectFile").resolves(err(new UserCancelError("VSC")));
      const publish = sinon.spy(handlers.core, "publishInDeveloperPortal");
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(vscode.commands, "executeCommand");
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readdir").resolves(["test.zip", "test.json"] as any);

      const res = await handlers.publishInDeveloperPortalHandler();
      chai.assert.isTrue(res.isOk());
      chai.assert.isFalse(publish.calledOnce);
    });
  });

  describe("openAppManagement", async () => {
    afterEach(() => {
      sinon.restore();
    });

    it("open link with loginHint", async () => {
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(M365TokenInstance, "getStatus").resolves(
        ok({
          status: signedIn,
          token: undefined,
          accountInfo: { upn: "test" },
        })
      );
      const openUrl = sinon.stub(extension.VS_CODE_UI, "openUrl").resolves(ok(true));

      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");

      const res = await handlers.openAppManagement();

      chai.assert.isTrue(openUrl.calledOnce);
      chai.assert.isTrue(res.isOk());
      chai.assert.equal(openUrl.args[0][0], `${DeveloperPortalHomeLink}?login_hint=test`);
    });

    it("open link without loginHint", async () => {
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      sinon.stub(M365TokenInstance, "getStatus").resolves(
        ok({
          status: signedOut,
          token: undefined,
          accountInfo: { upn: "test" },
        })
      );
      const openUrl = sinon.stub(extension.VS_CODE_UI, "openUrl").resolves(ok(true));

      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");

      const res = await handlers.openAppManagement();

      chai.assert.isTrue(openUrl.calledOnce);
      chai.assert.isTrue(res.isOk());
      chai.assert.equal(openUrl.args[0][0], DeveloperPortalHomeLink);
    });
  });

  describe("installAppInTeams", () => {
    afterEach(() => {
      sinon.restore();
    });

    it("happy path", async () => {
      sinon.stub(debugCommonUtils, "triggerV3Migration").resolves();
      const result = await handlers.installAppInTeams();
      chai.assert.equal(result, undefined);
    });

    it("migration error", async () => {
      sinon.stub(debugCommonUtils, "triggerV3Migration").throws(err({ foo: "bar" } as any));
      sinon.stub(handlers, "showError").resolves();
      const result = await handlers.installAppInTeams();
      chai.assert.equal(result, "1");
    });
  });

  describe("callBackFunctions", () => {
    it("checkCopilotCallback()", async () => {
      sinon.stub(localizeUtils, "localize").returns("");
      let showMessageCalledCount = 0;
      sinon.stub(extension, "VS_CODE_UI").value({
        showMessage: async () => {
          showMessageCalledCount += 1;
          return Promise.resolve(ok("Enroll"));
        },
      });

      handlers.checkCopilotCallback();

      chai.expect(showMessageCalledCount).to.be.equal(1);
      sinon.restore();
    });

    it("checkSideloadingCallback()", async () => {
      sinon.stub(localizeUtils, "localize").returns("");
      let showMessageCalledCount = 0;
      sinon.stub(extension, "VS_CODE_UI").value({
        showMessage: async () => {
          showMessageCalledCount += 1;
          return Promise.resolve(ok("Learn More"));
        },
      });
      const createOrShow = sinon.stub(WebviewPanel, "createOrShow");

      handlers.checkSideloadingCallback();

      chai.expect(showMessageCalledCount).to.be.equal(1);
      sinon.assert.calledOnceWithExactly(createOrShow, PanelType.AccountHelp);
      sinon.restore();
    });

    it("signinAzureCallback", async () => {
      sinon.stub(AzureAccountManager.prototype, "getAccountInfo").returns({});
      const getIdentityCredentialStub = sinon.stub(
        AzureAccountManager.prototype,
        "getIdentityCredentialAsync"
      );

      await handlers.signinAzureCallback([{}, { status: 0 }]);

      chai.assert.isTrue(getIdentityCredentialStub.calledOnce);
      sinon.restore();
    });
  });

  describe("validateAzureDependenciesHandler", () => {
    afterEach(() => {
      sinon.restore();
    });

    it("happy path", async () => {
      sinon.stub(debugCommonUtils, "triggerV3Migration").resolves();
      const result = await handlers.validateAzureDependenciesHandler();
      chai.assert.equal(result, undefined);
    });

    it("migration error", async () => {
      sinon.stub(debugCommonUtils, "triggerV3Migration").throws(err({ foo: "bar" } as any));
      sinon.stub(handlers, "showError").resolves();
      const result = await handlers.validateAzureDependenciesHandler();
      chai.assert.equal(result, "1");
    });
  });

  describe("validateLocalPrerequisitesHandler", () => {
    afterEach(() => {
      sinon.restore();
    });

    it("happy path", async () => {
      sinon.stub(debugCommonUtils, "triggerV3Migration").resolves();
      const result = await handlers.validateLocalPrerequisitesHandler();
      chai.assert.equal(result, undefined);
    });

    it("migration error", async () => {
      sinon.stub(debugCommonUtils, "triggerV3Migration").throws(err({ foo: "bar" } as any));
      sinon.stub(handlers, "showError").resolves();
      const result = await handlers.validateLocalPrerequisitesHandler();
      chai.assert.equal(result, "1");
    });
  });

  describe("backendExtensionsInstallHandler", () => {
    it("happy path", async () => {
      sinon.stub(debugCommonUtils, "triggerV3Migration").resolves();
      const result = await handlers.backendExtensionsInstallHandler();
      chai.assert.equal(result, undefined);
      sinon.restore();
    });

    it("migration error", async () => {
      sinon.stub(debugCommonUtils, "triggerV3Migration").throws(err({ foo: "bar" } as any));
      sinon.stub(handlers, "showError").resolves();
      const result = await handlers.backendExtensionsInstallHandler();
      chai.assert.equal(result, "1");
      sinon.restore();
    });
  });

  describe("preDebugCheckHandler", () => {
    it("happy path", async () => {
      sinon.stub(debugCommonUtils, "triggerV3Migration").resolves();
      const result = await handlers.preDebugCheckHandler();
      chai.assert.equal(result, undefined);
      sinon.restore();
    });

    it("happy path", async () => {
      sinon.stub(debugCommonUtils, "triggerV3Migration").throws(err({ foo: "bar" } as any));
      sinon.stub(handlers, "showError").resolves();
      const result = await handlers.preDebugCheckHandler();
      chai.assert.equal(result, "1");
      sinon.restore();
    });
  });

  describe("migrateTeamsTabAppHandler", () => {
    afterEach(() => {
      sinon.restore();
    });

    it("happy path", async () => {
      sinon.stub(ExtTelemetry, "sendTelemetryEvent").returns();
      sinon.stub(localizeUtils, "localize").callsFake((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      sinon.stub(extension, "VS_CODE_UI").value({
        showMessage: () => Promise.resolve(ok("teamstoolkit.migrateTeamsTabApp.upgrade")),
        selectFolder: () => Promise.resolve(ok({ type: "success", result: "test" })),
        createProgressBar: () => progressHandler,
      });
      sinon.stub(VsCodeLogInstance, "info").returns();
      sinon.stub(TeamsAppMigrationHandler.prototype, "updatePackageJson").resolves(ok(true));
      sinon.stub(TeamsAppMigrationHandler.prototype, "updateCodes").resolves(ok([]));

      const result = await handlers.migrateTeamsTabAppHandler();

      chai.assert.deepEqual(result, ok(null));
    });

    it("happy path: failed files", async () => {
      sinon.stub(ExtTelemetry, "sendTelemetryEvent").returns();
      sinon.stub(localizeUtils, "localize").callsFake((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      sinon.stub(extension, "VS_CODE_UI").value({
        showMessage: () => Promise.resolve(ok("teamstoolkit.migrateTeamsTabApp.upgrade")),
        selectFolder: () => Promise.resolve(ok({ type: "success", result: "test" })),
        createProgressBar: () => progressHandler,
      });
      sinon.stub(VsCodeLogInstance, "info").returns();
      const warningStub = sinon.stub(VsCodeLogInstance, "warning");
      sinon.stub(TeamsAppMigrationHandler.prototype, "updatePackageJson").resolves(ok(true));
      sinon
        .stub(TeamsAppMigrationHandler.prototype, "updateCodes")
        .resolves(ok(["test1", "test2"]));

      const result = await handlers.migrateTeamsTabAppHandler();

      chai.assert.deepEqual(result, ok(null));
      chai.expect(warningStub.calledOnce).to.be.true;
    });

    it("error", async () => {
      sinon.stub(ExtTelemetry, "sendTelemetryEvent").returns();
      const sendTelemetryErrorEventStub = sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(localizeUtils, "localize").callsFake((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      sinon.stub(extension, "VS_CODE_UI").value({
        showMessage: () => Promise.resolve(ok("teamstoolkit.migrateTeamsTabApp.upgrade")),
        selectFolder: () => Promise.resolve(ok({ type: "success", result: "test" })),
        createProgressBar: () => progressHandler,
      });
      sinon.stub(VsCodeLogInstance, "info").returns();
      sinon.stub(TeamsAppMigrationHandler.prototype, "updatePackageJson").resolves(ok(true));
      sinon
        .stub(TeamsAppMigrationHandler.prototype, "updateCodes")
        .resolves(err({ foo: "bar" } as any));

      const result = await handlers.migrateTeamsTabAppHandler();

      chai.assert.isTrue(result.isErr());
      chai.expect(sendTelemetryErrorEventStub.calledOnce).to.be.true;
    });

    it("user cancel", async () => {
      sinon.stub(ExtTelemetry, "sendTelemetryEvent").returns();
      sinon.stub(localizeUtils, "localize").callsFake((key: string) => key);
      const sendTelemetryErrorEventStub = sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(extension, "VS_CODE_UI").value({
        showMessage: () => Promise.resolve(ok("teamstoolkit.migrateTeamsTabApp.upgrade")),
        selectFolder: () => Promise.resolve(ok({ type: "skip" })),
      });

      const result = await handlers.migrateTeamsTabAppHandler();

      chai.assert.deepEqual(result, ok(null));
      chai.expect(sendTelemetryErrorEventStub.calledOnce).to.be.true;
    });

    it("user cancel: skip folder selection", async () => {
      sinon.stub(ExtTelemetry, "sendTelemetryEvent").returns();
      sinon.stub(localizeUtils, "localize").callsFake((key: string) => key);
      const sendTelemetryErrorEventStub = sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(extension, "VS_CODE_UI").value({
        showMessage: () => Promise.resolve(ok("cancel")),
      });

      const result = await handlers.migrateTeamsTabAppHandler();

      chai.assert.deepEqual(result, ok(null));
      chai.expect(sendTelemetryErrorEventStub.calledOnce).to.be.true;
    });

    it("no change in package.json", async () => {
      sinon.stub(ExtTelemetry, "sendTelemetryEvent").returns();
      sinon.stub(localizeUtils, "localize").callsFake((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      sinon.stub(extension, "VS_CODE_UI").value({
        showMessage: () => Promise.resolve(ok("teamstoolkit.migrateTeamsTabApp.upgrade")),
        selectFolder: () => Promise.resolve(ok({ type: "success", result: "test" })),
        createProgressBar: () => progressHandler,
      });
      sinon.stub(VsCodeLogInstance, "info").returns();
      sinon.stub(VsCodeLogInstance, "warning").returns();
      sinon.stub(TeamsAppMigrationHandler.prototype, "updatePackageJson").resolves(ok(false));

      const result = await handlers.migrateTeamsTabAppHandler();

      chai.assert.deepEqual(result, ok(null));
    });
  });

  describe("migrateTeamsManifestHandler", () => {
    afterEach(() => {
      sinon.restore();
    });

    it("happy path", async () => {
      sinon.stub(ExtTelemetry, "sendTelemetryEvent").returns();
      sinon.stub(localizeUtils, "localize").callsFake((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      sinon.stub(extension, "VS_CODE_UI").value({
        showMessage: () => Promise.resolve(ok("teamstoolkit.migrateTeamsManifest.upgrade")),
        selectFile: () => Promise.resolve(ok({ type: "success", result: "test" })),
        createProgressBar: () => progressHandler,
      });
      sinon.stub(VsCodeLogInstance, "info").returns();
      sinon.stub(TeamsAppMigrationHandler.prototype, "updateManifest").resolves(ok(null));

      const result = await handlers.migrateTeamsManifestHandler();

      chai.assert.deepEqual(result, ok(null));
    });

    it("user cancel: skip file selection", async () => {
      sinon.stub(ExtTelemetry, "sendTelemetryEvent").returns();
      const sendTelemetryErrorEventStub = sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(localizeUtils, "localize").callsFake((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      sinon.stub(extension, "VS_CODE_UI").value({
        showMessage: () => Promise.resolve(ok("teamstoolkit.migrateTeamsManifest.upgrade")),
        selectFile: () => Promise.resolve(ok({ type: "skip" })),
        createProgressBar: () => progressHandler,
      });
      sinon.stub(VsCodeLogInstance, "info").returns();
      sinon.stub(TeamsAppMigrationHandler.prototype, "updateManifest").resolves(ok(null));

      const result = await handlers.migrateTeamsManifestHandler();

      chai.assert.deepEqual(result, ok(null));
      chai.expect(sendTelemetryErrorEventStub.calledOnce).to.be.true;
    });

    it("error", async () => {
      sinon.stub(ExtTelemetry, "sendTelemetryEvent").returns();
      sinon.stub(localizeUtils, "localize").callsFake((key: string) => key);
      const sendTelemetryErrorEventStub = sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const progressHandler = new ProgressHandler("title", 1);
      sinon.stub(extension, "VS_CODE_UI").value({
        showMessage: () => Promise.resolve(ok("teamstoolkit.migrateTeamsManifest.upgrade")),
        selectFile: () => Promise.resolve(ok({ type: "success", result: "test" })),
        createProgressBar: () => progressHandler,
      });
      sinon.stub(VsCodeLogInstance, "info").returns();
      sinon
        .stub(TeamsAppMigrationHandler.prototype, "updateManifest")
        .resolves(err(new UserError("source", "name", "")));
      sinon.stub(handlers, "showError").callsFake(async () => {});

      const result = await handlers.migrateTeamsManifestHandler();

      chai.assert.isTrue(result.isErr());
      chai.expect(sendTelemetryErrorEventStub.calledOnce).to.be.true;
    });
  });

  describe("openDocumentHandler", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("opens upgrade guide when clicked from sidebar", async () => {
      const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      const openUrl = sandbox.stub(extension.VS_CODE_UI, "openUrl").resolves(ok(true));

      await handlers.openDocumentHandler([
        extTelemetryEvents.TelemetryTriggerFrom.SideBar,
        "learnmore",
      ]);

      chai.assert.isTrue(sendTelemetryStub.calledOnceWith("documentation"));
      chai.assert.isTrue(openUrl.calledOnceWith("https://aka.ms/teams-toolkit-5.0-upgrade"));
    });
  });

  it("refreshSPFxTreeOnFileChanged", () => {
    const initGlobalVariables = sandbox.stub(globalVariables, "initializeGlobalVariables");
    const updateTreeViewsOnSPFxChanged = sandbox
      .stub(TreeViewManagerInstance, "updateTreeViewsOnSPFxChanged")
      .resolves();

    handlers.refreshSPFxTreeOnFileChanged();

    chai.expect(initGlobalVariables.calledOnce).to.be.true;
    chai.expect(updateTreeViewsOnSPFxChanged.calledOnce).to.be.true;
  });

  describe("getPathDelimiterHandler", () => {
    it("happy path", async () => {
      const actualPath = await handlers.getPathDelimiterHandler();
      chai.assert.equal(actualPath, path.delimiter);
    });
  });

  describe("others", function () {
    const sandbox = sinon.createSandbox();
    afterEach(() => {
      sandbox.restore();
    });

    it("cmpAccountsHandler", async () => {
      const AzureSignOutStub = sandbox.stub(AzureAccountManager.prototype, "signout");
      const M365SignOutStub = sandbox.stub(M365TokenInstance, "signout");
      sandbox
        .stub(M365TokenInstance, "getStatus")
        .resolves(ok({ status: "SignedIn", accountInfo: { upn: "test.email.com" } }));
      sandbox
        .stub(AzureAccountManager.prototype, "getStatus")
        .resolves({ status: "SignedIn", accountInfo: { upn: "test.email.com" } });
      const stubQuickPick = {
        items: [],
        onDidChangeSelection: () => {
          return {
            dispose: () => {},
          };
        },
        onDidHide: () => {
          return {
            dispose: () => {},
          };
        },
        show: () => {},
        onDidAccept: () => {},
      };
      sandbox.stub(vscode.window, "createQuickPick").returns(stubQuickPick as any);
      sandbox.stub(extension.VS_CODE_UI, "selectOption").resolves(ok({ result: "unknown" } as any));

      await handlers.cmpAccountsHandler([]);

      for (const i of stubQuickPick.items) {
        await (i as any).function();
      }

      chai.assert.isTrue(AzureSignOutStub.calledOnce);
      chai.assert.isTrue(M365SignOutStub.calledOnce);
    });

    it("updatePreviewManifest", async () => {
      sandbox.stub(handlers, "core").value(new MockCore());
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const openTextDocumentStub = sandbox
        .stub(vscode.workspace, "openTextDocument")
        .returns(Promise.resolve("" as any));

      await handlers.updatePreviewManifest([]);

      chai.assert.isTrue(openTextDocumentStub.calledOnce);
    });
  });
});

describe("openPreviewAadFile", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("manifest file not exists", async () => {
    const core = new MockCore();
    sandbox.stub(handlers, "core").value(core);
    sandbox.stub(projectSettingsHelper, "isValidProject").returns(true);
    sandbox.stub(fs, "existsSync").returns(false);
    sandbox.stub(environmentManager, "listAllEnvConfigs").resolves(ok(["dev"]));
    sandbox.stub(extension.VS_CODE_UI, "selectOption").resolves(
      ok({
        type: "success",
        result: "dev",
      })
    );
    sandbox.stub(handlers, "askTargetEnvironment").resolves(ok("dev"));
    sandbox.stub(handlers, "showError").callsFake(async () => {});
    sandbox.stub(handlers.core, "buildAadManifest").resolves(ok(undefined));
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent").resolves();
    const res = await handlers.openPreviewAadFile([]);
    chai.assert.isTrue(res.isErr());
  });

  it("happy path", async () => {
    const core = new MockCore();
    sandbox.stub(handlers, "core").value(core);
    sandbox.stub(projectSettingsHelper, "isValidProject").returns(true);
    sandbox.stub(fs, "existsSync").returns(true);
    sandbox.stub(environmentManager, "listAllEnvConfigs").resolves(ok(["dev"]));
    sandbox.stub(extension.VS_CODE_UI, "selectOption").resolves(
      ok({
        type: "success",
        result: "dev",
      })
    );
    sandbox.stub(handlers, "askTargetEnvironment").resolves(ok("dev"));
    sandbox.stub(handlers, "showError").callsFake(async () => {});
    sandbox.stub(handlers.core, "buildAadManifest").resolves(ok(undefined));
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent").resolves();
    sandbox.stub(vscode.workspace, "openTextDocument").resolves();
    sandbox.stub(vscode.window, "showTextDocument").resolves();

    const res = await handlers.openPreviewAadFile([]);
    chai.assert.isTrue(res.isOk());
  });
});

describe("editAadManifestTemplate", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("happy path", async () => {
    const workspacePath = "/test/workspace/path";
    const workspaceUri = vscode.Uri.file(workspacePath);
    sinon.stub(globalVariables, "workspaceUri").value(workspaceUri);

    const openTextDocumentStub = sandbox
      .stub(vscode.workspace, "openTextDocument")
      .resolves({} as any);
    const showTextDocumentStub = sandbox.stub(vscode.window, "showTextDocument");

    await handlers.editAadManifestTemplate([null, "testTrigger"]);

    sandbox.assert.calledOnceWithExactly(
      openTextDocumentStub as any,
      `${workspaceUri.fsPath}/aad.manifest.json`
    );
  });

  it("happy path: no parameter", async () => {
    const workspacePath = "/test/workspace/path";
    const workspaceUri = vscode.Uri.file(workspacePath);
    sinon.stub(globalVariables, "workspaceUri").value(workspaceUri);

    const openTextDocumentStub = sandbox
      .stub(vscode.workspace, "openTextDocument")
      .resolves({} as any);
    const showTextDocumentStub = sandbox.stub(vscode.window, "showTextDocument");

    await handlers.editAadManifestTemplate([]);

    chai.assert.isTrue(showTextDocumentStub.callCount === 0);
  });

  it("happy path: workspaceUri is undefined", async () => {
    const workspaceUri = undefined;
    sinon.stub(globalVariables, "workspaceUri").value(undefined);

    const openTextDocumentStub = sandbox
      .stub(vscode.workspace, "openTextDocument")
      .resolves({} as any);
    const showTextDocumentStub = sandbox.stub(vscode.window, "showTextDocument");

    await handlers.editAadManifestTemplate([null, "testTrigger"]);

    sandbox.assert.calledOnceWithExactly(
      openTextDocumentStub as any,
      `${workspaceUri}/aad.manifest.json`
    );
  });
});

describe("autoOpenProjectHandler", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("opens walk through", async () => {
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "fx-extension.openWalkThrough") {
        return true;
      } else {
        return false;
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const executeCommandFunc = sandbox.stub(vscode.commands, "executeCommand");

    await handlers.autoOpenProjectHandler();

    chai.assert.isTrue(sendTelemetryStub.calledOnce);
    chai.assert.isTrue(executeCommandFunc.calledOnce);
  });

  it("opens README", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    sandbox.stub(globalVariables, "isTeamsFxProject").resolves(false);
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(undefined);
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "fx-extension.openReadMe") {
        return vscode.Uri.file("test").fsPath;
      } else {
        return "";
      }
    });
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok({} as any));
    sandbox.stub(ManifestUtil, "parseCommonProperties").resolves({ isCopilotPlugin: false });
    sandbox.stub(globalState, "globalStateUpdate");
    const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

    await handlers.autoOpenProjectHandler();

    chai.assert.isTrue(sendTelemetryStub.calledOnce);
  });

  it("opens sample README", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    sandbox.stub(globalVariables, "isTeamsFxProject").resolves(false);
    const showMessageStub = sandbox.stub(vscode.window, "showInformationMessage");
    sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
    sandbox.stub(vscode.workspace, "openTextDocument");
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "fx-extension.openSampleReadMe") {
        return true;
      } else {
        return "";
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    await handlers.autoOpenProjectHandler();

    chai.assert.isTrue(executeCommandStub.calledOnce);
  });

  it("opens README and show warnings successfully", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    sandbox.stub(globalVariables, "isTeamsFxProject").resolves(false);
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(undefined);
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "fx-extension.openReadMe") {
        return vscode.Uri.file("test").fsPath;
      } else if (key === GlobalKey.CreateWarnings) {
        return JSON.stringify([{ type: "type", content: "content" }]);
      } else {
        return "";
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(
      ok({
        name: { short: "short", full: "full" },
        description: { short: "short", full: "" },
        composeExtensions: [{ commands: [{ id: "command1" }] }],
      } as any)
    );
    const parseRes = {
      id: "",
      version: "",
      capabilities: [""],
      manifestVersion: "",
      isApiME: true,
      isSPFx: false,
      isApiBasedMe: true,
    };
    const parseManifestStub = sandbox.stub(ManifestUtil, "parseCommonProperties").returns(parseRes);
    VsCodeLogInstance.outputChannel = {
      show: () => {},
      info: () => {},
    } as unknown as vscode.OutputChannel;
    const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

    await handlers.autoOpenProjectHandler();

    chai.assert.isTrue(sendTelemetryStub.called);
    chai.assert.isTrue(sendTelemetryStub.calledTwice);
    chai.assert.isTrue(parseManifestStub.called);
  });

  it("skip show warnings if parsing error", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    sandbox.stub(globalVariables, "isTeamsFxProject").resolves(false);
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(undefined);
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "fx-extension.openReadMe") {
        return vscode.Uri.file("test").fsPath;
      } else if (key === GlobalKey.CreateWarnings) {
        return "string";
      } else {
        return "";
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const sendErrorTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");

    await handlers.autoOpenProjectHandler();

    chai.assert.isTrue(sendErrorTelemetryStub.called);
  });

  it("skip show warnings if cannot get manifest", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    sandbox.stub(globalVariables, "isTeamsFxProject").resolves(false);
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(undefined);
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "fx-extension.openReadMe") {
        return vscode.Uri.file("test").fsPath;
      } else if (key === GlobalKey.CreateWarnings) {
        return "string";
      } else {
        return "";
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox
      .stub(manifestUtils, "_readAppManifest")
      .resolves(err(new UserError("source", "name", "", "")));

    const sendErrorTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");

    await handlers.autoOpenProjectHandler();

    chai.assert.isTrue(sendErrorTelemetryStub.called);
  });

  it("auto install dependency", async () => {
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "teamsToolkit:autoInstallDependency") {
        return true;
      } else {
        return false;
      }
    });
    const globalStateStub = sandbox.stub(globalState, "globalStateUpdate");
    const runCommandStub = sandbox.stub(extension.VS_CODE_UI, "runCommand");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

    await handlers.autoOpenProjectHandler();

    chai.assert.isTrue(globalStateStub.calledWith("teamsToolkit:autoInstallDependency", false));
    chai.assert.isTrue(runCommandStub.calledOnce);
  });

  it("openFolderHandler()", async () => {
    const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

    const result = await handlers.openFolderHandler();

    chai.assert.isTrue(sendTelemetryStub.called);
    chai.assert.isTrue(result.isOk());
  });

  it("runUserTask() - error", async () => {
    const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    sandbox.stub(handlers, "core").value(undefined);
    sandbox.stub(commonUtils, "getTeamsAppTelemetryInfoByEnv");
    sandbox.stub(VsCodeLogInstance, "error");

    const result = await handlers.runUserTask({ namespace: "test", method: "test" }, "test", true);

    chai.assert.isTrue(sendTelemetryStub.called);
    chai.assert.isTrue(result.isErr());
  });

  it("validateGetStartedPrerequisitesHandler() - error", async () => {
    const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox
      .stub(localPrerequisites, "checkPrerequisitesForGetStarted")
      .resolves(err(new SystemError("test", "test", "test")));

    const result = await handlers.validateGetStartedPrerequisitesHandler();

    chai.assert.isTrue(sendTelemetryStub.called);
    chai.assert.equal(result, "1");
  });

  it("registerAccountMenuCommands() - signedinM365", async () => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox
      .stub(vscode.commands, "registerCommand")
      .callsFake((command: string, callback: (...args: any[]) => any) => {
        callback({ contextValue: "signedinM365" }).then(() => {});
        return {
          dispose: () => {},
        };
      });
    sandbox.stub(vscode.extensions, "getExtension");
    const signoutStub = sandbox.stub(M365TokenInstance, "signout");

    await handlers.registerAccountMenuCommands({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);

    chai.assert.isTrue(signoutStub.called);
  });

  it("registerAccountMenuCommands() - signedinAzure", async () => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox
      .stub(vscode.commands, "registerCommand")
      .callsFake((command: string, callback: (...args: any[]) => any) => {
        callback({ contextValue: "signedinAzure" }).then(() => {});
        return {
          dispose: () => {},
        };
      });
    sandbox.stub(vscode.extensions, "getExtension");
    const signoutStub = sandbox.stub(AzureAccountManager.prototype, "signout");

    await handlers.registerAccountMenuCommands({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);

    chai.assert.isTrue(signoutStub.called);
  });

  it("registerAccountMenuCommands() - error", async () => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox
      .stub(vscode.commands, "registerCommand")
      .callsFake((command: string, callback: (...args: any[]) => any) => {
        callback({ contextValue: "signedinAzure" }).then(() => {});
        return {
          dispose: () => {},
        };
      });
    sandbox.stub(vscode.extensions, "getExtension");
    const signoutStub = sandbox
      .stub(AzureAccountManager.prototype, "signout")
      .throws(new UserCancelError());

    await handlers.registerAccountMenuCommands({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);

    chai.assert.isTrue(signoutStub.called);
  });

  it("openSampleReadmeHandler() - trigger from walkthrough", async () => {
    sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
    sandbox.stub(vscode.workspace, "openTextDocument");
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");

    await handlers.openSampleReadmeHandler(["WalkThrough"]);

    chai.assert.isTrue(executeCommandStub.calledOnce);
  });

  it("showLocalDebugMessage()", async () => {
    sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
    sandbox.stub(vscode.workspace, "openTextDocument");
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");

    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "ShowLocalDebugMessage") {
        return true;
      } else {
        return false;
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .callsFake(
        (title: string, options: vscode.MessageOptions, ...items: vscode.MessageItem[]) => {
          return Promise.resolve({
            title: "Debug",
            run: (options as any).run,
          } as vscode.MessageItem);
        }
      );

    await handlers.showLocalDebugMessage();

    chai.assert.isTrue(executeCommandStub.notCalled);
  });

  it("installAdaptiveCardExt()", async () => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(vscode.extensions, "getExtension").returns(undefined);
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");

    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves("Install" as unknown as vscode.MessageItem);

    await handlers.installAdaptiveCardExt();

    chai.assert.isTrue(executeCommandStub.calledOnce);
  });

  describe("acpInstalled()", () => {
    afterEach(() => {
      mockfs.restore();
    });

    it("already installed", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(vscode.extensions, "getExtension").returns({} as any);

      const installed = handlers.acpInstalled();

      chai.assert.isTrue(installed);
    });

    it("not installed", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(vscode.extensions, "getExtension").returns(undefined);

      const installed = handlers.acpInstalled();

      chai.assert.isFalse(installed);
    });
  });

  it("signInAzure()", async () => {
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");

    await handlers.signInAzure();

    chai.assert.isTrue(executeCommandStub.calledOnce);
  });

  it("signInM365()", async () => {
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");

    await handlers.signInM365();

    chai.assert.isTrue(executeCommandStub.calledOnce);
  });

  it("openLifecycleTreeview() - TeamsFx Project", async () => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "isTeamsFxProject").value(true);
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");

    await handlers.openLifecycleTreeview();

    chai.assert.isTrue(executeCommandStub.calledWith("teamsfx-lifecycle.focus"));
  });

  it("openLifecycleTreeview() - non-TeamsFx Project", async () => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "isTeamsFxProject").value(false);
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");

    await handlers.openLifecycleTreeview();

    chai.assert.isTrue(executeCommandStub.calledWith("workbench.view.extension.teamsfx"));
  });

  it("treeViewDebugInTestToolHandler", async () => {
    sinon.stub(handlers, "core").value(new MockCore());
    sinon.stub(ExtTelemetry, "sendTelemetryEvent");
    const executeCommandStub = sinon.stub(vscode.commands, "executeCommand");

    await handlers.debugInTestToolHandler("treeview");

    chai.assert.isTrue(
      executeCommandStub.calledOnceWith("workbench.action.quickOpen", "debug Debug in Test Tool")
    );
    sinon.restore();
  });

  it("messageDebugInTestToolHandler", async () => {
    sinon.stub(handlers, "core").value(new MockCore());
    sinon.stub(ExtTelemetry, "sendTelemetryEvent");
    const executeCommandStub = sinon.stub(vscode.commands, "executeCommand");

    await handlers.debugInTestToolHandler("message");

    chai.assert.isTrue(
      executeCommandStub.calledOnceWith("workbench.action.quickOpen", "debug Debug in Test Tool")
    );
    sinon.restore();
  });
});
