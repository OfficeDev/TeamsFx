// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import {commands, Uri, window, workspace, ExtensionContext, env, ViewColumn} from "vscode";
import {
    Result,
    FxError,
    err,
    ok,
    Stage,
    Platform,
    Func,
    UserError,
    SystemError,
    returnSystemError,
    ConfigFolderName,
    traverse,
    RemoteFuncExecutor,
    Inputs,
    ConfigMap,
    InputResult,
    InputResultType,
    VsCodeEnv,
    AppStudioTokenProvider
} from "fx-api";
import AzureAccountManager from "./commonlib/azureLogin";
import AppStudioTokenInstance from "./commonlib/appStudioLogin";
import AppStudioCodeSpaceTokenInstance from "./commonlib/appStudioCodeSpaceLogin";
import VsCodeLogInstance from "./commonlib/log";
import {CommandsTreeViewProvider, TreeViewCommand} from "./commandsTreeViewProvider";
import {ext} from "./extensionVariables";
import {ExtTelemetry} from "./telemetry/extTelemetry";
import {
    TelemetryEvent,
    TelemetryProperty,
    TelemetryTiggerFrom,
    TelemetrySuccess
} from "./telemetry/extTelemetryEvents";
import * as commonUtils from "./debug/commonUtils";
import {ExtensionErrors, ExtensionSource} from "./error";
import {WebviewPanel} from "./controls/webviewPanel";
import * as constants from "./debug/constants";
import {isFeatureFlag} from "./utils/commonUtils";
import * as fs from "fs-extra";
import * as vscode from "vscode";
import {VS_CODE_UI} from "./qm/vsc_ui";
import {DepsChecker} from "./debug/depsChecker/checker";
import {backendExtensionsInstall} from "./debug/depsChecker/backendExtensionsInstall";
import {FuncToolChecker} from "./debug/depsChecker/funcToolChecker";
import {DotnetChecker} from "./debug/depsChecker/dotnetChecker";
import {PanelType} from "./controls/PanelType";
import {NodeChecker} from "./debug/depsChecker/nodeChecker";
import {TeamsCore} from "fx-core";
import {ContextFactory} from "./context";

export const core = TeamsCore.getInstance();

const runningTasks = new Set<string>(); // to control state of task execution

export async function createNewProjectHandler(args?: any[]): Promise<Result<null, FxError>> {
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.CreateProjectStart, {
        [TelemetryProperty.TriggerFrom]:
            args && args[0] === CommandsTreeViewProvider.TreeViewFlag
                ? TelemetryTiggerFrom.TreeView
                : TelemetryTiggerFrom.CommandPalette
    });
    return await runCommand(Stage.create);
}

export async function updateProjectHandler(): Promise<Result<null, FxError>> {
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.UpdateProjectStart, {
        [TelemetryProperty.TriggerFrom]: TelemetryTiggerFrom.CommandPalette
    });
    return await runCommand(Stage.update);
}

export async function validateManifestHandler(): Promise<Result<null, FxError>> {
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.ValidateManifest, {
        [TelemetryProperty.TriggerFrom]: TelemetryTiggerFrom.CommandPalette
    });

    const func: Func = {
        namespace: "fx-solution-azure",
        method: "validateManifest"
    };
    const ctx = await ContextFactory.get(Stage.userTask);
    return await core.executeUserTask(ctx, func);
}

export async function buildPackageHandler(): Promise<Result<null, FxError>> {
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.BuildPackage, {
        [TelemetryProperty.TriggerFrom]: TelemetryTiggerFrom.CommandPalette
    });

    const func: Func = {
        namespace: "fx-solution-azure",
        method: "buildPackage"
    };
    const ctx = await ContextFactory.get(Stage.userTask);
    return await core.executeUserTask(ctx, func);
}

export async function provisionHandler(): Promise<Result<null, FxError>> {
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.ProvisionStart, {
        [TelemetryProperty.TriggerFrom]: TelemetryTiggerFrom.CommandPalette
    });
    return await runCommand(Stage.provision);
}

export async function deployHandler(): Promise<Result<null, FxError>> {
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.DeployStart, {
        [TelemetryProperty.TriggerFrom]: TelemetryTiggerFrom.CommandPalette
    });
    return await runCommand(Stage.deploy);
}

export async function publishHandler(): Promise<Result<null, FxError>> {
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.PublishStart, {
        [TelemetryProperty.TriggerFrom]: TelemetryTiggerFrom.CommandPalette
    });
    return await runCommand(Stage.publish);
}

const coreExeceutor: RemoteFuncExecutor = async function (
    func: Func,
    answers: Inputs | ConfigMap
): Promise<Result<unknown, FxError>> {
    return await core.callFunc(await ContextFactory.get(Stage.userTask), func, answers as ConfigMap);
};

export async function runCommand(stage: Stage): Promise<Result<null, FxError>> {
    const eventName = ExtTelemetry.stageToEvent(stage);
    let result: Result<null, FxError> = ok(null);

    try {
        // 1. check concurrent lock
        if (runningTasks.size > 0 && stage !== Stage.create) {
            result = err(
                new UserError(
                    ExtensionErrors.ConcurrentTriggerTask,
                    `task '${Array.from(runningTasks).join(",")}' is still running, please wait!`,
                    ExtensionSource
                )
            );
            await processResult(eventName, result);
            return result;
        }

        // 2. lock
        runningTasks.add(stage);

        // 3. check core not empty
        const checkCoreRes = checkCoreNotEmpty();
        if (checkCoreRes.isErr()) {
            throw checkCoreRes.error;
        }

        const answers = new ConfigMap();
        answers.set("stage", stage);
        answers.set("platform", Platform.VSCode);

        // 4. getQuestions
        const qres = await core.getQuestions(await ContextFactory.get(stage));
        if (qres.isErr()) {
            throw qres.error;
        }

        const vscenv = detectVsCodeEnv();
        answers.set("vscenv", vscenv);
        VsCodeLogInstance.info(`VS Code Environment: ${vscenv}`);

        // 5. run question model
        const node = qres.value;
        if (node) {
            VsCodeLogInstance.info(`Question tree:${JSON.stringify(node, null, 4)}`);
            const res: InputResult = await traverse(node, answers, VS_CODE_UI, coreExeceutor);
            VsCodeLogInstance.info(`User input:${JSON.stringify(res, null, 4)}`);
            if (res.type === InputResultType.error) {
                throw res.error!;
            } else if (res.type === InputResultType.cancel) {
                throw new UserError(ExtensionErrors.UserCancel, "User Cancel", ExtensionSource);
            }
        }

        // 6. run task
        const ctx = await ContextFactory.get(stage);
        switch (stage) {
            case (Stage.create): {
                const tmpResult = await core.create(ctx, answers);
                if (tmpResult.isErr()) {
                    result = err(tmpResult.error);
                } else {
                    result = ok(null);
                    // TODO @long open the project
                }
            }
            case (Stage.update): {
                result = await core.update(ctx, answers);
            }
            case (Stage.provision): {
                result = await core.provision(ctx, answers);
            }
            case (Stage.deploy): {
                result = await core.deploy(ctx, answers);
            }
            case (Stage.debug): {
                result = await core.localDebug(ctx, answers);
            }
            case (Stage.publish): {
                result = await core.publish(ctx, answers);
            }
            default: {
                throw new SystemError(
                    ExtensionErrors.UnsupportedOperation,
                    `Operation not support:${stage}`,
                    ExtensionSource
                );
            }
        }
    } catch (e) {
        result = wrapError(e);
    }

    // 7. unlock
    runningTasks.delete(stage);

    // 8. send telemetry and show error
    await processResult(eventName, result);

    return result;
}

export function detectVsCodeEnv(): VsCodeEnv {
    // extensionKind returns ExtensionKind.UI when running locally, so use this to detect remote
    const extension = vscode.extensions.getExtension("Microsoft.teamsfx-extension");

    if (extension?.extensionKind === vscode.ExtensionKind.Workspace) {
        // running remotely
        // Codespaces browser-based editor will return UIKind.Web for uiKind
        if (vscode.env.uiKind === vscode.UIKind.Web) {
            return VsCodeEnv.codespaceBrowser;
        } else {
            return VsCodeEnv.codespaceVsCode;
        }
    } else {
        // running locally
        return VsCodeEnv.local;
    }
}

async function runUserTask(func: Func): Promise<Result<null, FxError>> {
    const eventName = func.method;
    let result: Result<null, FxError> = ok(null);

    try {
        // 1. check concurrent lock
        if (runningTasks.size > 0) {
            result = err(
                new UserError(
                    ExtensionErrors.ConcurrentTriggerTask,
                    `task '${Array.from(runningTasks).join(",")}' is still running, please wait!`,
                    ExtensionSource
                )
            );
            await processResult(eventName, result);
            return result;
        }

        // 2. lock
        runningTasks.add(eventName);

        // 3. check core not empty
        const checkCoreRes = checkCoreNotEmpty();
        if (checkCoreRes.isErr()) {
            throw checkCoreRes.error;
        }

        const answers = new ConfigMap();
        answers.set("task", eventName);
        answers.set("platform", Platform.VSCode);

        // 4. getQuestions
        const ctx = await ContextFactory.get(Stage.userTask);
        const qres = await core.getQuestionsForUserTask(ctx, func);
        if (qres.isErr()) {
            throw qres.error;
        }

        // 5. run question model
        const node = qres.value;
        if (node) {
            VsCodeLogInstance.info(`Question tree:${JSON.stringify(node, null, 4)}`);
            const res: InputResult = await traverse(node, answers, VS_CODE_UI, coreExeceutor);
            VsCodeLogInstance.info(`User input:${JSON.stringify(res, null, 4)}`);
            if (res.type === InputResultType.error && res.error) {
                throw res.error;
            } else if (res.type === InputResultType.cancel) {
                throw new UserError(ExtensionErrors.UserCancel, "User Cancel", ExtensionSource);
            }
        }

        // 6. run task
        result = await core.executeUserTask(ctx, func, answers);
    } catch (e) {
        result = wrapError(e);
    }

    // 7. unlock
    runningTasks.delete(eventName);

    // 8. send telemetry and show error
    await processResult(eventName, result);

    return result;
}

//TODO workaround
function isCancelWarning(error: FxError): boolean {
    return (
        (!!error.name && error.name === ExtensionErrors.UserCancel) ||
        (!!error.message && error.message.includes("User Cancel"))
    );
}
//TODO workaround
function isLoginFaiureError(error: FxError): boolean {
    return !!error.message && error.message.includes("Cannot get user login information");
}

async function processResult(eventName: string, result: Result<null, FxError>) {
    if (result.isErr()) {
        ExtTelemetry.sendTelemetryErrorEvent(eventName, result.error);
        const error = result.error;
        if (isCancelWarning(error)) {
            // window.showWarningMessage(`Operation is canceled!`);
            return;
        }
        if (isLoginFaiureError(error)) {
            window.showErrorMessage(`Login failed, the operation is terminated.`);
            return;
        }
        showError(error);
    } else {
        ExtTelemetry.sendTelemetryEvent(eventName, {
            [TelemetryProperty.Success]: TelemetrySuccess.Yes
        });
    }
}

function wrapError(e: Error): Result<null, FxError> {
    if (
        e instanceof UserError ||
        e instanceof SystemError ||
        (e.constructor &&
            e.constructor.name &&
            (e.constructor.name === "SystemError" || e.constructor.name === "UserError"))
    ) {
        return err(e as FxError);
    }
    return err(returnSystemError(e, ExtensionSource, ExtensionErrors.UnknwonError));
}

function checkCoreNotEmpty(): Result<null, SystemError> {
    if (!core) {
        return err(
            returnSystemError(
                new Error("Core module is not ready!\n Can't do other actions!"),
                ExtensionSource,
                ExtensionErrors.UnsupportedOperation
            )
        );
    }
    return ok(null);
}

/**
 * manually added customized command
 */
export async function updateAADHandler(): Promise<Result<null, FxError>> {
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.UpdateAadStart, {
        [TelemetryProperty.TriggerFrom]: TelemetryTiggerFrom.CommandPalette
    });
    const func: Func = {
        namespace: "fx-solution-azure/fx-resource-aad-app-for-teams",
        method: "aadUpdatePermission"
    };
    return await runUserTask(func);
}


export async function addCapabilityHandler(): Promise<Result<null, FxError>> {
    // ExtTelemetry.sendTelemetryEvent(TelemetryEvent.AddCapStart, {
    //   [TelemetryProperty.TriggerFrom]: TelemetryTiggerFrom.CommandPalette
    // });
    const func: Func = {
        namespace: "fx-solution-azure",
        method: "addCapability"
    };
    return await runUserTask(func);
}

/**
 * check & install required dependencies during local debug.
 */
export async function validateDependenciesHandler(): Promise<void> {
    const depsChecker = new DepsChecker([new NodeChecker(), new FuncToolChecker(), new DotnetChecker()]);
    const shouldContinue = await depsChecker.resolve();
    if (!shouldContinue) {
        // TODO: better mechanism to stop the tasks and debug session.
        throw new Error("debug stopped.");
    }
}

/**
 * install functions binding before launch local debug
 */
export async function backendExtensionsInstallHandler(): Promise<void> {
    if (workspace.workspaceFolders && workspace.workspaceFolders.length > 0) {
        const workspaceFolder = workspace.workspaceFolders[0];
        const backendRoot = await commonUtils.getProjectRoot(
            workspaceFolder.uri.fsPath,
            constants.backendFolderName
        );

        if (backendRoot) {
            await backendExtensionsInstall(backendRoot);
        }
    }
}

/**
 * call localDebug on core, then call customized function to return result
 */
export async function preDebugCheckHandler(): Promise<void> {
    let result: Result<any, FxError> = ok(null);
    result = await runCommand(Stage.debug);
    if (result.isErr()) {
        throw result.error;
    }
    // } catch (e) {
    //   result = wrapError(e);
    //   const eventName = ExtTelemetry.stageToEvent(Stage.debug);
    //   await processResult(eventName, result);
    //   // If debug stage fails, throw error to terminate the debug process
    //   throw result;
    // }
}

export async function mailtoHandler(): Promise<boolean> {
    return env.openExternal(Uri.parse("https://github.com/OfficeDev/teamsfx/issues/new"));
}

export async function openDocumentHandler(): Promise<boolean> {
    return env.openExternal(Uri.parse("https://github.com/OfficeDev/teamsfx/"));
}

export async function devProgramHandler(): Promise<boolean> {
    return env.openExternal(Uri.parse("https://developer.microsoft.com/en-us/microsoft-365/dev-program"));
}

export async function openWelcomeHandler() {
    if (isFeatureFlag()) {
        WebviewPanel.createOrShow(ext.context.extensionPath, PanelType.QuickStart);
    } else {
        const welcomePanel = window.createWebviewPanel("react", "Teams Toolkit", ViewColumn.One, {
            enableScripts: true,
            retainContextWhenHidden: true
        });
        welcomePanel.webview.html = getHtmlForWebview();
    }
}

export async function openSamplesHandler() {
    WebviewPanel.createOrShow(ext.context.extensionPath, PanelType.SampleGallery);
}

export async function openAppManagement() {
    return env.openExternal(Uri.parse("https://dev.teams.microsoft.com/apps"));
}

export async function openBotManagement() {
    return env.openExternal(Uri.parse("https://dev.teams.microsoft.com/bots"));
}

export async function openReportIssues() {
    return env.openExternal(Uri.parse("https://github.com/OfficeDev/TeamsFx/issues"));
}

export async function openManifestHandler(): Promise<Result<null, FxError>> {
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.OpenManifestEditor, {
        [TelemetryProperty.TriggerFrom]: TelemetryTiggerFrom.TreeView
    });
    if (workspace.workspaceFolders && workspace.workspaceFolders.length > 0) {
        const workspaceFolder = workspace.workspaceFolders[0];
        const configRoot = await commonUtils.getProjectRoot(
            workspaceFolder.uri.fsPath,
            `.${ConfigFolderName}`
        );
        const manifestFile = `${configRoot}/${constants.manifestFileName}`;
        if (fs.existsSync(manifestFile)) {
            workspace.openTextDocument(manifestFile).then((document) => {
                window.showTextDocument(document);
            });
            return ok(null);
        } else {
            const FxError: FxError = {
                name: "FileNotFound",
                source: ExtensionSource,
                message: `${manifestFile} not found, cannot open it.`,
                timestamp: new Date()
            };
            showError(FxError);
            return err(FxError);
        }
    } else {
        const FxError: FxError = {
            name: "NoWorkspace",
            source: ExtensionSource,
            message: `No open workspace`,
            timestamp: new Date()
        };
        showError(FxError);
        return err(FxError);
    }
}

// TODO: remove this once welcome page is ready
function getHtmlForWebview() {
    return `<!DOCTYPE html>
  <html>

  <head>
    <meta charset="utf-8" />
    <title>Teams Toolkit</title>
  </head>

  <body>
    <div class="message-container">
      <div class="message">
        Coming Soon...
      </div>
    </div>
    <style type="text/css">
      html {
        height: 100%;
      }

      body {
        box-sizing: border-box;
        min-height: 100%;
        margin: 0;
        padding: 15px 30px;
        display: flex;
        flex-direction: column;
        color: white;
        font-family: "Segoe UI", "Helvetica Neue", "Helvetica", Arial, sans-serif;
        background-color: #2C2C32;
      }

      .message-container {
        flex-grow: 1;
        display: flex;
        align-items: center;
        justify-content: center;
        margin: 0 30px;
      }

      .message {
        font-weight: 300;
        font-size: 1.4rem;
      }
    </style>
  </body>
  </html>`;
}

export async function cmdHdlLoadTreeView(context: ExtensionContext) {
    const treeViewProvider = CommandsTreeViewProvider.getInstance();
    const provider = window.registerTreeDataProvider("teamsfx", treeViewProvider);
    context.subscriptions.push(provider);

    // Register SignOut tree view command
    commands.registerCommand("fx-extension.signOut", async (node: TreeViewCommand) => {
        switch (node.contextValue) {
            case "signedinM365": {
                let appstudioLogin: AppStudioTokenProvider = AppStudioTokenInstance;
                const vscodeEnv = detectVsCodeEnv();
                if (vscodeEnv === VsCodeEnv.codespaceBrowser || vscodeEnv === VsCodeEnv.codespaceVsCode) {
                    appstudioLogin = AppStudioCodeSpaceTokenInstance;
                }
                const result = await appstudioLogin.signout();
                if (result) {
                    await CommandsTreeViewProvider.getInstance().refresh([
                        {
                            commandId: "fx-extension.signinM365",
                            label: "Sign In M365...",
                            contextValue: "signinM365"
                        }
                    ]);
                }
                break;
            }
            case "signedinAzure": {
                const result = await AzureAccountManager.signout();
                if (result) {
                    await CommandsTreeViewProvider.getInstance().refresh([
                        {
                            commandId: "fx-extension.signinAzure",
                            label: "Sign In Azure...",
                            contextValue: "signinAzure"
                        }
                    ]);
                    await CommandsTreeViewProvider.getInstance().remove([
                        {
                            commandId: "fx-extension.selectSubscription",
                            label: "",
                            parent: "fx-extension.signinAzure"
                        }
                    ]);
                }
                break;
            }
        }
    });
}

export function cmdHdlDisposeTreeView() {
    CommandsTreeViewProvider.getInstance().dispose();
}

export async function showError(e: FxError) {
    VsCodeLogInstance.error(`code:${e.source}.${e.name}, message: ${e.message}, stack: ${e.stack}`);

    const errorCode = `${e.source}.${e.name}`;
    if (e instanceof UserError && e.helpLink && typeof e.helpLink != "undefined") {
        const help = {
            title: "Get Help",
            run: async (): Promise<void> => {
                commands.executeCommand("vscode.open", Uri.parse(`${e.helpLink}#${errorCode}`));
            }
        };

        const button = await window.showErrorMessage(`[${errorCode}]: ${e.message}`, help);
        if (button) await button.run();
    } else if (e instanceof SystemError && e.issueLink && typeof e.issueLink != "undefined") {
        const path = e.issueLink.replace(/\/$/, "") + "?";
        const param = `title=new+bug+report: ${errorCode}&body=${e.message}\n\n${e.stack}`;
        const issue = {
            title: "Report Issue",
            run: async (): Promise<void> => {
                commands.executeCommand("vscode.open", Uri.parse(`${path}${param}`));
            }
        };

        const button = await window.showErrorMessage(`[${errorCode}]: ${e.message}`, issue);
        if (button) await button.run();
    } else {
        await window.showErrorMessage(`[${errorCode}]: ${e.message}`);
    }
}
