// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import * as vscode from "vscode";

import {
  AppPackageFolderName,
  BuildFolderName,
  ConfigFolderName,
  CreateProjectResult,
  FxError,
  Result,
} from "@microsoft/teamsfx-api";
import {
  AuthSvcScopes,
  Correlator,
  VersionState,
  initializePreviewFeatureFlags,
  setRegion,
} from "@microsoft/teamsfx-core";

import {
  AadAppTemplateCodeLensProvider,
  ApiPluginCodeLensProvider,
  CopilotPluginCodeLensProvider,
  CryptoCodeLensProvider,
  ManifestTemplateCodeLensProvider,
  OfficeDevManifestCodeLensProvider,
  PermissionsJsonFileCodeLensProvider,
  ProjectSettingsCodeLensProvider,
  TeamsAppYamlCodeLensProvider,
} from "./codeLensProvider";
import commandController from "./commandController";
import AzureAccountManager from "./commonlib/azureLogin";
import VsCodeLogInstance from "./commonlib/log";
import M365TokenInstance from "./commonlib/m365Login";
import { openWelcomePageAfterExtensionInstallation } from "./controls/openWelcomePage";
import { getLocalDebugSessionId, startLocalDebugSession } from "./debug/commonUtils";
import { disableRunIcon, registerRunIcon } from "./debug/runIconHandler";
import { TeamsfxDebugProvider } from "./debug/teamsfxDebugProvider";
import { registerTeamsfxTaskAndDebugEvents } from "./debug/teamsfxTaskHandler";
import { TeamsfxTaskProvider } from "./debug/teamsfxTaskProvider";
import * as exp from "./exp";
import { TreatmentVariableValue, TreatmentVariables } from "./exp/treatmentVariables";
import {
  initializeGlobalVariables,
  isExistingUser,
  isSPFxProject,
  isTeamsFxProject,
  isOfficeAddInProject,
  setUriEventHandler,
  unsetIsTeamsFxProject,
  workspaceUri,
} from "./globalVariables";
import * as handlers from "./handlers";
import * as officeDevHandlers from "./officeDevHandlers";
import { ManifestTemplateHoverProvider } from "./hoverProvider";
import { VsCodeUI } from "./qm/vsc_ui";
import { ExtTelemetry } from "./telemetry/extTelemetry";
import { TelemetryEvent, TelemetryTriggerFrom } from "./telemetry/extTelemetryEvents";
import accountTreeViewProviderInstance from "./treeview/account/accountTreeViewProvider";
import TreeViewManagerInstance from "./treeview/treeViewManager";
import { UriHandler } from "./uriHandler";
import {
  FeatureFlags,
  delay,
  hasAdaptiveCardInWorkspace,
  isM365Project,
} from "./utils/commonUtils";
import { loadLocalizedStrings } from "./utils/localizeUtils";
import { checkProjectTypeAndSendTelemetry } from "./utils/projectChecker";
import { ReleaseNote } from "./utils/releaseNote";
import { ExtensionSurvey } from "./utils/survey";
import { configMgr } from "./config";
import officeDevTreeViewManager from "./treeview/officeDevTreeViewManager";
import {
  CHAT_CREATE_OFFICEADDIN_SAMPLE_COMMAND_ID,
  CHAT_CREATE_SAMPLE_COMMAND_ID,
  CHAT_EXECUTE_COMMAND_ID,
  CHAT_OPENURL_COMMAND_ID,
  IsChatParticipantEnabled,
  chatParticipantName,
  officeAddinChatParticipantName,
} from "./chat/consts";
import followupProvider from "./chat/followupProvider";
import {
  chatCreateCommandHandler,
  chatExecuteCommandHandler,
  chatRequestHandler,
  openUrlCommandHandler,
  handleFeedback,
  officeAddinChatRequestHandler,
} from "./chat/handlers";
import { CommandKey as CommandKeys } from "./constants";

export let VS_CODE_UI: VsCodeUI;

export async function activate(context: vscode.ExtensionContext) {
  initializePreviewFeatureFlags();

  configMgr.registerConfigChangeCallback();

  context.subscriptions.push(new ExtTelemetry.Reporter(context));

  VS_CODE_UI = new VsCodeUI(context);
  initializeGlobalVariables(context);
  loadLocalizedStrings();

  const uriHandler = new UriHandler();
  setUriEventHandler(uriHandler);
  context.subscriptions.push(vscode.window.registerUriHandler(uriHandler));

  registerActivateCommands(context);

  registerInternalCommands(context);

  registerChatParticipant(context);

  registerOfficeAddinChatParticipant(context);

  if (isTeamsFxProject) {
    activateTeamsFxRegistration(context);
  }

  if (isOfficeAddInProject) {
    activateOfficeDevRegistration(context);
  }

  // Call activate function of toolkit core.
  handlers.activate();

  // Init VSC context key
  await initializeContextKey(context, isTeamsFxProject);

  // UI is ready to show & interact
  await vscode.commands.executeCommand("setContext", "fx-extension.isTeamsFx", isTeamsFxProject);

  // control whether to show chat participant entries
  await vscode.commands.executeCommand(
    "setContext",
    "fx-extension.isChatParticipantEnabled",
    IsChatParticipantEnabled
  );

  process.env[FeatureFlags.ChatParticipant] = IsChatParticipantEnabled.toString();

  await vscode.commands.executeCommand(
    "setContext",
    "fx-extension.isOfficeAddIn",
    isOfficeAddInProject
  );

  void VsCodeLogInstance.info("Teams Toolkit extension is now active!");

  // Don't wait this async method to let it run in background.
  void runBackgroundAsyncTasks(context, isTeamsFxProject);
  await vscode.commands.executeCommand("setContext", "fx-extension.initialized", true);
}

// this method is called when your extension is deactivated
export async function deactivate() {
  await ExtTelemetry.cacheTelemetryEventAsync(TelemetryEvent.Deactivate);
  await ExtTelemetry.dispose();
  handlers.cmdHdlDisposeTreeView();
  await disableRunIcon();
}

function activateTeamsFxRegistration(context: vscode.ExtensionContext) {
  registerTreeViewCommandsInDevelopment(context);
  registerTreeViewCommandsInLifecycle(context);
  registerTreeViewCommandsInHelper(context);
  registerTeamsFxCommands(context);
  registerMenuCommands(context);
  handlers.registerAccountMenuCommands(context);

  TreeViewManagerInstance.registerTreeViews(context);
  accountTreeViewProviderInstance.subscribeToStatusChanges({
    azureAccountProvider: AzureAccountManager,
    m365TokenProvider: M365TokenInstance,
  });
  // Set region for M365 account every
  void M365TokenInstance.setStatusChangeMap(
    "set-region",
    { scopes: AuthSvcScopes },
    async (status, token, accountInfo) => {
      if (status === "SignedIn") {
        const tokenRes = await M365TokenInstance.getAccessToken({ scopes: AuthSvcScopes });
        if (tokenRes.isOk()) {
          await setRegion(tokenRes.value);
        }
      }
    }
  );

  if (vscode.workspace.isTrusted) {
    registerCodelensAndHoverProviders(context);
  }

  registerDebugConfigProviders(context);

  // Register task and debug event handlers, as well as sending telemetries
  registerTeamsfxTaskAndDebugEvents();

  registerRunIcon().catch(() => {
    // do nothing
  });

  // Register teamsfx task provider
  const taskProvider: TeamsfxTaskProvider = new TeamsfxTaskProvider();
  context.subscriptions.push(
    vscode.tasks.registerTaskProvider(TeamsfxTaskProvider.type, taskProvider)
  );

  context.subscriptions.push(
    vscode.workspace.onWillSaveTextDocument(handlers.saveTextDocumentHandler)
  );
}

function activateOfficeDevRegistration(context: vscode.ExtensionContext) {
  registerOfficeDevMenuCommands(context);
  officeDevTreeViewManager.registerOfficeDevTreeViews(context);
  if (vscode.workspace.isTrusted) {
    registerOfficeDevCodeLensProviders(context);
  }
}

/**
 * Commands that always show in command palette. They will activate extension and wait for its completion.
 * They are usually used in welcome view and walkthrough.
 */
function registerActivateCommands(context: vscode.ExtensionContext) {
  // non-teamsfx project upgrade
  const checkUpgradeCmd = vscode.commands.registerCommand(
    "fx-extension.checkProjectUpgrade",
    (...args) => Correlator.run(handlers.checkUpgrade, args)
  );
  context.subscriptions.push(checkUpgradeCmd);

  // user can manage account in non-teamsfx project
  const cmpAccountsCmd = vscode.commands.registerCommand("fx-extension.cmpAccounts", (...args) =>
    Correlator.run(handlers.cmpAccountsHandler, args)
  );
  context.subscriptions.push(cmpAccountsCmd);

  // Create a new Teams app
  registerInCommandController(
    context,
    CommandKeys.Create,
    handlers.createNewProjectHandler,
    "createProject"
  );
  context.subscriptions.push(
    vscode.commands.registerCommand("fx-extension.createFromWalkthrough", async (...args) => {
      const res: Result<CreateProjectResult, FxError> = await Correlator.run(
        handlers.createProjectFromWalkthroughHandler,
        args
      );
      if (res.isOk()) {
        const fileUri = vscode.Uri.file(res.value.projectPath);
        const warnings = res.value.warnings;
        await handlers.updateAutoOpenGlobalKey(true, fileUri, warnings, args);
        await ExtTelemetry.dispose();
        await delay(2000);
        return { openFolder: fileUri };
      }
    })
  );

  // Show lifecycle view
  const openLifecycleTreeview = vscode.commands.registerCommand(
    "fx-extension.openLifecycleTreeview",
    (...args) => Correlator.run(handlers.openLifecycleTreeview, args)
  );
  context.subscriptions.push(openLifecycleTreeview);

  // Documentation
  registerInCommandController(context, CommandKeys.OpenDocument, handlers.openDocumentHandler);

  // README
  registerInCommandController(context, CommandKeys.OpenReadMe, handlers.openReadMeHandler);

  // View samples
  registerInCommandController(context, CommandKeys.OpenSamples, handlers.openSamplesHandler);

  // Quick start
  registerInCommandController(context, CommandKeys.OpenWelcome, handlers.openWelcomeHandler);

  // Tutorials
  registerInCommandController(
    context,
    "fx-extension.selectTutorials",
    handlers.selectTutorialsHandler
  );

  // Sign in to M365
  registerInCommandController(context, CommandKeys.SigninM365, handlers.signinM365Callback);

  // Prerequisites check
  registerInCommandController(
    context,
    CommandKeys.ValidateGetStartedPrerequisites,
    handlers.validateGetStartedPrerequisitesHandler
  );

  // Upgrade command to update Teams manifest
  const migrateTeamsManifestCmd = vscode.commands.registerCommand(
    "fx-extension.migrateTeamsManifest",
    () => Correlator.run(handlers.migrateTeamsManifestHandler)
  );
  context.subscriptions.push(migrateTeamsManifestCmd);

  // Upgrade command to update Teams Client SDK
  const migrateTeamsTabAppCmd = vscode.commands.registerCommand(
    "fx-extension.migrateTeamsTabApp",
    () => Correlator.run(handlers.migrateTeamsTabAppHandler)
  );
  context.subscriptions.push(migrateTeamsTabAppCmd);

  // Register local debug run icon
  const runIconCmd = vscode.commands.registerCommand("fx-extension.selectAndDebug", (...args) =>
    Correlator.run(handlers.selectAndDebugHandler, args)
  );
  context.subscriptions.push(runIconCmd);
}

/**
 * Internal commands that will not show in command palette and only be called via executeCommand()
 */
function registerInternalCommands(context: vscode.ExtensionContext) {
  registerInCommandController(
    context,
    "fx-extension.openFromTdp",
    handlers.scaffoldFromDeveloperPortalHandler,
    "openFromTdp"
  );

  const showOutputChannel = vscode.commands.registerCommand(
    "fx-extension.showOutputChannel",
    (...args) => Correlator.run(handlers.showOutputChannel, args)
  );
  context.subscriptions.push(showOutputChannel);

  // Register backend extensions install command
  const backendExtensionsInstallCmd = vscode.commands.registerCommand(
    "fx-extension.backend-extensions-install",
    () => Correlator.runWithId(getLocalDebugSessionId(), handlers.backendExtensionsInstallHandler)
  );
  context.subscriptions.push(backendExtensionsInstallCmd);

  // Referenced by tasks.json
  const getPathDelimiterCmd = vscode.commands.registerCommand(
    "fx-extension.get-path-delimiter",
    () => Correlator.run(handlers.getPathDelimiterHandler)
  );
  context.subscriptions.push(getPathDelimiterCmd);

  const getDotnetPathCmd = vscode.commands.registerCommand("fx-extension.get-dotnet-path", () =>
    Correlator.run(handlers.getDotnetPathHandler)
  );
  context.subscriptions.push(getDotnetPathCmd);

  const installAppInTeamsCmd = vscode.commands.registerCommand(
    "fx-extension.install-app-in-teams",
    () => Correlator.runWithId(getLocalDebugSessionId(), handlers.installAppInTeams)
  );
  context.subscriptions.push(installAppInTeamsCmd);

  const openSurveyCmd = vscode.commands.registerCommand("fx-extension.openSurvey", (...args) =>
    Correlator.run(handlers.openSurveyHandler, [TelemetryTriggerFrom.TreeView])
  );
  context.subscriptions.push(openSurveyCmd);

  const openTutorial = vscode.commands.registerCommand("fx-extension.openTutorial", (...args) =>
    Correlator.run(handlers.openTutorialHandler, [
      TelemetryTriggerFrom.QuickPick,
      ...(args as unknown[]),
    ])
  );
  context.subscriptions.push(openTutorial);

  const preDebugCheckCmd = vscode.commands.registerCommand("fx-extension.pre-debug-check", () =>
    Correlator.runWithId(getLocalDebugSessionId(), handlers.preDebugCheckHandler)
  );
  context.subscriptions.push(preDebugCheckCmd);

  // localdebug session starts from environment checker
  const validateDependenciesCmd = vscode.commands.registerCommand(
    "fx-extension.validate-dependencies",
    () => Correlator.runWithId(startLocalDebugSession(), handlers.validateAzureDependenciesHandler)
  );
  context.subscriptions.push(validateDependenciesCmd);

  // localdebug session starts from prerequisites checker
  const validatePrerequisitesCmd = vscode.commands.registerCommand(
    "fx-extension.validate-local-prerequisites",
    // Do not run with Correlator because it is handled inside `validateLocalPrerequisitesHandler()`.
    handlers.validateLocalPrerequisitesHandler
  );
  context.subscriptions.push(validatePrerequisitesCmd);

  registerInCommandController(context, CommandKeys.SigninAzure, handlers.signinAzureCallback);
}

/**
 * Copilot Chat Participant
 */
function registerChatParticipant(context: vscode.ExtensionContext) {
  const participant = vscode.chat.createChatParticipant(chatParticipantName, (...args) =>
    Correlator.run(chatRequestHandler, ...args)
  );
  participant.iconPath = vscode.Uri.joinPath(context.extensionUri, "media", "teams.png");
  participant.followupProvider = followupProvider;
  participant.onDidReceiveFeedback((...args) => Correlator.run(handleFeedback, ...args));

  context.subscriptions.push(
    participant,
    vscode.commands.registerCommand(CHAT_CREATE_SAMPLE_COMMAND_ID, chatCreateCommandHandler),
    vscode.commands.registerCommand(CHAT_EXECUTE_COMMAND_ID, chatExecuteCommandHandler),
    vscode.commands.registerCommand(CHAT_OPENURL_COMMAND_ID, openUrlCommandHandler)
  );

  const generateManifestGUID = vscode.commands.registerCommand(
    "fx-extension.generateManifestGUID",
    () => Correlator.run(officeDevHandlers.generateManifestGUID)
  );
  context.subscriptions.push(generateManifestGUID);
}

/**
 * Copilot Chat Participant for Office Add-in
 */
function registerOfficeAddinChatParticipant(context: vscode.ExtensionContext) {
  const participant = vscode.chat.createChatParticipant(
    officeAddinChatParticipantName,
    officeAddinChatRequestHandler
  );
  participant.iconPath = vscode.Uri.joinPath(context.extensionUri, "media", "teams.png");
  participant.followupProvider = followupProvider;
  participant.onDidReceiveFeedback((e) => handleFeedback(e));

  context.subscriptions.push(
    participant,
    vscode.commands.registerCommand(
      CHAT_CREATE_OFFICEADDIN_SAMPLE_COMMAND_ID,
      chatCreateCommandHandler
    )
    // vscode.commands.registerCommand(CHAT_EXECUTE_COMMAND_ID, chatExecuteCommandHandler)
    // vscode.commands.registerCommand(CHAT_OPENURL_COMMAND_ID, openUrlCommandHandler)
  );
}

function registerTreeViewCommandsInDevelopment(context: vscode.ExtensionContext) {
  // Open adaptive card
  registerInCommandController(
    context,
    "fx-extension.OpenAdaptiveCardExt",
    handlers.installAdaptiveCardExt
  );

  registerInCommandController(
    context,
    "fx-extension.addWebpart",
    handlers.addWebpart,
    "addWebpart"
  );
}

function registerTreeViewCommandsInLifecycle(context: vscode.ExtensionContext) {
  // Provision in the cloud
  registerInCommandController(
    context,
    CommandKeys.Provision,
    handlers.provisionHandler,
    "provision"
  );

  // Zip Teams metadata package
  registerInCommandController(
    context,
    "fx-extension.build",
    handlers.buildPackageHandler,
    "buildPackage"
  );

  // Deploy to the cloud
  registerInCommandController(context, CommandKeys.Deploy, handlers.deployHandler, "deploy");

  // Publish to Teams
  registerInCommandController(context, CommandKeys.Publish, handlers.publishHandler, "publish");

  // Publish in Developer Portal
  registerInCommandController(
    context,
    "fx-extension.publishInDeveloperPortal",
    handlers.publishInDeveloperPortalHandler,
    "publishInDeveloperPortal"
  );

  // Developer Portal for Teams
  registerInCommandController(
    context,
    "fx-extension.openAppManagement",
    handlers.openAppManagement
  );
}

function registerTreeViewCommandsInHelper(context: vscode.ExtensionContext) {
  // Report issues on GitHub
  registerInCommandController(context, "fx-extension.openReportIssues", handlers.openReportIssues);
}

/**
 * TeamsFx related commands, they will show in command palette after extension is initialized
 */
function registerTeamsFxCommands(context: vscode.ExtensionContext) {
  const createNewEnvironment = vscode.commands.registerCommand(
    // TODO: fix trigger from
    "fx-extension.addEnvironment",
    (...args) => Correlator.run(handlers.createNewEnvironment, args)
  );
  context.subscriptions.push(createNewEnvironment);

  const updateAadAppManifest = vscode.commands.registerCommand(
    "fx-extension.updateAadAppManifest",
    (...args) => Correlator.run(handlers.updateAadAppManifest, args)
  );
  context.subscriptions.push(updateAadAppManifest);

  const updateManifestCmd = vscode.commands.registerCommand(
    "fx-extension.updatePreviewFile",
    (...args) => Correlator.run(handlers.updatePreviewManifest, args)
  );
  context.subscriptions.push(updateManifestCmd);

  const validateManifestCmd = vscode.commands.registerCommand(
    "fx-extension.validateManifest",
    (...args) => Correlator.run(handlers.validateManifestHandler, args)
  );
  context.subscriptions.push(validateManifestCmd);

  const openBotManagementCmd = vscode.commands.registerCommand(
    "fx-extension.openBotManagement",
    (...args) => Correlator.run(handlers.openBotManagement, args)
  );
  context.subscriptions.push(openBotManagementCmd);

  const decryptCmd = vscode.commands.registerCommand(
    "fx-extension.decryptSecret",
    (cipher: string, selection) => Correlator.run(handlers.decryptSecret, cipher, selection)
  );
  context.subscriptions.push(decryptCmd);

  const openConfigStateCmd = vscode.commands.registerCommand(
    "fx-extension.openConfigState",
    (...args) => Correlator.run(handlers.openConfigStateFile, args)
  );
  context.subscriptions.push(openConfigStateCmd);

  const editAadManifestTemplateCmd = vscode.commands.registerCommand(
    "fx-extension.editAadManifestTemplate",
    (...args) => Correlator.run(handlers.editAadManifestTemplate, args)
  );
  context.subscriptions.push(editAadManifestTemplateCmd);

  registerInCommandController(context, CommandKeys.Preview, handlers.treeViewPreviewHandler);

  registerInCommandController(context, "fx-extension.openFolder", handlers.openFolderHandler);

  const checkSideloading = vscode.commands.registerCommand(
    "fx-extension.checkSideloading",
    (...args) => Correlator.run(handlers.checkSideloadingCallback, args)
  );
  context.subscriptions.push(checkSideloading);

  const checkCopilotCallback = vscode.commands.registerCommand(
    "fx-extension.checkCopilotCallback",
    (...args) => Correlator.run(handlers.checkCopilotCallback, args)
  );
  context.subscriptions.push(checkCopilotCallback);
}

/**
 * Commands used in menus, e.g. Explorer context & view item title/context
 */
function registerMenuCommands(context: vscode.ExtensionContext) {
  const createNewEnvironmentWithIcon = vscode.commands.registerCommand(
    "fx-extension.addEnvironmentWithIcon",
    (...args) =>
      Correlator.run(handlers.createNewEnvironment, [TelemetryTriggerFrom.ViewTitleNavigation])
  );
  context.subscriptions.push(createNewEnvironmentWithIcon);

  const azureAccountSettingsCmd = vscode.commands.registerCommand(
    "fx-extension.azureAccountSettings",
    () => Correlator.run(handlers.openAzureAccountHandler)
  );
  context.subscriptions.push(azureAccountSettingsCmd);

  const createAccountCmd = vscode.commands.registerCommand(
    "fx-extension.createAccount",
    (...args) =>
      Correlator.run(handlers.createAccountHandler, [TelemetryTriggerFrom.ViewTitleNavigation])
  );
  context.subscriptions.push(createAccountCmd);

  const manageCollaborator = vscode.commands.registerCommand(
    "fx-extension.manageCollaborator",
    async (node: Record<string, string>) => {
      const envName = node.identifier;
      await Correlator.run(handlers.manageCollaboratorHandler, envName);
    }
  );
  context.subscriptions.push(manageCollaborator);

  registerInCommandController(context, CommandKeys.LocalDebug, handlers.treeViewLocalDebugHandler);

  registerInCommandController(
    context,
    "fx-extension.localdebugWithIcon",
    handlers.treeViewLocalDebugHandler
  );

  registerInCommandController(
    context,
    "fx-extension.debugInTestToolWithIcon",
    handlers.debugInTestToolHandler("treeview")
  );

  registerInCommandController(
    context,
    CommandKeys.DebugInTestToolFromMessage,
    handlers.debugInTestToolHandler("message")
  );

  const m365AccountSettingsCmd = vscode.commands.registerCommand(
    "fx-extension.m365AccountSettings",
    () => Correlator.run(handlers.openM365AccountHandler)
  );
  context.subscriptions.push(m365AccountSettingsCmd);

  const openAccountLinkCmd = vscode.commands.registerCommand(
    "fx-extension.openAccountLink",
    (...args) =>
      Correlator.run(handlers.openAccountLinkHandler, [TelemetryTriggerFrom.ViewTitleNavigation])
  );
  context.subscriptions.push(openAccountLinkCmd);

  const openLifecycleLinkCmd = vscode.commands.registerCommand(
    "fx-extension.openLifecycleLink",
    (...args) =>
      Correlator.run(handlers.openLifecycleLinkHandler, [TelemetryTriggerFrom.ViewTitleNavigation])
  );
  context.subscriptions.push(openLifecycleLinkCmd);

  const openDevelopmentLinkCmd = vscode.commands.registerCommand(
    "fx-extension.openDevelopmentLink",
    (...args) =>
      Correlator.run(handlers.openDevelopmentLinkHandler, [
        TelemetryTriggerFrom.ViewTitleNavigation,
      ])
  );
  context.subscriptions.push(openDevelopmentLinkCmd);

  const openEnvLinkCmd = vscode.commands.registerCommand("fx-extension.openEnvLink", (...args) =>
    Correlator.run(handlers.openEnvLinkHandler, [TelemetryTriggerFrom.ViewTitleNavigation])
  );
  context.subscriptions.push(openEnvLinkCmd);

  const openHelpFeedbackLinkCmd = vscode.commands.registerCommand(
    "fx-extension.openHelpFeedbackLink",
    (...args) =>
      Correlator.run(handlers.openHelpFeedbackLinkHandler, [
        TelemetryTriggerFrom.ViewTitleNavigation,
      ])
  );
  context.subscriptions.push(openHelpFeedbackLinkCmd);

  const openDocumentLinkCmd = vscode.commands.registerCommand(
    "fx-extension.openDocumentLink",
    (...args) => Correlator.run(handlers.openDocumentLinkHandler, args)
  );
  context.subscriptions.push(openDocumentLinkCmd);

  const azureAccountSignOutHelpCmd = vscode.commands.registerCommand(
    "fx-extension.azureAccountSignOutHelp",
    (...args) => Correlator.run(handlers.azureAccountSignOutHelpHandler, args)
  );
  context.subscriptions.push(azureAccountSignOutHelpCmd);

  const aadManifestTemplateCodeLensCmd = vscode.commands.registerCommand(
    "fx-extension.openPreviewAadFile",
    (...args) => Correlator.run(handlers.openPreviewAadFile, args)
  );
  context.subscriptions.push(aadManifestTemplateCodeLensCmd);

  const openResourceGroupInPortal = vscode.commands.registerCommand(
    "fx-extension.openResourceGroupInPortal",
    async (node: Record<string, string>) => {
      const envName = node.identifier;
      await Correlator.run(handlers.openResourceGroupInPortal, envName);
    }
  );
  context.subscriptions.push(openResourceGroupInPortal);

  const openManifestSchemaCmd = vscode.commands.registerCommand(
    "fx-extension.openSchema",
    async (...args) => {
      await Correlator.run(handlers.openExternalHandler, args);
    }
  );
  context.subscriptions.push(openManifestSchemaCmd);

  const addAPICmd = vscode.commands.registerCommand(
    "fx-extension.copilotPluginAddAPI",
    async (...args) => {
      await Correlator.run(handlers.copilotPluginAddAPIHandler, args);
    }
  );
  context.subscriptions.push(addAPICmd);

  const openSubscriptionInPortal = vscode.commands.registerCommand(
    "fx-extension.openSubscriptionInPortal",
    async (node: Record<string, string>) => {
      const envName = node.identifier;
      await Correlator.run(handlers.openSubscriptionInPortal, envName);
    }
  );
  context.subscriptions.push(openSubscriptionInPortal);

  registerInCommandController(
    context,
    "fx-extension.previewWithIcon",
    handlers.treeViewPreviewHandler
  );

  const refreshEnvironment = vscode.commands.registerCommand(
    "fx-extension.refreshEnvironment",
    (...args) =>
      Correlator.run(handlers.refreshEnvironment, [TelemetryTriggerFrom.ViewTitleNavigation])
  );
  context.subscriptions.push(refreshEnvironment);

  const refreshSideloading = vscode.commands.registerCommand(
    "fx-extension.refreshSideloading",
    (...args) => Correlator.run(handlers.refreshSideloadingCallback, args)
  );
  context.subscriptions.push(refreshSideloading);

  const refreshCopilot = vscode.commands.registerCommand("fx-extension.refreshCopilot", (...args) =>
    Correlator.run(handlers.refreshCopilotCallback, args)
  );
  context.subscriptions.push(refreshCopilot);
}

/**
 * Commands used in office dev tree view menus, e.g. Explorer context & view item title/context
 */
function registerOfficeDevMenuCommands(context: vscode.ExtensionContext) {
  // development
  const openDevelopmentLinkCmd = vscode.commands.registerCommand(
    "fx-extension.openOfficeDevDevelopmentLink",
    (...args) => Correlator.run(officeDevHandlers.openDevelopmentLinkHandler, args)
  );
  context.subscriptions.push(openDevelopmentLinkCmd);

  //fx-extension.create and fx-extension.openSamples are registered in registerActivateCommands
  const installDependencyCmd = vscode.commands.registerCommand(
    "fx-extension.installDependency",
    () => Correlator.run(officeDevHandlers.installOfficeAddInDependencies)
  );
  context.subscriptions.push(installDependencyCmd);

  registerInCommandController(context, CommandKeys.LocalDebug, handlers.treeViewLocalDebugHandler);

  const stopDebugging = vscode.commands.registerCommand("fx-extension.stopDebugging", () =>
    Correlator.run(officeDevHandlers.stopOfficeAddInDebug)
  );
  context.subscriptions.push(stopDebugging);

  // lifecycle
  const openLifecycleLinkCmd = vscode.commands.registerCommand(
    "fx-extension.openOfficeDevLifecycleLink",
    (...args) => Correlator.run(officeDevHandlers.openLifecycleLinkHandler, args)
  );
  context.subscriptions.push(openLifecycleLinkCmd);

  const openDeployLinkCmd = vscode.commands.registerCommand(
    "fx-extension.officeDevDeploy",
    (...args) => Correlator.run(officeDevHandlers.openOfficeDevDeployHandler, args)
  );
  context.subscriptions.push(openDeployLinkCmd);

  const publishToAppSourceCmd = vscode.commands.registerCommand(
    "fx-extension.publishToAppSource",
    () => Correlator.run(officeDevHandlers.publishToAppSourceHandler)
  );
  context.subscriptions.push(publishToAppSourceCmd);

  // utility
  const validateManifest = vscode.commands.registerCommand(
    "fx-extension.validateApplication",
    (...args) => Correlator.run(officeDevHandlers.validateOfficeAddInManifest, args)
  );
  context.subscriptions.push(validateManifest);

  const openScriptLabLinkCmd = vscode.commands.registerCommand(
    "fx-extension.openSciptLabLink",
    (...args) => Correlator.run(officeDevHandlers.openScriptLabLink, args)
  );
  context.subscriptions.push(openScriptLabLinkCmd);

  // help and feedback
  const openHelpFeedbackLinkCmd = vscode.commands.registerCommand(
    "fx-extension.openOfficeDevHelpFeedbackLink",
    (...args) => Correlator.run(officeDevHandlers.openHelpFeedbackLinkHandler, args)
  );
  context.subscriptions.push(openHelpFeedbackLinkCmd);

  const openOfficeDevDocumentLinkCmd = vscode.commands.registerCommand(
    "fx-extension.openOfficeDevDocument",
    (...args) => Correlator.run(officeDevHandlers.openDocumentHandler, args)
  );
  context.subscriptions.push(openOfficeDevDocumentLinkCmd);

  const openGetStartedLinkCmd = vscode.commands.registerCommand(
    "fx-extension.openGetStarted",
    (...args) => Correlator.run(officeDevHandlers.openGetStartedLinkHandler, args)
  );
  context.subscriptions.push(openGetStartedLinkCmd);

  const openOfficePartnerCenterLinkCmd = vscode.commands.registerCommand(
    "fx-extension.officePartnerCenter",
    (...args) => Correlator.run(officeDevHandlers.openOfficePartnerCenterHandler, args)
  );
  context.subscriptions.push(openOfficePartnerCenterLinkCmd);

  const reportIssueCmd = vscode.commands.registerCommand(
    "fx-extension.openOfficeDevReportIssues",
    (...args) => Correlator.run(officeDevHandlers.openReportIssues, args)
  );
  context.subscriptions.push(reportIssueCmd);
}

async function initializeContextKey(context: vscode.ExtensionContext, isTeamsFxProject: boolean) {
  await vscode.commands.executeCommand("setContext", "fx-extension.isSPFx", isSPFxProject);

  await vscode.commands.executeCommand(
    "setContext",
    "fx-extension.isM365",
    workspaceUri && (await isM365Project(workspaceUri.fsPath))
  );

  if (isTeamsFxProject) {
    const aadTemplateWatcher = vscode.workspace.createFileSystemWatcher("**/aad.template.json");

    aadTemplateWatcher.onDidCreate(async (event) => {
      await setAadManifestEnabledContext();
    });
  }

  const ymlFileWatcher = vscode.workspace.createFileSystemWatcher(
    "**/teamsapp.yml",
    false,
    true,
    true
  );
  ymlFileWatcher.onDidCreate(async (event) => {
    await detectedTeamsFxProject(context);
  });

  await setAadManifestEnabledContext();
  await setTDPIntegrationEnabledContext();

  const upgradeable = await checkProjectUpgradable();
  if (upgradeable) {
    await vscode.commands.executeCommand("setContext", "fx-extension.canUpgradeV3", true);
    await handlers.checkUpgrade([TelemetryTriggerFrom.Auto]);
  }
}

async function setAadManifestEnabledContext() {
  await vscode.commands.executeCommand("setContext", "fx-extension.isAadManifestEnabled", true);
}

async function setTDPIntegrationEnabledContext() {
  await vscode.commands.executeCommand(
    "setContext",
    "fx-extension.isTDPIntegrationEnabled", // Currently it will return whether v3 is enabled or not.
    true
  );
}

function registerCodelensAndHoverProviders(context: vscode.ExtensionContext) {
  // Setup CodeLens provider for userdata file
  const codelensProvider = new CryptoCodeLensProvider();
  const envDataSelector = {
    scheme: "file",
    pattern: "**/.env.*",
  };

  const projectSettingsCodeLensProvider = new ProjectSettingsCodeLensProvider();
  const projectSettingsSelector = {
    language: "json",
    scheme: "file",
    pattern: `**/.${ConfigFolderName}/configs/projectSettings.json`,
  };

  const manifestTemplateCodeLensProvider = new ManifestTemplateCodeLensProvider();
  const manifestTemplateSelector = {
    language: "json",
    scheme: "file",
    pattern: `**/${AppPackageFolderName}/manifest.json`,
  };
  const localManifestTemplateSelector = {
    language: "json",
    scheme: "file",
    pattern: `**/${AppPackageFolderName}/manifest.local.json`,
  };

  const manifestPreviewSelector = {
    language: "json",
    scheme: "file",
    pattern: `**/${BuildFolderName}/${AppPackageFolderName}/manifest.*.json`,
  };

  const smeOpenapiSpecSelector = {
    language: "yaml",
    scheme: "file",
    pattern: `**/${AppPackageFolderName}/apiSpecificationFile/*.{yml,yaml}`,
  };
  const apiPluginOpenapiSpecSelector: vscode.DocumentSelector = {
    scheme: "file",
    pattern: `**/${AppPackageFolderName}/apiSpecificationFile/*.{yml,yaml,json}`,
  };

  const aadAppTemplateCodeLensProvider = new AadAppTemplateCodeLensProvider();

  const aadAppTemplateSelectorV3 = {
    language: "json",
    scheme: "file",
    pattern: `**/aad.manifest.json`,
  };

  const permissionsJsonFileCodeLensProvider = new PermissionsJsonFileCodeLensProvider();
  const permissionsJsonFileSelector = {
    language: "json",
    scheme: "file",
    pattern: `**/permissions.json`,
  };

  context.subscriptions.push(
    vscode.languages.registerCodeLensProvider(envDataSelector, codelensProvider)
  );
  context.subscriptions.push(
    vscode.languages.registerCodeLensProvider(
      projectSettingsSelector,
      projectSettingsCodeLensProvider
    )
  );
  context.subscriptions.push(
    vscode.languages.registerCodeLensProvider(
      manifestTemplateSelector,
      manifestTemplateCodeLensProvider
    )
  );
  const copilotPluginCodeLensProvider = new CopilotPluginCodeLensProvider();
  context.subscriptions.push(
    vscode.languages.registerCodeLensProvider(
      manifestTemplateSelector,
      copilotPluginCodeLensProvider
    )
  );

  const apiPluginCodeLensProvider = new ApiPluginCodeLensProvider();
  context.subscriptions.push(
    vscode.languages.registerCodeLensProvider(
      apiPluginOpenapiSpecSelector,
      apiPluginCodeLensProvider
    )
  );

  context.subscriptions.push(
    vscode.languages.registerCodeLensProvider(
      localManifestTemplateSelector,
      manifestTemplateCodeLensProvider
    )
  );
  context.subscriptions.push(
    vscode.languages.registerCodeLensProvider(
      manifestPreviewSelector,
      manifestTemplateCodeLensProvider
    )
  );
  context.subscriptions.push(
    vscode.languages.registerCodeLensProvider(
      smeOpenapiSpecSelector,
      manifestTemplateCodeLensProvider
    )
  );
  context.subscriptions.push(
    vscode.languages.registerCodeLensProvider(
      permissionsJsonFileSelector,
      permissionsJsonFileCodeLensProvider
    )
  );

  const aadManifestPreviewSelectorV3 = {
    language: "json",
    scheme: "file",
    pattern: `**/${BuildFolderName}/aad.*.json`,
  };

  const manifestTemplateHoverProvider = new ManifestTemplateHoverProvider();
  context.subscriptions.push(
    vscode.languages.registerHoverProvider(manifestTemplateSelector, manifestTemplateHoverProvider)
  );

  context.subscriptions.push(
    vscode.languages.registerCodeLensProvider(
      aadAppTemplateSelectorV3,
      aadAppTemplateCodeLensProvider
    )
  );

  context.subscriptions.push(
    vscode.languages.registerHoverProvider(
      localManifestTemplateSelector,
      manifestTemplateHoverProvider
    )
  );

  context.subscriptions.push(
    vscode.languages.registerHoverProvider(aadAppTemplateSelectorV3, manifestTemplateHoverProvider)
  );

  context.subscriptions.push(
    vscode.languages.registerCodeLensProvider(
      aadManifestPreviewSelectorV3,
      aadAppTemplateCodeLensProvider
    )
  );

  const yamlCodelensProvider = new TeamsAppYamlCodeLensProvider();
  const yamlFileSelector = {
    language: "yaml",
    scheme: "file",
    pattern: `**/teamsapp.yml`,
  };
  context.subscriptions.push(
    vscode.languages.registerCodeLensProvider(yamlFileSelector, yamlCodelensProvider)
  );
}

function registerOfficeDevCodeLensProviders(context: vscode.ExtensionContext) {
  const officeDevManifestCodeLensProvider = new OfficeDevManifestCodeLensProvider();
  const manifestFileSelector = {
    language: "xml",
    scheme: "file",
    pattern: `**/manifest*.xml`,
  };
  context.subscriptions.push(
    vscode.languages.registerCodeLensProvider(
      manifestFileSelector,
      officeDevManifestCodeLensProvider
    )
  );
}

function registerDebugConfigProviders(context: vscode.ExtensionContext) {
  const debugProvider: TeamsfxDebugProvider = new TeamsfxDebugProvider();
  context.subscriptions.push(
    vscode.debug.registerDebugConfigurationProvider("pwa-chrome", debugProvider)
  );
  context.subscriptions.push(
    vscode.debug.registerDebugConfigurationProvider("chrome", debugProvider)
  );
  context.subscriptions.push(
    vscode.debug.registerDebugConfigurationProvider("pwa-msedge", debugProvider)
  );
  context.subscriptions.push(
    vscode.debug.registerDebugConfigurationProvider("msedge", debugProvider)
  );
}

/**
 * Tasks that doesn't block the user interaction so that they can be processed in background.
 */
async function runBackgroundAsyncTasks(
  context: vscode.ExtensionContext,
  isTeamsFxProject: boolean
) {
  await exp.initialize(context);
  await vscode.commands.executeCommand(
    "setContext",
    "fx-extension.isNewUser",
    isExistingUser === "no"
  );
  TreatmentVariableValue.inProductDoc = await exp
    .getExpService()
    .getTreatmentVariableAsync(
      TreatmentVariables.VSCodeConfig,
      TreatmentVariables.InProductDoc,
      true
    );

  ExtTelemetry.settingsVersion = await handlers.getSettingsVersion();

  await ExtTelemetry.sendCachedTelemetryEventsAsync();
  const releaseNote = new ReleaseNote(context);
  await releaseNote.show();

  if (!isOfficeAddInProject) {
    await openWelcomePageAfterExtensionInstallation();
  }

  if (isTeamsFxProject) {
    await runTeamsFxBackgroundTasks();
  }

  if (isOfficeAddInProject) {
    await runOfficeDevBackgroundTasks();
  }

  const survey = ExtensionSurvey.getInstance();
  survey.activate();

  await recommendACPExtension();

  await checkProjectTypeAndSendTelemetry();
}

async function runTeamsFxBackgroundTasks() {
  const upgradeable = await checkProjectUpgradable();
  if (isTeamsFxProject) {
    await handlers.autoOpenProjectHandler();
    await TreeViewManagerInstance.updateTreeViewsByContent(upgradeable);
  }
}

async function runOfficeDevBackgroundTasks() {
  await officeDevHandlers.autoOpenOfficeDevProjectHandler();
}

function registerInCommandController(
  context: vscode.ExtensionContext,
  name: string,
  callback: (...args: unknown[]) => Promise<Result<unknown, FxError>>,
  runningLabelKey?: string
) {
  commandController.registerCommand(name, callback, runningLabelKey);
  const command = vscode.commands.registerCommand(name, (...args) => {
    if (args[1] === TelemetryTriggerFrom.CopilotChat) {
      return Correlator.runWithId(args[0], runCommand, name, ...args.slice(1));
    }
    return Correlator.run(runCommand, name, ...args);
  });
  context.subscriptions.push(command);
}

function runCommand(commandName: string, ...args: unknown[]) {
  return commandController.runCommand(commandName, ...args);
}

async function checkProjectUpgradable(): Promise<boolean> {
  const versionCheckResult = await handlers.projectVersionCheck();
  if (versionCheckResult.isErr()) {
    unsetIsTeamsFxProject();
    return false;
  }
  const upgradeable = versionCheckResult.isOk()
    ? versionCheckResult.value.isSupport == VersionState.upgradeable
    : false;
  return upgradeable;
}

async function detectedTeamsFxProject(context: vscode.ExtensionContext) {
  const wasTeamsFxProject = isTeamsFxProject;
  initializeGlobalVariables(context);
  if (isTeamsFxProject && !wasTeamsFxProject) {
    activateTeamsFxRegistration(context);

    await vscode.commands.executeCommand("setContext", "fx-extension.isTeamsFx", isTeamsFxProject);

    const aadTemplateWatcher = vscode.workspace.createFileSystemWatcher("**/aad.template.json");

    aadTemplateWatcher.onDidCreate(async (event) => {
      await setAadManifestEnabledContext();
    });

    void runTeamsFxBackgroundTasks();
  }

  const upgradeable = await checkProjectUpgradable();
  if (isTeamsFxProject) {
    await vscode.commands.executeCommand("setContext", "fx-extension.canUpgradeV3", upgradeable);
    await TreeViewManagerInstance.updateTreeViewsByContent(upgradeable);
  }
}

async function recommendACPExtension(): Promise<void> {
  if (!handlers.acpInstalled() && (await hasAdaptiveCardInWorkspace())) {
    await handlers.installAdaptiveCardExt(TelemetryTriggerFrom.Auto);
  }
}
