// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
export enum TelemetryEvent {
  ManageAccount = "manage-account",
  CreateAccountStart = "create-account-start",
  CreateAccount = "create-account",

  GetStarted = "quick-start",

  Samples = "samples",

  Documentation = "documentation",

  LoginClick = "login-click",
  LoginStart = "login-start",
  Login = "login",

  SignOutStart = "sign-out-start",
  SignOut = "sign-out",

  SelectSubscription = "select-subscription",

  CreateProjectStart = "create-project-start",
  CreateProject = "create-project",

  InitProjectStart = "init-project-start",
  InitProject = "init-project",

  RunIconDebugStart = "run-icon-debug-start",
  RunIconDebug = "run-icon-debug",

  AddFeatureStart = "add-feature-start",
  AddFeature = "add-feature",

  OpenManifestEditorStart = "open-manifest-editor-start",
  OpenManifestEditor = "open-manifest-editor",

  ValidateManifestStart = "validate-manifest-start",
  ValidateManifest = "validate-manifest",

  UpdatePreviewManifestStart = "update-preview-manifest-start",
  UpdatePreviewManifest = "update-preview-manifest",

  EditManifestTemplate = "edit-manifest-template",
  EditAadManifestTemplate = "edit-aad-manifest-template",

  getManifestTemplatePath = "get-manifest-path",

  BuildStart = "build-start",
  Build = "build",

  BuildAadManifestStart = "build-aad-manifest-start",
  BuildAadManifest = "build-aad-manifest",

  ProvisionStart = "provision-start",
  Provision = "provision",

  DeployStart = "deploy-start",
  Deploy = "deploy",

  DeployAadManifestStart = "deploy-aad-manifest-start",
  DeployAadManifest = "deploy-aad-manifest",

  UpdateAadStart = "update-aad-start",
  UpdateAad = "update-aad",

  PublishStart = "publish-start",
  Publish = "publish",

  ManageTeamsApp = "manage-teams-app",

  ManageTeamsBot = "manage-teams-bot",

  ReportIssues = "report-issues",

  OpenM365Portal = "open-m365-portal",

  OpenAzurePortal = "open-azure-portal",

  DownloadSampleStart = "download-sample-start",
  DownloadSample = "download-sample",

  WatchVideo = "watch-video",
  PauseVideo = "pause-video",

  DisplayCommands = "display-commands",

  OpenDownloadNode = "open-download-node",

  NextStep = "next-step",

  ClickOpenDeploymentTreeview = "click-open-deployment-tree-view",
  ClickValidatePrerequisites = "click-validate-prerequisites",
  ClickOpenReadMe = "click-open-read-me",
  ViewGuidedTutorials = "view-guided-tutorials",
  OpenTutorial = "open-tutorial",

  GetStartedPrerequisitesStart = "get-started-prerequisites-start",
  GetStartedPrerequisites = "get-started-prerequisites",

  DebugEnvCheckStart = "debug-envcheck-start",
  DebugEnvCheck = "debug-envcheck",
  DebugPreCheckStart = "debug-precheck-start",
  DebugPreCheck = "debug-precheck",
  DebugPrerequisitesStart = "debug-prerequisites-start",
  DebugPrerequisites = "debug-prerequisites",
  DebugStart = "debug-start",
  DebugStop = "debug-stop",
  DebugNpmInstallStart = "debug-npm-install-start",
  DebugNpmInstall = "debug-npm-install",
  DebugServiceStart = "debug-service-start",
  DebugService = "debug-service",
  DebugPrereqsCheckM365Account = "debug-prereqs-check-m365-account",
  DebugPrereqsCheckM365AccountSignIn = "debug-prereqs-check-m365-account-sign-in",
  DebugPrereqsCheckM365Sideloading = "debug-prereqs-check-m365-sideloading",
  DebugPrereqsCheckNode = "debug-prereqs-check-node",
  DebugPrereqsCheckPorts = "debug-prereqs-check-ports",
  DebugPrereqsCheckCert = "debug-prereqs-check-cert",
  DebugPrereqsCheckDependencies = "debug-prereqs-check-dependencies",
  DebugPrereqsCheckNpmInstall = "debug-prereqs-check-npm-install",
  DebugPrereqsInstallPackages = "debug-prereqs-install-packages",
  DebugPreCheckCoreLocalDebug = "debug-precheck-core-local-debug",
  DebugTaskProvider = "debug-task-provider",
  DebugProviderResolveDebugConfiguration = "debug-provider-resolve-debug-configuration",
  DebugAllStart = "debug-all-start",
  DebugAll = "debug-all",

  AutomaticNpmInstallStart = "automatic-npm-install-start",
  AutomaticNpmInstall = "automatic-npm-install",
  ClickDisableAutomaticNpmInstall = "click-disable-automatic-npm-install",

  Survey = "survey",
  SurveyData = "survey-data",

  EditSecretStart = "edit-secret-start",
  EditSecret = "edit-secret",

  OpenManifestConfigStateStart = "open-manifest-config-state-start",
  OpenManifestConfigState = "open-manifest-config-state",

  OpenAadConfigStateStart = "open-aad-config-state-start",
  OpenAadConfigState = "open-aad-config-state",

  OpenTeamsApp = "open-teams-app",
  UpdateTeamsApp = "update-teams-app",

  CreateNewEnvironmentStart = "create-new-environment-start",
  CreateNewEnvironment = "create-new-environment",

  OpenSubscriptionInPortal = "open-subscription-in-portal",
  OpenResourceGroupInPortal = "open-resource-group-in-portal",

  ListCollaboratorStart = "list-collaborator-start",
  ListCollaborator = "list-collaborator",

  GrantPermissionStart = "grant-permission-start",
  GrantPermission = "grant-permission",

  CheckPermissionStart = "check-permission-start",
  CheckPermission = "check-permission",
  OpenSideloadingJoinM365 = "open-sideloading-joinm365",
  OpenSideloadingReadmore = "open-sideloading-readmore",

  ShowWhatIsNewNotification = "show-what-is-new-notification",
  ShowWhatIsNewContext = "show-what-is-new-context",

  ShowLocalDebugNotification = "show-local-debug-notification",
  ShowLocalPreviewNotification = "show-local-preview-notification",
  ClickLocalDebug = "click-local-debug",
  ClickLocalPreview = "click-local-preview",
  PreviewAdaptiveCard = "open-adaptivecard-preview",

  PreviewManifestFile = "preview-manifest",
  PreviewAadManifestFile = "preview-aad-manifest",

  MigrateTeamsTabAppStart = "migrate-teams-tab-app-start",
  MigrateTeamsTabApp = "migrate-teams-tab-app",
  MigrateTeamsTabAppCode = "migrate-teams-tab-app-code",
  MigrateTeamsManifestStart = "migrate-teams-manifest-start",
  MigrateTeamsManifest = "migrate-teams-manifest",

  TreeViewLocalDebug = "treeview-localdebug",

  TreeViewPreviewStart = "treeview-preview-start",
  TreeViewPreview = "treeview-preview",

  ShowOutputChannel = "show-output-channel",
  OpenFolder = "open-folder",

  // To track the effect of UX changes
  // that prevents user performing concurrent operations.
  TreeViewCommandConcurrentExecution = "treeview-command-concurrent-execution",

  Activate = "activate",
  Deactivate = "deactivate",
}

export enum TelemetryProperty {
  Component = "component",
  ProjectId = "project-id",
  CorrelationId = "correlation-id",
  AppId = "appid",
  TenantId = "tenant-id",
  UserId = "hashed-userid",
  AccountType = "account-type",
  TriggerFrom = "trigger-from",
  Success = "success",
  ErrorType = "error-type",
  ErrorCode = "error-code",
  ErrorMessage = "error-message",
  Errors = "errors",
  DebugSessionId = "session-id",
  DebugType = "type",
  DebugRequest = "request",
  DebugPort = "port",
  DebugRemote = "remote",
  DebugAppId = "debug-appid",
  DebugProjectComponents = "debug-project-components",
  DebugDevCertStatus = "debug-dev-cert-status",
  DebugCheckResults = "debug-check-results",
  DebugCheckResultsSafe = "debug-check-results-safe",
  DebugErrorCodes = "debug-error-codes",
  DebugNpmInstallName = "debug-npm-install-name",
  DebugNpmInstallAlreadyInstalled = "debug-npm-install-already-installed",
  DebugNpmInstallExitCode = "debug-npm-install-exit-code",
  DebugNpmInstallErrorMessage = "debug-npm-install-error-message",
  DebugNpmInstallNodeVersion = "debug-npm-install-node-version",
  DebugNpmInstallNpmVersion = "debug-npm-install-npm-version",
  DebugServiceName = "debug-service-name",
  DebugServiceExitCode = "debug-service-exit-code",
  DebugPrereqsDepsType = "debug-prereqs-deps-type",
  DebugFailedServices = "debug-failed-services",
  DebugPortsInUse = "debug-ports-in-use",
  DebugM365AccountStatus = "debug-m365-account-status",
  DebugIsSideloadingAllowed = "debug-is-sideloading-allowed",
  DebugConcurrentCorrelationId = "debug-concurrent-correlation-id",
  DebugConcurrentLastEventName = "debug-concurrent-last-event-name",
  Internal = "internal",
  InternalAlias = "internal-alias",
  OSArch = "os-arch",
  OSRelease = "os-release",
  SampleAppName = "sample-app-name",
  CurrentAction = "current-action",
  VideoPlayFrom = "video-play-from",
  FeatureFlags = "feature-flags",
  UpdateTeamsAppReason = "update-teams-app-reason",
  IsExistingUser = "is-existing-user",
  CollaborationState = "collaboration-state",
  Env = "env",
  SourceEnv = "sourceEnv",
  TargetEnv = "targetEnv",
  IsFromSample = "is-from-sample",
  IsSpfx = "is-spfx",
  IsM365 = "is-m365",
  IsCreatingM365 = "is-creating-m365",
  SettingsVersion = "settings-version",
  UpdateFailedFiles = "update-failed-files",
  NewProjectId = "new-project-id",
  // Used with TreeViewCommandConcurrentExecution
  RunningCommand = "running-command",
  BlockedCommand = "blocked-command",
  // Used with OpenTutorial
  TutorialName = "tutorial-name",
  DocumentationName = "documentation-name",
  // Used with Deactivate
  Timestamp = "timestamp",
  ProgrammingLanguage = "programming-language",
  HostType = "host-type",
  // Used with Activate
  IsTeamsFx = "is-teamsfx",
  ActivationTime = "activation-time",
  UIDisplayTime = "ui-display-time",
  BackgroundTaskFinishTime = "background-task-finish-time",
}

export enum TelemetryMeasurements {
  Duration = "duration",
  DebugPrecheckGapDuration = "debug-precheck-gap-duration",
  DebugServicesGapDuration = "debug-services-gap-duration",
}

export enum TelemetrySuccess {
  Yes = "yes",
  No = "no",
}

export enum TelemetryTriggerFrom {
  CommandPalette = "CommandPalette",
  TreeView = "TreeView",
  ViewTitleNavigation = "ViewTitleNavigation",
  Webview = "Webview",
  CodeLens = "CodeLens",
  EditorTitle = "EditorTitle",
  SideBar = "SideBar",
  WalkThrough = "WalkThrough",
  Notification = "Notification",
  QuickPick = "QuickPick",
  Other = "Other",
  Auto = "Auto",
  Unknow = "Unknow",
}

export enum WatchVideoFrom {
  WatchVideoBtn = "WatchVideoBtn",
  PlayBtn = "PlayBtn",
  WatchOnBrowserBtn = "WatchOnBrowserBtn",
}

export enum TelemetryErrorType {
  UserError = "user",
  SystemError = "system",
}

export enum AccountType {
  M365 = "m365",
  Azure = "azure",
}

export enum TelemetryUpdateAppReason {
  Manual = "manual",
  AfterDelay = "afterDelay",
  FocusOut = "focusOut",
}

export enum TelemetrySurveyDataProperty {
  Q1Title = "q1-title",
  Q1Result = "q1-result",
  Q2Title = "q2-title",
  Q2Result = "q2-result",
  Q3Title = "q3-title",
  Q3Result = "q3-result",
  Q4Title = "q4-title",
  Q4Result = "q4-result",
  Q5Title = "q5-title",
  Q5Result = "q5-result",
}

export enum TelemetryDebugDevCertStatus {
  Disabled = "disabled",
  AlreadyTrusted = "already-trusted",
  Trusted = "trusted",
  NotTrusted = "not-trusted",
}

export interface TelemetryEventCache {
  type: "error" | "normal";
  occurTime: Date;
  sendTime?: Date;
  eventName: string;
  properties?: { [p: string]: string };
  measurements?: { [p: string]: number };
}

export const TelemetryComponentType = "extension";
