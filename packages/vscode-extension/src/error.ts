// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export const ExtensionSource = "Ext";

export enum ExtensionErrors {
  UnknwonError = "UnknwonError",
  UnsupportedOperation = "UnsupportedOperation",
  UserCancel = "UserCancel",
  ConcurrentTriggerTask = "ConcurrentTriggerTask",
  EmptySelectOption = "EmptySelectOption",
  UnsupportedNodeType = "UnsupportedNodeType",
  UnknownSubscription = "UnknownSubscription",
  PortAlreadyInUse = "PortAlreadyInUse",
  PrerequisitesValidationError = "PrerequisitesValidationError",
  PrerequisitesNoM365AccountError = "PrerequisitesNoM365AccountError",
  PrerequisitesSideloadingDisabledError = "PrerequisitesSideloadingDisabledError",
  PrerequisitesInstallPackagesError = "PrerequisitesPackageInstallError",
  DebugServiceFailedBeforeStartError = "DebugServiceFailedBeforeStartError",
  DebugNpmInstallError = "DebugNpmInstallError",
  OpenExternalFailed = "OpenExternalFailed",
  FolderAlreadyExist = "FolderAlreadyExist",
  InvalidProject = "InvalidProject",
  InvalidArgs = "InvalidArgs",
  FetchSampleError = "FetchSampleError",
  EnvConfigNotFoundError = "EnvConfigNotFoundError",
  EnvStateNotFoundError = "EnvStateNotFoundError",
  EnvFileNotFoundError = "EnvFileNotFoundError",
  EnvResourceInfoNotFoundError = "EnvResourceInfoNotFoundError",
  NoWorkspaceError = "NoWorkSpaceError",
  UpdatePackageJsonError = "UpdatePackageJsonError",
  UpdateManifestError = "UpdateManifestError",
  UpdateCodeError = "UpdateCodeError",
  UpdateCodesError = "UpdateCodesError",
  TeamsAppIdNotFoundError = "TeamsAppIdNotFoundError",
  TaskDefinitionError = "TaskDefinitionError",
  TaskCancelError = "TaskCancelError",
  NoTunnelServiceError = "NoTunnelServiceError",
  MultipleTunnelServiceError = "MultipleTunnelServiceError",
  NgrokStoppedError = "NgrokStoppedError",
  NgrokProcessError = "NgrokProcessError",
  NgrokNotFoundError = "NgrokNotFoundError",
  NgrokInstallationError = "NgrokInstallationError",
  TunnelServiceNotStartedError = "TunnelServiceNotStartedError",
  TunnelEndpointNotFoundError = "TunnelEndpointNotFoundError",
  TunnelEnvError = "TunnelEnvError",
  StartTunnelError = "StartTunnelError",
  NgrokTimeoutError = "NgrokTimeoutError",
  LaunchTeamsWebClientError = "LaunchTeamsWebClientError",
  SetUpTabError = "SetUpTabError",
  SetUpBotError = "SetUpBotError",
  SetUpSSOError = "SetUpSSOError",
  PrepareManifestError = "PrepareManifestError",
  LoginCacheError = "LoginCacheError",
  DefaultManifestTemplateNotExistsError = "DefaultManifestTemplateNotExistsError",
  DefaultAppPackageNotExistsError = "DefaultAppPackageNotExistsError",
  DevTunnelStartError = "DevTunnelStartError",
}
