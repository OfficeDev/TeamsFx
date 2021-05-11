/* eslint-disable @typescript-eslint/no-var-requires */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// NOTE:
// DO NOT EDIT this file in function plugin.
// The source of truth of this file is in packages/vscode-extension/src/debug/depsChecker.
// If you need to edit this file, please edit it in the above folder
// and run the scripts (tools/depsChecker/copyfiles.sh or tools/depsChecker/copyfiles.ps1 according to your OS)
// to copy you changes to function plugin.

import * as os from "os";
const opn = require("opn");

export async function openUrl(url: string): Promise<void> {
  // Using this functionality is blocked by https://github.com/Microsoft/vscode/issues/25852
  // Specifically, opening the Live Metrics Stream for Linux Function Apps doesn't work in this extension.
  // await vscode.env.openExternal(vscode.Uri.parse(url));

  opn(url);
}

export function isWindows(): boolean {
  return os.type() === "Windows_NT";
}

export function isMacOS(): boolean {
  return os.type() === "Darwin";
}

export function isLinux(): boolean {
  return os.type() === "Linux";
}

// help links
export const defaultHelpLink = "https://aka.ms/teamsfx-envchecker-help";

export const nodeNotFoundHelpLink = `${defaultHelpLink}#nodenotfound`;
export const nodeNotSupportedForAzureHelpLink = `${defaultHelpLink}#nodenotsupportedazure-hosting`;
export const nodeNotSupportedForSPFxHelpLink = `${defaultHelpLink}#nodenotsupportedspfx-hosting`;

export const dotnetExplanationHelpLink = `${defaultHelpLink}#overall`;
export const dotnetFailToInstallHelpLink = `${defaultHelpLink}#failtoinstalldotnet`;
export const dotnetManualInstallHelpLink = `${defaultHelpLink}#dotnetnotfound`;
export const dotnetNotSupportTargetVersionHelpLink = `${defaultHelpLink}#dotnetnotsupporttargetversion`;

export const Messages = {
  learnMoreButtonText: "Learn more",
  continueButtonText: "Continue",

  defaultErrorMessage: "Please install the required dependencies manually.",

  // since FuncToolChecker is disabled and azure functions core tools will be installed as devDependencies now,
  // below messages related to FuncToolChecker won't be displayed to end user.
  startInstallFunctionCoreTool: `Downloading and installing @NameVersion.`,
  finishInstallFunctionCoreTool: `Successfully installed @NameVersion.`,
  needReplaceWithFuncCoreToolV3: `You must replace with @NameVersion to debug your local functions.`,
  needInstallFuncCoreTool: `You must have @NameVersion installed to debug your local functions.`,
  failToInstallFuncCoreTool: `@NameVersion installation has failed and will have to be installed manually.`,
  failToValidateFuncCoreTool: `Failed to validate @NameVersion after its installation.`,

  downloadDotnet: `Downloading and installing the portable version of @NameVersion, which will be installed to @InstallDir and won't affect the development environment.`,
  finishInstallDotnet: `Successfully installed @NameVersion.`,
  useGlobalDotnet: `Use global dotnet from PATH.`,
  dotnetInstallStderr: `dotnet-install command failed without error exit code but with non-empty standard error.`,
  dotnetInstallErrorCode: `dotnet-install command failed.`,
  failToInstallDotnet: `Failed to install @NameVersion. Please install @NameVersion manually and restart all your Visual Studio Code instances`,

  NodeNotFound: `The toolkit cannot find Node.js on your machine.

As a fundamental language runtime for Teams app, these dependencies are required. Node.js is required and the recommended version is v12.

Click "Learn more" to learn how to install the Node.js.`,
  NodeNotSupported: `Current installed Node.js (@CurrentVersion) is not in the supported version list (@SupportedVersions), which might not work as expected for some functionalities.

Click "Learn more" to learn more about the supported Node.js versions.
Click "Continue" to continue local debugging.`,

  dotnetNotFound: `The toolkit cannot find @NameVersion on your machine. As a fundamental runtime context for Teams app, it's required. For the details why .NET SDK is needed, please refer to ${dotnetExplanationHelpLink}`,
  depsNotFound: `The toolkit cannot find @SupportedPackages on your machine.

As a fundamental runtime context for Teams app, these dependencies are required. Following steps will help you to install the appropriate version to run the Microsoft Teams Toolkit.

Please notice that these dependencies only need to be installed once.

Click "Install" to install @InstallPackages.`,

  linuxDepsNotFound: `The toolkit cannot find @SupportedPackages on your machine.

As a fundamental runtime context for Teams app, these dependencies are required.

Please install the required dependencies manually.

Click "Continue" to continue.`
};

export enum DepsCheckerEvent {
  // since FuncToolChecker is disabled and azure functions core tools will be installed as devDependencies now,
  // below events related to FuncToolChecker won't be displayed to end user.
  funcCheck = "func-check",
  funcCheckSkipped = "func-check-skipped",
  funcInstall = "func-install",
  funcInstallCompleted = "func-install-completed",
  funcValidation = "func-validation",
  funcValidationCompleted = "func-validation-completed",
  funcV1Installed = "func-v1-installed",
  funcV2Installed = "func-v2-installed",
  funcV3Installed = "func-v3-installed",

  dotnetCheckSkipped = "dotnet-check-skipped",
  dotnetAlreadyInstalled = "dotnet-already-installed",
  dotnetInstallCompleted = "dotnet-install-completed",
  dotnetInstallError = "dotnet-install-error",
  dotnetInstallScriptCompleted = "dotnet-install-script-completed",
  dotnetInstallScriptError = "dotnet-install-script-error",
  dotnetValidationError = "dotnet-validation-error",

  nodeNotFound = "node-not-found",
  nodeNotSupportedForAzure = "node-not-supported-for-azure",
  nodeNotSupportedForSPFx = "node-not-supported-for-spfx"
}

export enum TelemtryMessages {
  failedToInstallFunc = "failed to install Func core tools.",
  funcV1Installed = "func v1 is installed by user.",
  NPMNotFound = "npm is not found.",
  failedToExecDotnetScript = "failed to exec dotnet script.",
  failedToValidateDotnet = "failed to validate dotnet."
}

export enum TelemetryMessurement {
  completionTime = "completion-time"
}
