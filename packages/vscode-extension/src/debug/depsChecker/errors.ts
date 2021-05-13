// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// NOTE:
// DO NOT EDIT this file in function plugin.
// The source of truth of this file is in packages/vscode-extension/src/debug/depsChecker.
// If you need to edit this file, please edit it in the above folder
// and run the scripts (tools/depsChecker/copyfiles.sh or tools/depsChecker/copyfiles.ps1 according to your OS)
// to copy you changes to function plugin.

export class DepsCheckerError extends Error {
  public readonly helpLink: string;

  constructor(message: string, helpLink: string) {
    super(message);

    this.helpLink = helpLink;
    Object.setPrototypeOf(this, DepsCheckerError.prototype);
  }
}

export class NodeNotFoundError extends DepsCheckerError {
  constructor(message: string, helpLink: string) {
    super(message, helpLink);

    Object.setPrototypeOf(this, NodeNotFoundError.prototype);
  }
}

export class NodeNotSupportedError extends DepsCheckerError {
  constructor(message: string, helpLink: string) {
    super(message, helpLink);

    Object.setPrototypeOf(this, NodeNotSupportedError.prototype);
  }
}

export class BackendExtensionsInstallError extends DepsCheckerError {
  constructor(message: string, helpLink: string) {
    super(message, helpLink);

    Object.setPrototypeOf(this, BackendExtensionsInstallError.prototype);
  }
}
