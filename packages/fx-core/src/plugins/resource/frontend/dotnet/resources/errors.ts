// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Logger } from "../../utils/logger";
import { ConfigFolderName } from "@microsoft/teamsfx-api";
import { FrontendPluginError } from "../../resources/errors";

export enum ErrorType {
  User,
  System,
}

const tips = {
  checkLog: "Check log for more information.",
  doProvision: `Run 'Provision Resource' before this command.`,
  reProvision: `Run 'Provision' command again.`,
  reDeploy: "Run 'Deploy' command again.",
  checkNetwork: "Check your network connection.",
  checkFsPermissions: "Check if you have Read/Write permissions to your file system.",
  restoreEnvironment: `If you manually updated configuration files (under directory .${ConfigFolderName}), recover them.`,
};

export class DotnetPluginError extends FrontendPluginError {
  public innerError?: Error;

  constructor(
    errorType: ErrorType,
    code: string,
    messages: [string, string],
    suggestions: string[],
    helpLink?: string,
    innerError?: Error
  ) {
    super(errorType, code, messages, suggestions, helpLink);
    this.innerError = innerError;
  }

  getMessage(): string {
    return `${this.messages[0]} Suggestions: ${this.suggestions.join(" ")}`;
  }
  getDefaultMessage(): string {
    return `${this.messages[1]} Suggestions: ${this.suggestions.join(" ")}`;
  }
  setInnerError(error: Error): void {
    this.innerError = error;
  }

  getInnerError(): Error | undefined {
    return this.innerError;
  }
}

export class NoProjectSettingError extends DotnetPluginError {
  constructor() {
    super(
      ErrorType.System,
      "NoProjectSettingError",
      ["Failed to load project setting", "Failed to load project setting"],
      []
    );
  }
}

export class FetchConfigError extends DotnetPluginError {
  constructor(key: string) {
    super(
      ErrorType.User,
      "FetchConfigError",
      [`Failed to find ${key} from configuration`, `Failed to find ${key} from configuration`],
      [tips.restoreEnvironment]
    );
  }
}

export class ProjectPathError extends DotnetPluginError {
  constructor(projectFilePath: string) {
    super(
      ErrorType.User,
      "ProjectPathError",
      [
        `Failed to find target project ${projectFilePath}.`,
        `Failed to find target project ${projectFilePath}.`,
      ],
      [tips.checkLog, tips.restoreEnvironment]
    );
  }
}

export class BuildError extends DotnetPluginError {
  constructor(innerError?: Error) {
    super(
      ErrorType.User,
      "BuildError",
      ["Failed to build Dotnet project.", "Failed to build Dotnet project."],
      [tips.checkLog, tips.reDeploy],
      undefined,
      innerError
    );
  }
}

export class ZipError extends DotnetPluginError {
  constructor() {
    super(
      ErrorType.User,
      "ZipError",
      ["Failed to generate zip package.", "Failed to generate zip package."],
      [tips.checkFsPermissions, tips.reDeploy]
    );
  }
}

export class PublishCredentialError extends DotnetPluginError {
  constructor() {
    super(
      ErrorType.User,
      "PublishCredentialError",
      ["Failed to retrieve publish credential.", "Failed to retrieve publish credential."],
      [tips.doProvision, tips.reDeploy]
    );
  }
}

export class UploadZipError extends DotnetPluginError {
  constructor() {
    super(
      ErrorType.User,
      "UploadZipError",
      ["Failed to upload zip package.", "Failed to upload zip package."],
      [tips.checkNetwork, tips.reDeploy]
    );
  }
}

export class FileIOError extends DotnetPluginError {
  constructor(path: string) {
    super(
      ErrorType.User,
      "FileIOError",
      [`Failed to read/write ${path}.`, `Failed to read/write ${path}.`],
      [tips.checkFsPermissions, tips.checkLog]
    );
  }
}

export const UnhandledErrorCode = "UnhandledError";
export const UnhandledErrorMessage = "Unhandled error.";

export async function runWithErrorCatchAndThrow<T>(
  error: DotnetPluginError,
  fn: () => T | Promise<T>
): Promise<T> {
  try {
    return await Promise.resolve(fn());
  } catch (e: any) {
    Logger.error(e.toString());
    error.setInnerError(e);
    throw error;
  }
}

export async function runWithErrorCatchAndWrap<T>(
  wrap: (error: any) => DotnetPluginError,
  fn: () => T | Promise<T>
): Promise<T> {
  try {
    return await Promise.resolve(fn());
  } catch (e: any) {
    Logger.error(e.toString());
    const error = wrap(e);
    error.setInnerError(e);
    throw error;
  }
}
