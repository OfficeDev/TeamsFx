// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { workspace, debug, WorkspaceConfiguration } from "vscode";
import { configurationPrefix, validateDotnetSdkKey, Messages, dotnetHelpLink } from "../constants";
import { DotnetChecker } from "./dotnetChecker";
import * as commonUtils from "../commonUtils";
import logger from "../../commonlib/log";

/**
 * This file contains the extension related functionalities of dotnet checker.
 */

export function dotnetCheckerEnabled(): boolean {
  const configuration: WorkspaceConfiguration = workspace.getConfiguration(configurationPrefix);
  return configuration.get<boolean>(validateDotnetSdkKey, false);
}

export async function tryValidateDotnetInstalled(): Promise<boolean> {
  if (!dotnetCheckerEnabled()) {
    return true;
  }

  try {
    if (!(await DotnetChecker.ensureDotnet())) {
      // TODO(aochengwang): remove this code after using exception to handle errors for ensureDotnet()
      throw new Error("Failed to ensureDotnet(), reason: unknown");
    }
  } catch (e) {
    logger.error(`Failed to ensureDotnet(), exception: '${e}'`);
    await dotnetCheckerFailurePopup();

    // Stop debugging to prevent error popup
    // TODO(aochengwang): stopDebugging() won't stop the "backend extensions install" task
    await debug.stopDebugging();
    return false;
  }

  return true;
}

async function dotnetCheckerFailurePopup() {
  const message = Messages.failToDetectOrInstallDotnet.replace(
    "@ConfigPath",
    DotnetChecker.getDotnetConfigPath()
  );
  commonUtils.displayLearnMore(message, dotnetHelpLink);
}
