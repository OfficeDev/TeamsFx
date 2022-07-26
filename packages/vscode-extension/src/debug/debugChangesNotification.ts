// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { LocalEnvManager } from "@microsoft/teamsfx-core";
import * as ExtTelemetry from "../telemetry/telemetry";
import VsCodeLogInstance from "../commonlib/log";
import * as globalVariables from "../globalVariables";
import { vscodeHelper } from "./depsChecker/vscodeHelper";
import { VS_CODE_UI } from "../extension";
import {
  skipNgrokHelpLink,
  skipNgrokRetiredNotification,
  trustDevCertHelpLink,
  trustDevCertRetiredNotification,
} from "./constants";
import { commands } from "vscode";
import { localize } from "../utils/localizeUtils";

// TODO: remove the notification
export async function showDebugChangesNotification(): Promise<void> {
  const localEnvManager = new LocalEnvManager(VsCodeLogInstance, ExtTelemetry.reporter);
  if (!globalVariables.workspaceUri?.fsPath) {
    return;
  }
  const localSettings = await localEnvManager.getLocalSettings(globalVariables.workspaceUri.fsPath);
  if (localSettings?.frontend?.trustDevCert === false && vscodeHelper.isTrustDevCertEnabled()) {
    showNotification(trustDevCertRetiredNotification, trustDevCertHelpLink);
  }

  if (localSettings?.bot?.skipNgrok === true && vscodeHelper.isNgrokCheckerEnabled()) {
    showNotification(skipNgrokRetiredNotification, skipNgrokHelpLink);
  }
}

function showNotification(message: string, url: string): void {
  VS_CODE_UI.showMessage(
    "warn",
    message,
    false,
    localize("teamstoolkit.localDebug.openSettings"),
    localize("teamstoolkit.localDebug.learnMore")
  ).then(async (result) => {
    if (result.isOk()) {
      if (result.value === localize("teamstoolkit.localDebug.learnMore")) {
        await VS_CODE_UI.openUrl(url);
      }
      if (result.value === localize("teamstoolkit.localDebug.openSettings")) {
        await commands.executeCommand(
          "workbench.action.openSettings",
          "fx-extension.prerequisiteCheck"
        );
      }
    }
  });
}
