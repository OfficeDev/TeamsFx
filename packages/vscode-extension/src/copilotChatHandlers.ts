// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as util from "util";
import * as vscode from "vscode";

import { FxError, Result, SystemError, err, ok } from "@microsoft/teamsfx-api";
import { assembleError } from "@microsoft/teamsfx-core";
import VsCodeLogInstance from "./commonlib/log";
import { ExtTelemetry } from "./telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
  TelemetryTriggerFrom,
} from "./telemetry/extTelemetryEvents";
import { getTriggerFromProperty } from "./utils/commonUtils";
import { localize } from "./utils/localizeUtils";
import { UserCancelError, sleep } from "@microsoft/vscode-ui";
import { showOutputChannel } from "./handlers";

const githubCopilotChatExtensionId = "github.copilot-chat";

function githubCopilotInstalled(): boolean {
  const extension = vscode.extensions.getExtension(githubCopilotChatExtensionId);
  return !!extension;
}

async function openGithubCopilotChat(query: string): Promise<Result<null, FxError>> {
  const eventName = "openCopilotChat";
  try {
    const options = {
      query,
      isPartialQuery: true,
    };
    await vscode.commands.executeCommand("workbench.panel.chat.view.copilot.focus");
    await vscode.commands.executeCommand("workbench.action.chat.open", options);
    return ok(null);
  } catch (e) {
    const error = new SystemError(
      eventName,
      "openCopilotError",
      localize("teamstoolkit.handlers.chatTeamsAgentError"),
      localize("teamstoolkit.handlers.chatTeamsAgentError")
    );
    VsCodeLogInstance.error(error.message);
    ExtTelemetry.sendTelemetryErrorEvent(eventName, error);

    const assembledError = assembleError(e);
    if (assembledError.message) {
      VsCodeLogInstance.error(assembledError.message);
    }

    return err(error);
  }
}

export async function installGithubCopilotChatExtension(
  triggerFrom: TelemetryTriggerFrom
): Promise<Result<any, FxError>> {
  const eventName = "installCopilotChat";
  const telemetryProperties = {
    [TelemetryProperty.TriggerFrom]: triggerFrom,
  };
  ExtTelemetry.sendTelemetryEvent(eventName, telemetryProperties);
  try {
    const vscodeVersion = vscode.version;
    const confirmRes = await vscode.window.showInformationMessage(
      localize("teamstoolkit.handlers.askInstallCopilot"),
      localize("teamstoolkit.handlers.askInstallCopilot.install"),
      localize("teamstoolkit.handlers.askInstallCopilot.cancel")
    );

    if (confirmRes !== localize("teamstoolkit.handlers.askInstallCopilot.install")) {
      const error = new UserCancelError(eventName, "cancel");
      ExtTelemetry.sendTelemetryErrorEvent(eventName, error, telemetryProperties);
      return err(error);
    } else {
      await vscode.commands.executeCommand(
        "workbench.extensions.installExtension",
        githubCopilotChatExtensionId,
        {
          installPreReleaseVersion: vscodeVersion.includes("insider"), // VSCode insider need to install Github Copilot Chat of pre-release version
          enable: true,
        }
      );

      ExtTelemetry.sendTelemetryEvent(eventName, {
        ...telemetryProperties,
        [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      });

      return ok(null);
    }
  } catch (e) {
    const error = new SystemError(
      eventName,
      "InstallCopilotError",
      localize("teamstoolkit.handlers.installCopilotError"),
      localize("teamstoolkit.handlers.installCopilotError")
    );
    VsCodeLogInstance.error(error.message);
    ExtTelemetry.sendTelemetryErrorEvent(eventName, error, telemetryProperties);

    const assembledError = assembleError(e);
    if (assembledError.message) {
      VsCodeLogInstance.error(assembledError.message);
    }

    return err(error);
  }
}

export async function invokeTeamsAgent(args?: any[]): Promise<Result<null, FxError>> {
  const eventName = TelemetryEvent.InvokeTeamsAgent;
  const triggerFromProperty = getTriggerFromProperty(args);
  ExtTelemetry.sendTelemetryEvent(eventName, triggerFromProperty);

  const query =
    triggerFromProperty["trigger-from"] === TelemetryTriggerFrom.TreeView
      ? "@teams"
      : "@teams /create";
  let res;

  const isExtensionInstalled = githubCopilotInstalled();
  if (isExtensionInstalled) {
    res = await openGithubCopilotChat(query);
  } else {
    VsCodeLogInstance.info(
      util.format(localize("teamstoolkit.handlers.installExtension.output"), "Github Copilot Chat")
    );
    showOutputChannel();

    const maxRetry = 5;
    const installRes = await installGithubCopilotChatExtension(
      triggerFromProperty[TelemetryProperty.TriggerFrom]
    );
    if (installRes.isOk()) {
      let checkCount = 0;
      let verifyExtensionInstalled = false;
      while (checkCount < maxRetry) {
        verifyExtensionInstalled = githubCopilotInstalled();
        if (!verifyExtensionInstalled) {
          await sleep(3000);
          checkCount++;
        } else {
          break;
        }
      }

      if (verifyExtensionInstalled) {
        await sleep(2000); // wait for extension activation
        res = await openGithubCopilotChat(query);
      } else {
        const error = new SystemError(
          eventName,
          "CannotVerifyGithubCopilotChat",
          localize("teamstoolkit.handlers.verifyCopilotExtensionError"),
          localize("teamstoolkit.handlers.verifyCopilotExtensionError")
        );
        VsCodeLogInstance.error(`Cannot start Teams Agent due to ${error.message}`);
        res = err(error);
      }
    } else {
      res = installRes;
    }
  }
  if (res.isErr()) {
    ExtTelemetry.sendTelemetryErrorEvent(eventName, res.error, triggerFromProperty);
  } else {
    ExtTelemetry.sendTelemetryEvent(eventName, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      ...triggerFromProperty,
    });
  }
  return res;
}
