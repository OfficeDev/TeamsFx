// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetryTriggerFrom,
} from "../telemetry/extTelemetryEvents";
import { getTriggerFromProperty } from "../utils/telemetryUtils";
import { FxError, ok, Result } from "@microsoft/teamsfx-api";
import { TreeViewCommand } from "../treeview/treeViewCommand";
import { AppStudioScopes, featureFlagManager, FeatureFlags } from "@microsoft/teamsfx-core";
import { VS_CODE_UI } from "../qm/vsc_ui";
import { DeveloperPortalHomeLink, PublishAppLearnMoreLink } from "../constants";
import { signedIn } from "../commonlib/common/constant";
import M365TokenInstance from "../commonlib/m365Login";

export async function openEnvLinkHandler(args: any[]): Promise<boolean> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Documentation, {
    ...getTriggerFromProperty(args),
    [TelemetryProperty.DocumentationName]: "environment",
  });
  return vscode.env.openExternal(vscode.Uri.parse("https://aka.ms/teamsfx-treeview-environment"));
}

export async function openDevelopmentLinkHandler(args: any[]): Promise<boolean> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Documentation, {
    ...getTriggerFromProperty(args),
    [TelemetryProperty.DocumentationName]: "development",
  });
  return vscode.env.openExternal(vscode.Uri.parse("https://aka.ms/teamsfx-treeview-development"));
}

export async function openLifecycleLinkHandler(args: any[]): Promise<boolean> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Documentation, {
    ...getTriggerFromProperty(args),
    [TelemetryProperty.DocumentationName]: "lifecycle",
  });
  return vscode.env.openExternal(vscode.Uri.parse("https://aka.ms/teamsfx-treeview-deployment"));
}

export async function openHelpFeedbackLinkHandler(args: any[]): Promise<boolean> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Documentation, {
    ...getTriggerFromProperty(args),
    [TelemetryProperty.DocumentationName]: "help&feedback",
  });
  return vscode.env.openExternal(vscode.Uri.parse("https://aka.ms/teamsfx-treeview-helpnfeedback"));
}

export async function openWelcomeHandler(...args: unknown[]): Promise<Result<unknown, FxError>> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.GetStarted, getTriggerFromProperty(args));
  const data = await vscode.commands.executeCommand(
    "workbench.action.openWalkthrough",
    getWalkThroughId()
  );
  return Promise.resolve(ok(data));
}

export async function openDocumentLinkHandler(args?: any[]): Promise<Result<boolean, FxError>> {
  if (!args || args.length < 1) {
    // should never happen
    return Promise.resolve(ok(false));
  }
  const node = args[0] as TreeViewCommand;
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Documentation, {
    [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.TreeView,
    [TelemetryProperty.DocumentationName]: node.contextValue!,
  });
  switch (node.contextValue) {
    case "signinM365": {
      await vscode.commands.executeCommand("workbench.action.openWalkthrough", {
        category: getWalkThroughId(),
        step: `${getWalkThroughId()}#teamsToolkitCreateFreeAccount`,
      });
      return Promise.resolve(ok(true));
    }
    case "signinAzure": {
      return VS_CODE_UI.openUrl("https://portal.azure.com/");
    }
    case "fx-extension.create":
    case "fx-extension.openSamples": {
      return VS_CODE_UI.openUrl("https://aka.ms/teamsfx-create-project");
    }
    case "fx-extension.provision": {
      return VS_CODE_UI.openUrl("https://aka.ms/teamsfx-provision-cloud-resource");
    }
    case "fx-extension.build": {
      return VS_CODE_UI.openUrl("https://aka.ms/teams-store-validation");
    }
    case "fx-extension.deploy": {
      return VS_CODE_UI.openUrl("https://aka.ms/teamsfx-deploy");
    }
    case "fx-extension.publish": {
      return VS_CODE_UI.openUrl("https://aka.ms/teamsfx-publish");
    }
    case "fx-extension.publishInDeveloperPortal": {
      return VS_CODE_UI.openUrl(PublishAppLearnMoreLink);
    }
  }
  return Promise.resolve(ok(false));
}

export async function openM365AccountHandler() {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.OpenM365Portal);
  return vscode.env.openExternal(vscode.Uri.parse("https://admin.microsoft.com/Adminportal/"));
}

export async function openAzureAccountHandler() {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.OpenAzurePortal);
  return vscode.env.openExternal(vscode.Uri.parse("https://portal.azure.com/"));
}

export function getWalkThroughId(): string {
  return featureFlagManager.getBooleanValue(FeatureFlags.ChatParticipant)
    ? "TeamsDevApp.ms-teams-vscode-extension#teamsToolkitGetStartedWithChat"
    : "TeamsDevApp.ms-teams-vscode-extension#teamsToolkitGetStarted";
}

export async function openAppManagement(...args: unknown[]): Promise<Result<boolean, FxError>> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.ManageTeamsApp, getTriggerFromProperty(args));
  const accountRes = await M365TokenInstance.getStatus({ scopes: AppStudioScopes });

  if (accountRes.isOk() && accountRes.value.status === signedIn) {
    const loginHint = accountRes.value.accountInfo?.upn as string;
    return VS_CODE_UI.openUrl(`${DeveloperPortalHomeLink}?login_hint=${loginHint}`);
  } else {
    return VS_CODE_UI.openUrl(DeveloperPortalHomeLink);
  }
}

export async function openBotManagement(args?: any[]) {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.ManageTeamsBot, getTriggerFromProperty(args));
  return vscode.env.openExternal(vscode.Uri.parse("https://dev.teams.microsoft.com/bots"));
}

export async function openAccountLinkHandler(args: any[]): Promise<boolean> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Documentation, {
    ...getTriggerFromProperty(args),
    [TelemetryProperty.DocumentationName]: "account",
  });
  return vscode.env.openExternal(vscode.Uri.parse("https://aka.ms/teamsfx-treeview-account"));
}

export async function openReportIssues(...args: unknown[]): Promise<Result<boolean, FxError>> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.ReportIssues, getTriggerFromProperty(args));
  return VS_CODE_UI.openUrl("https://github.com/OfficeDev/TeamsFx/issues");
}

export async function openDocumentHandler(...args: unknown[]): Promise<Result<boolean, FxError>> {
  let documentName = "general";
  if (args && args.length >= 2) {
    documentName = args[1] as string;
  }
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Documentation, {
    ...getTriggerFromProperty(args),
    [TelemetryProperty.DocumentationName]: documentName,
  });
  let url = "https://aka.ms/teamsfx-build-first-app";
  if (documentName === "learnmore") {
    url = "https://aka.ms/teams-toolkit-5.0-upgrade";
  }
  return VS_CODE_UI.openUrl(url);
}
