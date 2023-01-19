// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

"use strict";
import * as vscode from "vscode";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import { TelemetryEvent } from "../telemetry/extTelemetryEvents";
import * as versionUtil from "./versionUtil";
import { isV3Enabled } from "@microsoft/teamsfx-core";
import { PrereleaseState } from "../constants";
import * as folder from "../folder";
import VsCodeLogInstance from "../commonlib/log";
export class PrereleasePage {
  private context: vscode.ExtensionContext;
  constructor(context: vscode.ExtensionContext) {
    this.context = context;
  }
  public async checkAndShow() {
    const extensionId = versionUtil.getExtensionId();
    const teamsToolkit = vscode.extensions.getExtension(extensionId);
    const teamsToolkitVersion = teamsToolkit?.packageJSON.version;
    const prereleaseVersion = this.context.globalState.get<string>(PrereleaseState.Version);
    this.context.globalState.update(PrereleaseState.Version, teamsToolkitVersion);
    if (
      isV3Enabled() &&
      (prereleaseVersion === undefined ||
        (versionUtil.isPrereleaseVersion(prereleaseVersion) &&
          teamsToolkitVersion != prereleaseVersion))
    ) {
      this.context.globalState.update(PrereleaseState.Version, teamsToolkitVersion);
      this.show();
    }
  }
  private async show() {
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.ShowWhatIsNewNotification);
    const uri = vscode.Uri.file(`${folder.getResourceFolder()}/PRERELEASE.md`);
    vscode.workspace.openTextDocument(uri).then(() => {
      const PreviewMarkdownCommand = "markdown.showPreview";
      vscode.commands.executeCommand(PreviewMarkdownCommand, uri);
    });
  }
}
