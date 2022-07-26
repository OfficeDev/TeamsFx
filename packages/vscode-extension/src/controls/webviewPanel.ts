// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, Stage } from "@microsoft/teamsfx-api";
import {
  AppStudioScopes,
  Correlator,
  globalStateGet,
  globalStateUpdate,
  sampleProvider,
} from "@microsoft/teamsfx-core";
import * as AdmZip from "adm-zip";
import axios from "axios";
import { execSync } from "child_process";
import * as fs from "fs-extra";
import { glob } from "glob";
import * as path from "path";
import * as uuid from "uuid";
import * as vscode from "vscode";
import M365TokenInstance from "../commonlib/m365Login";
import AzureAccountManager from "../commonlib/azureLogin";
import { GlobalKey } from "../constants";
import * as globalVariables from "../globalVariables";
import { downloadSample, getSystemInputs } from "../handlers";
import * as ExtTelemetry from "../telemetry/telemetry";
import {
  AccountType,
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
  TelemetryTriggerFrom,
} from "../telemetry/extTelemetryEvents";
import { isMacOS } from "../utils/commonUtils";
import { localize } from "../utils/localizeUtils";
import { Commands } from "./Commands";
import { EventMessages } from "./messages";
import { PanelType } from "./PanelType";

export class WebviewPanel {
  private static readonly viewType = "react";
  public static currentPanels: WebviewPanel[] = [];

  private panel: vscode.WebviewPanel;
  private panelType: PanelType = PanelType.SampleGallery;
  private disposables: vscode.Disposable[] = [];

  public static createOrShow(panelType: PanelType, isToSide?: boolean) {
    const column = vscode.window.activeTextEditor
      ? vscode.window.activeTextEditor.viewColumn
      : undefined;
    if (
      WebviewPanel.currentPanels &&
      WebviewPanel.currentPanels.findIndex((panel) => panel.panelType === panelType) > -1
    ) {
      WebviewPanel.currentPanels
        .find((panel) => panel.panelType === panelType)!
        .panel.reveal(column);
    } else {
      isToSide
        ? WebviewPanel.currentPanels.push(
            new WebviewPanel(panelType, column || vscode.ViewColumn.Two)
          )
        : WebviewPanel.currentPanels.push(
            new WebviewPanel(panelType, column || vscode.ViewColumn.One)
          );
    }
  }

  private constructor(panelType: PanelType, column: vscode.ViewColumn) {
    this.panelType = panelType;

    // Create and show a new webview panel
    this.panel = vscode.window.createWebviewPanel(
      WebviewPanel.viewType,
      this.getWebpageTitle(panelType),
      column,
      {
        // Enable javascript in the webview
        enableScripts: true,
        retainContextWhenHidden: true,
        localResourceRoots: [
          vscode.Uri.file(path.join(globalVariables.context.extensionPath, "out")),
        ],
      }
    );

    // Listen for when the panel is disposed
    // This happens when the user closes the panel or when the panel is closed programatically
    this.panel.onDidDispose(() => this.dispose(), null, this.disposables);

    // Handle messages from the webview
    this.panel.webview.onDidReceiveMessage(
      async (msg) => {
        switch (msg.command) {
          case Commands.OpenExternalLink:
            vscode.env.openExternal(vscode.Uri.parse(msg.data));
            break;
          case Commands.CloneSampleApp:
            Correlator.run(async () => {
              await this.downloadSampleApp(msg);
            });
            break;
          case Commands.DisplayCommands:
            vscode.commands.executeCommand("workbench.action.quickOpen", `>${msg.data}`);
            break;
          case Commands.SigninM365:
            Correlator.run(async () => {
              ExtTelemetry.sendTelemetryEvent(TelemetryEvent.LoginClick, {
                [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.Webview,
                [TelemetryProperty.AccountType]: AccountType.M365,
              });
              await M365TokenInstance.getJsonObject({ scopes: AppStudioScopes });
            });
            break;
          case Commands.SigninAzure:
            Correlator.run(async () => {
              ExtTelemetry.sendTelemetryEvent(TelemetryEvent.LoginClick, {
                [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.Webview,
                [TelemetryProperty.AccountType]: AccountType.Azure,
              });
              await AzureAccountManager.getAccountCredentialAsync(false);
            });
            break;
          case Commands.CreateNewProject:
            await vscode.commands.executeCommand(
              "fx-extension.create",
              TelemetryTriggerFrom.Webview
            );
            break;
          case Commands.SwitchPanel:
            WebviewPanel.createOrShow(msg.data);
            break;
          case Commands.InitAccountInfo:
            this.setStatusChangeMap();
            break;
          case Commands.UpdateGlobalStepsDone:
            await this.updateGlobalStepsDone(msg.data);
            break;
          case Commands.GetGlobalStepsDone:
            await this.getGlobalStepsDone();
            break;
          case Commands.SendTelemetryEvent:
            ExtTelemetry.sendTelemetryEvent(msg.data.eventName, msg.data.properties);
          case Commands.LoadSampleCollection:
            this.LoadSampleCollection();
            break;
          default:
            break;
        }
      },
      undefined,
      globalVariables.context.subscriptions
    );

    // Set the webview's initial html content
    this.panel.webview.html = this.getHtmlForWebview(panelType);
  }

  private async downloadSampleApp(msg: any) {
    const props: any = {
      [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.Webview,
      [TelemetryProperty.SampleAppName]: msg.data.appFolder,
    };
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.DownloadSampleStart, props);
    const inputs: Inputs = getSystemInputs();
    inputs["samples"] = msg.data.appFolder;
    inputs.projectId = inputs.projectId ?? uuid.v4();

    const res = await downloadSample(inputs);
    if (inputs.projectId) {
      props[TelemetryProperty.NewProjectId] = inputs.projectId;
    }
    if (res.isOk()) {
      props[TelemetryProperty.Success] = TelemetrySuccess.Yes;
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.DownloadSample, props);
      await globalStateUpdate(GlobalKey.OpenSampleReadMe, true);
      await globalStateUpdate(GlobalKey.ShowLocalDebugMessage, true);
      await ExtTelemetry.dispose();
      setTimeout(() => {
        vscode.commands.executeCommand("vscode.openFolder", res.value);
      }, 2000);
    } else {
      props[TelemetryProperty.Success] = TelemetrySuccess.No;
      ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.DownloadSample, res.error, props);
    }
  }

  private async updateGlobalStepsDone(data: any) {
    await globalStateUpdate("globalStepsDone", data);
  }

  private async getGlobalStepsDone() {
    const globalStepsDone = await globalStateGet("globalStepsDone", []);
    if (this.panel && this.panel.webview) {
      this.panel.webview.postMessage({
        message: "updateStepsDone",
        data: globalStepsDone,
      });
    }
  }

  private LoadSampleCollection() {
    if (this.panel && this.panel.webview) {
      this.panel.webview.postMessage({
        message: EventMessages.LoadSampleCollection,
        data: sampleProvider.SampleCollection,
      });
    }
  }

  private getWebpageTitle(panelType: PanelType) {
    switch (panelType) {
      case PanelType.SampleGallery:
        return localize("teamstoolkit.webview.samplePageTitle");
      case PanelType.Survey:
        return localize("teamstoolkit.webview.surveyPageTitle");
    }
  }

  private setStatusChangeMap() {
    const m365WebviewCallback = (
      status: string,
      token: string | undefined,
      accountInfo: Record<string, unknown> | undefined
    ) => {
      let email = undefined;
      if (status === "SignedIn") {
        email = (accountInfo as any).upn ? (accountInfo as any).upn : undefined;
      }

      if (this.panel && this.panel.webview) {
        this.panel.webview.postMessage({
          message: "m365AccountChange",
          data: email,
        });
      }

      return Promise.resolve();
    };

    M365TokenInstance.setStatusChangeMap(
      "quick-start-webview",
      { scopes: AppStudioScopes },
      m365WebviewCallback
    );

    AzureAccountManager.setStatusChangeMap(
      "quick-start-webview",
      async (status, token, accountInfo) => {
        let email = undefined;
        if (status === "SignedIn") {
          const token = await AzureAccountManager.getAccountCredentialAsync();
          if (token !== undefined) {
            email = (token as any).username ? (token as any).username : undefined;
          }
        }

        if (this.panel && this.panel.webview) {
          this.panel.webview.postMessage({
            message: "azureAccountChange",
            data: email,
          });
        }

        return Promise.resolve();
      }
    );
  }

  private async fetchCodeZip(url: string) {
    let retries = 3;
    let result = undefined;
    while (retries > 0) {
      retries--;
      try {
        result = await axios.get(url, {
          responseType: "arraybuffer",
        });
        if (result.status === 200 || result.status === 201) {
          return result;
        }
      } catch (e) {
        await new Promise<void>((resolve: () => void): NodeJS.Timer => setTimeout(resolve, 10000));
      }
    }
    return result;
  }

  private async saveFilesRecursively(
    zip: AdmZip,
    appFolder: string,
    dstPath: string
  ): Promise<void> {
    await Promise.all(
      zip
        .getEntries()
        .filter(
          (entry) =>
            !entry.isDirectory &&
            entry.entryName.includes(appFolder) &&
            entry.entryName.split("/").includes(appFolder)
        )
        .map(async (entry) => {
          const entryPath = entry.entryName.substring(entry.entryName.indexOf("/") + 1);
          const filePath = path.join(dstPath, entryPath);
          await fs.ensureDir(path.dirname(filePath));
          await fs.writeFile(filePath, entry.getData());
        })
    );
  }

  private async downloadSampleHook(sampleId: string, sampleAppPath: string) {
    // A temporary solution to avoid duplicate componentId
    if (sampleId === "todo-list-SPFx") {
      const originalId = "c314487b-f51c-474d-823e-a2c3ec82b1ff";
      const componentId = uuid.v4();
      glob.glob(`${sampleAppPath}/**/*.json`, { nodir: true, dot: true }, async (err, files) => {
        await Promise.all(
          files.map(async (file) => {
            let content = (await fs.readFile(file)).toString();
            const reg = new RegExp(originalId, "g");
            content = content.replace(reg, componentId);
            await fs.writeFile(file, content);
          })
        );
      });
    }
  }

  private getHtmlForWebview(panelType: PanelType) {
    const scriptBasePathOnDisk = vscode.Uri.file(
      path.join(globalVariables.context.extensionPath, "out/")
    );
    const scriptBaseUri = scriptBasePathOnDisk.with({ scheme: "vscode-resource" });

    const scriptPathOnDisk = vscode.Uri.file(
      path.join(globalVariables.context.extensionPath, "out/src", "client.js")
    );
    const scriptUri = scriptPathOnDisk.with({ scheme: "vscode-resource" });

    // Use a nonce to to only allow specific scripts to be run
    const nonce = this.getNonce();
    return `<!DOCTYPE html>
        <html lang="en">
          <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>ms-teams</title>
            <base href='${scriptBaseUri}' />
          </head>
          <body>
            <div id="root"></div>
            <script>
              const vscode = acquireVsCodeApi();
              const panelType = '${panelType}';
              const isSupportedNode = ${this.isValidNode()};
              const isMacPlatform = ${isMacOS()};
            </script>
            <script nonce="${nonce}"  type="module" src="${scriptUri}"></script>
          </body>
        </html>`;
  }

  private getNonce() {
    let text = "";
    // eslint-disable-next-line no-secrets/no-secrets
    const possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    for (let i = 0; i < 32; i++) {
      text += possible.charAt(Math.floor(Math.random() * possible.length));
    }
    return text;
  }

  isValidNode = () => {
    try {
      const supportedVersions = ["10", "12", "14"];
      const output = execSync("node --version");
      const regex = /v(?<major_version>\d+)\.(?<minor_version>\d+)\.(?<patch_version>\d+)/gm;

      const match = regex.exec(output.toString());
      if (!match) {
        return false;
      }

      const majorVersion = match.groups?.major_version;
      if (!majorVersion) {
        return false;
      }

      return supportedVersions.includes(majorVersion);
    } catch (e) {}
    return false;
  };

  public dispose() {
    WebviewPanel.currentPanels.splice(WebviewPanel.currentPanels.indexOf(this), 1);

    AzureAccountManager.removeStatusChangeMap("quick-start-webview");

    // Clean up our resources
    this.panel.dispose();

    while (this.disposables.length) {
      const x = this.disposables.pop();
      if (x) {
        x.dispose();
      }
    }
  }
}
