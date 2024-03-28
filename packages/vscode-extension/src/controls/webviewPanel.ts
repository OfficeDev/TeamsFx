// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as path from "path";
import * as uuid from "uuid";
import * as vscode from "vscode";

import { Inputs } from "@microsoft/teamsfx-api";
import {
  Correlator,
  SampleConfig,
  isValidOfficeAddInProject,
  sampleProvider,
} from "@microsoft/teamsfx-core";

import * as extensionPackage from "../../package.json";
import { TreatmentVariableValue } from "../exp/treatmentVariables";
import * as globalVariables from "../globalVariables";
import { downloadSample, getSystemInputs, openFolder } from "../handlers";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  InProductGuideInteraction,
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
  TelemetryTriggerFrom,
} from "../telemetry/extTelemetryEvents";
import { localize } from "../utils/localizeUtils";
import { compare } from "../utils/versionUtil";
import { Commands } from "./Commands";
import { PanelType } from "./PanelType";
import { isTriggerFromWalkThrough } from "../utils/commonUtils";
import { openOfficeDevFolder } from "../officeDevHandlers";

export class WebviewPanel {
  private static readonly viewType = "react";
  public static currentPanels: WebviewPanel[] = [];

  private panel: vscode.WebviewPanel;
  private panelType: PanelType = PanelType.SampleGallery;
  private disposables: vscode.Disposable[] = [];

  public static createOrShow(panelType: PanelType, args?: any[]) {
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
      const isToSide = isTriggerFromWalkThrough(args);
      isToSide
        ? WebviewPanel.currentPanels.push(
            new WebviewPanel(panelType, column || vscode.ViewColumn.Two)
          )
        : WebviewPanel.currentPanels.push(
            new WebviewPanel(panelType, column || vscode.ViewColumn.One)
          );
    }
    // if args empty or undefined, return
    if (!args?.length) {
      return;
    }
    if (panelType == PanelType.SampleGallery && args.length > 1) {
      try {
        const sampleId = args[1] as string;
        const panel = WebviewPanel.currentPanels.find((panel) => panel.panelType === panelType);
        if (panel) {
          void panel.panel.webview.postMessage({
            message: Commands.OpenDesignatedSample,
            sampleId: sampleId,
          });
        }
      } catch (e) {}
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

    this.panel.onDidChangeViewState(
      (e) => {
        const panel = e.webviewPanel;
        if (
          TreatmentVariableValue.inProductDoc &&
          (panelType === PanelType.RespondToCardActions ||
            panelType === PanelType.FunctionBasedNotificationBotReadme ||
            panelType === PanelType.RestifyServerNotificationBotReadme)
        ) {
          ExtTelemetry.sendTelemetryEvent(TelemetryEvent.InteractWithInProductDoc, {
            [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.InProductDoc,
            [TelemetryProperty.Interaction]: panel.visible
              ? InProductGuideInteraction.Show
              : InProductGuideInteraction.Hide,
            [TelemetryProperty.Identifier]: panelType,
          });
        } else if (panelType === PanelType.AccountHelp) {
          ExtTelemetry.sendTelemetryEvent(TelemetryEvent.InteractWithInProductDoc, {
            [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.AccountHelp,
            [TelemetryProperty.Interaction]: panel.visible
              ? InProductGuideInteraction.Show
              : InProductGuideInteraction.Hide,
            [TelemetryProperty.Identifier]: panelType,
          });
        }
      },
      null,
      globalVariables.context.subscriptions
    );

    // Handle messages from the webview
    this.panel.webview.onDidReceiveMessage(
      async (msg) => {
        switch (msg.command) {
          case Commands.OpenExternalLink:
            void vscode.env.openExternal(vscode.Uri.parse(msg.data));
            break;
          case Commands.CloneSampleApp:
            await Correlator.run(async () => {
              await this.downloadSampleApp(msg);
            });
            break;
          case Commands.DisplayCommands:
            // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
            await vscode.commands.executeCommand("workbench.action.quickOpen", `>${msg.data}`);
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
          case Commands.SendTelemetryEvent:
            ExtTelemetry.sendTelemetryEvent(msg.data.eventName, msg.data.properties);
            break;
          case Commands.LoadSampleCollection:
            await this.LoadSampleCollection();
            break;
          case Commands.LoadSampleReadme:
            await this.LoadSampleReadme(msg.data);
            break;
          case Commands.UpgradeToolkit:
            await this.OpenToolkitInExtensionView(msg.data.version);
            break;
          case Commands.StoreData:
            await globalVariables.context.globalState.update(msg.data.key, msg.data.value);
            break;
          case Commands.GetData:
            await this.panel.webview.postMessage({
              message: Commands.GetData,
              data: {
                key: msg.data.key,
                value: globalVariables.context.globalState.get(msg.data.key),
              },
            });
            break;
          case Commands.InvokeTeamsAgent:
            await vscode.commands.executeCommand(
              "fx-extension.invokeChat",
              TelemetryTriggerFrom.Webview
            );
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
    this.panel.iconPath = this.getWebviewPanelIconPath(panelType);
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
      if (isValidOfficeAddInProject((res.value as vscode.Uri).fsPath)) {
        await openOfficeDevFolder(res.value, true);
      } else {
        await openFolder(res.value, true);
      }
    } else {
      props[TelemetryProperty.Success] = TelemetrySuccess.No;
      ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.DownloadSample, res.error, props);
    }
  }

  private async LoadSampleCollection() {
    try {
      await sampleProvider.refreshSampleConfig();
    } catch (e: unknown) {
      await this.panel.webview.postMessage({
        message: Commands.LoadSampleCollection,
        samples: [],
        error: e,
      });
      return;
    }
    const sampleCollection = await sampleProvider.SampleCollection;
    const sampleData = sampleCollection.samples.map((sample) => {
      const extensionVersion = extensionPackage.version;
      let versionComparisonResult = 0;
      if (
        sample.maximumToolkitVersion &&
        compare(extensionVersion, sample.maximumToolkitVersion) > 0
      ) {
        versionComparisonResult = 1;
      }
      if (
        sample.minimumToolkitVersion &&
        compare(extensionVersion, sample.minimumToolkitVersion) < 0
      ) {
        versionComparisonResult = -1;
      }
      return {
        ...sample,
        versionComparisonResult,
      };
    });
    if (this.panel && this.panel.webview) {
      await this.panel.webview.postMessage({
        message: Commands.LoadSampleCollection,
        samples: sampleData,
        filterOptions: sampleCollection.filterOptions,
      });
    }
  }

  private async LoadSampleReadme(sample: SampleConfig) {
    let htmlContent = "";
    try {
      htmlContent = await sampleProvider.getSampleReadmeHtml(sample);
    } catch (e: unknown) {
      await this.panel.webview.postMessage({
        message: Commands.LoadSampleReadme,
        error: e,
        readme: "",
      });
      return;
    }
    if (this.panel && this.panel.webview) {
      const readme = this.replaceRelativeImagePaths(htmlContent, sample);
      await this.panel.webview.postMessage({
        message: Commands.LoadSampleReadme,
        readme: readme,
      });
    }
  }

  private async OpenToolkitInExtensionView(version: string) {
    // await vscode.commands.executeCommand(
    //   "workbench.extensions.installExtension",
    //   `teamsdevapp.ms-teams-vscode-extension@${version}`
    // );
    await vscode.commands.executeCommand("workbench.extensions.action.checkForUpdates");
  }

  private replaceRelativeImagePaths(htmlContent: string, sample: SampleConfig) {
    const urlInfo = sample.downloadUrlInfo;
    const imageUrl = `https://github.com/${urlInfo.owner}/${urlInfo.repository}/blob/${urlInfo.ref}/${urlInfo.dir}/${sample.thumbnailPath}?raw=1`;
    const imageRegex = /img\s+src="([^"]+)"/gm;
    return htmlContent.replace(imageRegex, `img src="${imageUrl}"`);
  }

  private getWebpageTitle(panelType: PanelType): string {
    switch (panelType) {
      case PanelType.SampleGallery:
        return localize("teamstoolkit.webview.samplePageTitle");
      case PanelType.Survey:
        return localize("teamstoolkit.webview.surveyPageTitle");
      case PanelType.RespondToCardActions:
        return localize("teamstoolkit.guides.cardActionResponse.label");
      case PanelType.AccountHelp:
        return localize("teamstoolkit.webview.accountHelp");
      case PanelType.RestifyServerNotificationBotReadme:
        return localize("teamstoolkit.guides.notificationBot.label");
      case PanelType.FunctionBasedNotificationBotReadme:
        return localize("teamstoolkit.guides.notificationBot.label");
    }
  }

  private getHtmlForWebview(panelType: PanelType) {
    const scriptBasePathOnDisk = vscode.Uri.file(
      path.join(globalVariables.context.extensionPath, "out/")
    );
    const scriptBaseUri = this.panel.webview.asWebviewUri(scriptBasePathOnDisk);

    const scriptPathOnDisk = vscode.Uri.file(
      path.join(globalVariables.context.extensionPath, "out/src", "client.js")
    );
    const scriptUri = this.panel.webview.asWebviewUri(scriptPathOnDisk);
    const codiconsUri = this.panel.webview.asWebviewUri(
      vscode.Uri.joinPath(globalVariables.context.extensionUri, "out", "resource", "codicon.css")
    );
    const dompurifyUri = this.panel.webview.asWebviewUri(
      vscode.Uri.joinPath(globalVariables.context.extensionUri, "out", "resource", "purify.min.js")
    );

    // Use a nonce to to only allow specific scripts to be run
    const nonce = this.getNonce();
    return `<!DOCTYPE html>
        <html lang="en">
          <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>ms-teams</title>
            <base href='${scriptBaseUri.toString()}' />
            <link href="${codiconsUri.toString()}" rel="stylesheet" />
          </head>
          <body>
            <div id="root"></div>
            <script>
              const vscode = acquireVsCodeApi();
              const panelType = '${panelType}';
            </script>
            <script nonce="${nonce}" type="module" src="${scriptUri.toString()}"></script>
            <script nonce="${nonce}" type="text/javascript" src="${dompurifyUri.toString()}"></script>
          </body>
        </html>`;
  }

  private getNonce() {
    let text = "";
    // eslint-disable-next-line
    const possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    for (let i = 0; i < 32; i++) {
      text += possible.charAt(Math.floor(Math.random() * possible.length));
    }
    return text;
  }

  private getWebviewPanelIconPath(panelType: PanelType) {
    if (panelType === PanelType.AccountHelp) {
      return vscode.Uri.file(
        path.join(globalVariables.context.extensionPath, "media/itp/m365icon.svg")
      );
    }
    return undefined;
  }

  public dispose() {
    const panelIndex = WebviewPanel.currentPanels.indexOf(this);
    if (TreatmentVariableValue.inProductDoc && this.panelType === PanelType.RespondToCardActions) {
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.InteractWithInProductDoc, {
        [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.InProductDoc,
        [TelemetryProperty.Interaction]: InProductGuideInteraction.Close,
        [TelemetryProperty.Identifier]: this.panelType,
      });
    } else if (this.panelType === PanelType.AccountHelp) {
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.InteractWithInProductDoc, {
        [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.AccountHelp,
        [TelemetryProperty.Interaction]: InProductGuideInteraction.Close,
        [TelemetryProperty.Identifier]: this.panelType,
      });
    }

    WebviewPanel.currentPanels.splice(panelIndex, 1);

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
