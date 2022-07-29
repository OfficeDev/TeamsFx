// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { FolderName } from "@microsoft/teamsfx-core";

export enum Browser {
  chrome = "chrome",
  edge = "edge",
  default = "default",
}

export enum Hub {
  teams = "teams",
  outlook = "outlook",
  office = "office",
}

export class LaunchUrl {
  public static readonly teams: string =
    "https://teams.microsoft.com/l/app/${teamsAppId}?installAppPackage=true&webjoin=true&${account-hint}";
  public static readonly outlookTab: string =
    "https://outlook.office.com/host/${teamsAppInternalId}?${account-hint}";
  public static readonly outlookBot: string = "https://outlook.office.com/mail?${account-hint}";
  public static readonly officeTab: string =
    "https://www.office.com/m365apps/${teamsAppInternalId}?auth=2&${account-hint}";
}

export const teamsAppIdPlaceholder = "${teamsAppId}";
export const teamsAppInternalIdPlaceholder = "${teamsAppInternalId}";
export const accountHintPlaceholder = "${account-hint}";

export const serviceLogHintMessage = "The log of this task can be found in:";
export const openBrowserHintMessage =
  "WARN: Failed to open the browser, please copy the preview url and paste it into your browser.";
export const waitCtrlPlusC =
  "WARN: Closing browser will not terminate the preview process, please press Ctrl+C to terminate.";

export const frontendHostingPluginName = "fx-resource-frontend-hosting";
export const functionPluginName = "fx-resource-function";
export const botPluginName = "fx-resource-bot";
export const solutionPluginName = "solution";
export const appstudioPluginName = "fx-resource-appstudio";
export const spfxPluginName = "fx-resource-spfx";

export const teamsAppTenantIdConfigKey = "teamsAppTenantId";
export const remoteTeamsAppIdConfigKey = "teamsAppId";
export const botIdConfigKey = "botId";

export const frontendStartPattern = /Compiled|Failed/g;
export const backendStartPattern =
  /Worker process started and initialized|Host lock lease acquired by instance ID/g;
// From vscode $tsc-watch problem matcher: https://github.com/microsoft/vscode/blob/5a0ab56492d0c99f08028ca62ac3d59edb37f30f/extensions/typescript-language-features/package.json#L1085
export const tscWatchPattern =
  /^\s*(?:message TS6042:|\[?\D*\d{1,2}[:.]\d{1,2}[:.]\d{1,2}\D*(├\D*\d{1,2}\D+┤)?(?:\]| -)) (?:Compilation complete\.|Found \d+ errors?\.) Watching for file changes\./g;
// make a copy to prevent accidental change
export const backendWatchPattern = new RegExp(tscWatchPattern);
export const funcHostedBotWatchPattern = new RegExp(tscWatchPattern);
export const authStartPattern = /.*/g;
export const ngrokStartPattern = /started tunnel|failed to reconnect session/g;
export const botStartPattern = /listening|[nodemon] app crashed/g;
export const funcHostedBotStartPattern =
  /Worker process started and initialized|Host lock lease acquired by instance ID/g;
export const funcHostedBotAzuritePattern = /successfully listening/g;
export const gulpServePattern = /^.*Finished subtask 'reload'.*/g;

export const spfxInstallStartMessage = `executing 'npm install' under ${FolderName.SPFx} folder.`;
export const gulpCertTitle = "gulp trust-dev-cert";
export const gulpCertStartMessage = `executing 'gulp trust-dev-cert' under ${FolderName.SPFx} folder.`;
export const gulpServeTitle = "gulp serve";
export const gulpServeStartMessage = `executing 'gulp serve' under ${FolderName.SPFx} folder.`;
export const frontendInstallStartMessage = `executing 'npm install' under ${FolderName.Frontend} folder.`;
export const frontendStartStartMessage = `executing 'react-scripts start' under ${FolderName.Frontend} folder.`;
export const frontendStartStartMessageNext = `executing 'npm run dev:teamsfx' under ${FolderName.Frontend} folder.`;
export const authStartStartMessage = "starting auth service.";
export const backendInstallStartMessage = `executing 'npm install' under ${FolderName.Function} folder.`;
export const backendExtensionsInstallStartMessage =
  "installing Azure Functions binding extensions.";
export const backendStartStartMessage = `executing 'func start' under ${FolderName.Function} folder.`;
export const backendStartStartMessageNext = `executing 'npm run dev:teamsfx' under ${FolderName.Function} folder.`;
export const backendWatchStartMessage = `executing 'tsc --watch' under ${FolderName.Function} folder.`;
export const backendWatchStartMessageNext = `executing 'npm run watch:teamsfx' under ${FolderName.Function} folder.`;
export const botInstallStartMessage = `executing 'npm install' under ${FolderName.Bot} folder.`;
export const botStartStartMessage = "starting bot.";
export const botStartStartMessageNext = `executing 'npm run dev:teamsfx' under ${FolderName.Bot} folder.`;
export const botWatchStartMessage = `executing 'npm run watch:teamsfx' under ${FolderName.Bot} folder.`;
export const botAzuriteStartMessage = `executing 'npm run prepare-storage:teamsfx' under ${FolderName.Bot} folder.`;
export const ngrokStartStartMessage = `executing 'ngrok http' under ${FolderName.Bot} folder.`;

export const previewTitle = "preview";
export const previewStartMessage = "opening Teams web client.";
export const previewSPFxTitle = "spfx preview";
export const previewSPFxStartMessage = "opening SharePoint workbench.";

export const frontendLocalEnvPrefix = "FRONTEND_";
export const backendLocalEnvPrefix = "BACKEND_";
export const authLocalEnvPrefix = "AUTH_";
export const authServicePathEnvKey = "AUTH_SERVICE_PATH";
export const botLocalEnvPrefix = "BOT_";

export const automaticNpmInstallHintMessage =
  'Automatically installing packages required for your project. You can disable this by setting the global config "automatic-npm-install" to "off".';

export const doctorResult = {
  NodeNotFound: `Cannot find Node.js.`,
  NodeNotSupported: `Node.js (@CurrentVersion) is not in the supported version list (@SupportedVersions).`,
  NodeSuccess: `Supported Node.js version (@Version) is installed`,
  InstallNode: "Go to https://nodejs.org/about/releases/ to install Node.js (v16 is recommended).",
  BypassNode12: `To continue to preview using Node.js v12, please follow the link (@Link) to disable Node.js check with TeamsFx CLI.`,
  BypassNode12AndFunction: `To continue to preview using Node.js v12, please follow the link (@Link) to disable Node.js and Azure Functions Core Tools check with TeamsFx CLI. Also make sure you install the Azure Functions Core Tools v3. https://github.com/Azure/azure-functions-core-tools`,
  Node12MatchFunction:
    "If you have your own Azure Functions Core Tools installed, make sure it works with new Node.js version. See (https://docs.microsoft.com/azure/azure-functions/functions-versions#languages) for Azure Functions supported Node versions",
  SideLoadingDisabled:
    "Your M365 tenant admin hasn't enabled sideloading permission for your account. You can't install your app to Teams!",
  NotSignIn: "No M365 account login",
  SignInSuccess: `M365 Account (@account) is logged in and sideloading enabled`,
  SkipTrustingCert: "Skip trusting development certificate for localhost",
  HelpLink: `Please refer to @Link for more information.`,
};

export const installApp = {
  description:
    "To continue to preview your application in Outlook or Office.com, you need to install the app via Teams manually.",
  finish: "Once you have finished the installation, please come back and click 'Continue'.",
  guide: "Click 'Install in Teams' will pop up Teams web client for you to install the app.",
  installInTeams: "Install in Teams",
  installInTeamsDescription: "Pop up Teams web client for you to install the app.",
  continue: "Continue",
  continueDescription: "Continue to preview in Outlook or Office.",
  cancel: "Cancel",
  cancelDescription: "Stop preview.",
  installAppTitle: "Install app in Teams or continue to Outlook or Office",
  nonInteractive: {
    notInstalled:
      "We detected that you have not yet installed the app in Teams first, please run 'teamsfx preview %s --m365-host teams' to install app.",
    manifestChanges:
      "If you changed the manifest file, please run 'teamsfx preview %s --m365-host teams' to install app again.",
  },
  bot: {
    description:
      "To continue to preview your application in Outlook, you need to follow two steps:",
    guide1: "First, please click 'Install in Teams' to install the app in Teams.",
    guide2:
      "Second, please click 'Configure Outlook', sign in to the portal with the same Microsoft 365 account you used in Teams Toolkit. Click the 'Save' button in the portal to connect your bot to the Outlook channel.",
    remoteGuide2:
      "Second, please click 'Configure Outlook', sign in to the portal with the same Azure 365 account you used in Teams Toolkit. Select 'Outlook' in the portal and click the 'Apply' button to connect your bot to the Outlook channel.",
    finish: "Once you have finished the above two steps, please come back and click 'Continue'.",
    configureOutlook: "Configure Outlook",
    configureOutlookDescription:
      "Pop up Bot Framework Portal for you to connect your bot to Outlook channel.",
    remoteConfigureOutlookDescription:
      "Pop up Azure Portal for you to connect your bot to Outlook channel.",
  },
};

export const m365TenantHintMessage =
  "WARN: Please note that after you enrolled your developer tenant in Office 365 Target Release, it may take couple days for the enrollment to take effect. Please click https://aka.ms/teamsfx-m365-apps-prerequisites for more information about setting up dev environment for extending Teams apps across Microsoft 365.";
export const m365SwitchedMessage = "You have switched M365 tenant.";
