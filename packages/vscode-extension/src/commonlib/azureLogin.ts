/* eslint-disable @typescript-eslint/no-empty-function */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { TokenCredential } from "@azure/core-auth";
import { DeviceTokenCredentials, TokenCredentialsBase } from "@azure/ms-rest-nodeauth";
import {
  AzureAccountProvider,
  UserError,
  SubscriptionInfo,
  SingleSelectConfig,
  OptionItem,
  ok,
  ConfigFolderName,
} from "@microsoft/teamsfx-api";
import { ExtensionErrors } from "../error";
import { AzureAccount } from "./azure-account.api";
import { LoginFailureError } from "./codeFlowLogin";
import * as vscode from "vscode";
import * as identity from "@azure/identity";
import {
  initializing,
  loggedIn,
  loggedOut,
  loggingIn,
  signedIn,
  signedOut,
  signingIn,
  subscriptionInfoFile,
} from "./common/constant";
import { login, LoginStatus } from "./common/login";
import * as util from "util";
import * as ExtTelemetry from "../telemetry/telemetry";
import VsCodeLogInstance from "./log";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
  AccountType,
  TelemetryErrorType,
} from "../telemetry/extTelemetryEvents";
import { VS_CODE_UI } from "../extension";
import * as path from "path";
import * as fs from "fs-extra";
import * as commonUtils from "../debug/commonUtils";
import { environmentManager } from "@microsoft/teamsfx-core";
import { getSubscriptionInfoFromEnv } from "../utils/commonUtils";
import { getDefaultString, localize } from "../utils/localizeUtils";
import * as globalVariables from "../globalVariables";
import accountTreeViewProviderInstance from "../treeview/account/accountTreeViewProvider";

export class AzureAccountManager extends login implements AzureAccountProvider {
  private static instance: AzureAccountManager;
  private static subscriptionId: string | undefined;
  private static subscriptionName: string | undefined;
  private static tenantId: string | undefined;
  private static currentStatus: string | undefined;

  private static statusChange?: (
    status: string,
    token?: string,
    accountInfo?: Record<string, unknown>
  ) => Promise<void>;

  private constructor() {
    super();
    this.addStatusChangeEvent();
  }

  /**
   * Gets instance
   * @returns instance
   */
  public static getInstance(): AzureAccountManager {
    if (!AzureAccountManager.instance) {
      AzureAccountManager.instance = new AzureAccountManager();
    }

    return AzureAccountManager.instance;
  }

  /**
   * Update TeamsFx project subscription information.
   */
  public async updateSubscriptionInfo(): Promise<void> {
    if (AzureAccountManager.currentStatus === "LoggedIn") {
      const subscriptioninfo = await this.readSubscription();
      if (subscriptioninfo) {
        this.setSubscription(subscriptioninfo.subscriptionId);
      }
    }
  }

  /**
   * Async get ms-rest-* [credential](https://github.com/Azure/ms-rest-nodeauth/blob/master/lib/credentials/tokenCredentialsBase.ts)
   */
  async getAccountCredentialAsync(showDialog = true): Promise<TokenCredentialsBase | undefined> {
    if (this.isUserLogin()) {
      return this.doGetAccountCredentialAsync();
    }

    let cred;
    try {
      await this.login(showDialog);
      cred = await this.doGetAccountCredentialAsync();
    } catch (e) {
      ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.Login, e, {
        [TelemetryProperty.AccountType]: AccountType.Azure,
        [TelemetryProperty.Success]: TelemetrySuccess.No,
        [TelemetryProperty.UserId]: "",
        [TelemetryProperty.Internal]: "",
        [TelemetryProperty.ErrorType]:
          e instanceof UserError ? TelemetryErrorType.UserError : TelemetryErrorType.SystemError,
        [TelemetryProperty.ErrorCode]: `${e.source}.${e.name}`,
        [TelemetryProperty.ErrorMessage]: `${e.message}`,
      });
      throw e;
    }

    const userid = cred ? cred.clientId : "";
    const internal = cred
      ? (cred as DeviceTokenCredentials).username.endsWith("@microsoft.com")
      : false;
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Login, {
      [TelemetryProperty.AccountType]: AccountType.Azure,
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      [TelemetryProperty.UserId]: userid,
      [TelemetryProperty.Internal]: internal ? "true" : "false",
    });
    return cred;
  }

  /**
   * Async get identity [crendential](https://github.com/Azure/azure-sdk-for-js/blob/master/sdk/core/core-auth/src/tokenCredential.ts)
   */
  async getIdentityCredentialAsync(showDialog = true): Promise<TokenCredential | undefined> {
    if (this.isUserLogin()) {
      return this.doGetIdentityCredentialAsync();
    }
    await this.login(showDialog);
    return this.doGetIdentityCredentialAsync();
  }

  private async updateLoginStatus(): Promise<void> {
    if (this.isUserLogin() && AzureAccountManager.statusChange !== undefined) {
      const credential = await this.doGetAccountCredentialAsync();
      const accessToken = await credential?.getToken();
      const accountJson = await this.getJsonObject();
      await AzureAccountManager.statusChange("SignedIn", accessToken?.accessToken, accountJson);
    }
  }

  private isUserLogin(): boolean {
    const azureAccount: AzureAccount =
      vscode.extensions.getExtension<AzureAccount>("ms-vscode.azure-account")!.exports;
    return azureAccount.status === "LoggedIn";
  }

  private async login(showDialog: boolean): Promise<void> {
    if (showDialog) {
      const userConfirmation: boolean = await this.doesUserConfirmLogin();
      if (!userConfirmation) {
        // throw user cancel error
        throw new UserError(
          "Login",
          ExtensionErrors.UserCancel,
          getDefaultString("teamstoolkit.common.userCancel"),
          localize("teamstoolkit.common.userCancel")
        );
      }
    }

    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.LoginStart, {
      [TelemetryProperty.AccountType]: AccountType.Azure,
    });
    await vscode.commands.executeCommand("azure-account.login");
    const azureAccount = this.getAzureAccount();
    if (azureAccount.status != loggedIn) {
      throw new UserError(
        getDefaultString("teamstoolkit.codeFlowLogin.loginComponent"),
        getDefaultString("teamstoolkit.codeFlowLogin.loginTimeoutTitle"),
        getDefaultString("teamstoolkit.codeFlowLogin.loginTimeoutDescription"),
        localize("teamstoolkit.codeFlowLogin.loginTimeoutDescription")
      );
    }
  }

  private doGetAccountCredentialAsync(): Promise<TokenCredentialsBase | undefined> {
    if (this.isUserLogin()) {
      const azureAccount: AzureAccount =
        vscode.extensions.getExtension<AzureAccount>("ms-vscode.azure-account")!.exports;
      // Choose one tenant credential when users have multi tenants. (TODO, need to optize after UX design)
      // 1. When azure-account-extension has at least one subscription, return the first one credential.
      // 2. When azure-account-extension has no subscription and has at at least one session, return the first session credential.
      // 3. When azure-account-extension has no subscription and no session, return undefined.
      return new Promise(async (resolve, reject) => {
        await azureAccount.waitForSubscriptions();
        if (azureAccount.subscriptions.length > 0) {
          let credential2 = azureAccount.subscriptions[0].session.credentials2;
          if (AzureAccountManager.tenantId) {
            for (let i = 0; i < azureAccount.sessions.length; ++i) {
              const item = azureAccount.sessions[i];
              if (item.tenantId == AzureAccountManager.tenantId) {
                credential2 = item.credentials2;
                break;
              }
            }
          }
          // TODO - If the correct process is always selecting subs before other calls, throw error if selected subs not exist.
          resolve(credential2);
        } else if (azureAccount.sessions.length > 0) {
          resolve(azureAccount.sessions[0].credentials2);
        } else {
          reject(LoginFailureError());
        }
      });
    }
    return Promise.reject(LoginFailureError());
  }

  private doGetIdentityCredentialAsync(): Promise<TokenCredential | undefined> {
    if (this.isUserLogin()) {
      return new Promise(async (resolve) => {
        const tokenJson = await this.getJsonObject();
        const tenantId = (tokenJson as any).tid;
        const vsCredential = new identity.VisualStudioCodeCredential({ tenantId: tenantId });
        resolve(vsCredential);
      });
    }
    return Promise.reject(LoginFailureError());
  }

  private async doesUserConfirmLogin(): Promise<boolean> {
    const message = localize("teamstoolkit.azureLogin.message");
    const signin = localize("teamstoolkit.common.signin");
    const readMore = localize("teamstoolkit.common.readMore");
    let userSelected: string | undefined;
    do {
      userSelected = await vscode.window.showInformationMessage(
        message,
        { modal: true },
        signin,
        readMore
      );
      if (userSelected === readMore) {
        vscode.env.openExternal(
          vscode.Uri.parse(
            "https://docs.microsoft.com/en-us/azure/cost-management-billing/manage/create-subscription"
          )
        );
      }
    } while (userSelected === readMore);

    return Promise.resolve(userSelected === signin);
  }

  private async doesUserConfirmSignout(): Promise<boolean> {
    const accountInfo = (await this.getStatus()).accountInfo;
    const email = (accountInfo as any).upn ? (accountInfo as any).upn : (accountInfo as any).email;
    const confirm = localize("teamstoolkit.common.signout");
    const userSelected: string | undefined = await vscode.window.showInformationMessage(
      util.format(localize("teamstoolkit.common.signOutOf"), email),
      { modal: true },
      confirm
    );
    return Promise.resolve(userSelected === confirm);
  }

  async getJsonObject(showDialog = true): Promise<Record<string, unknown> | undefined> {
    const credential = await this.getAccountCredentialAsync(showDialog);
    const token = await credential?.getToken();
    if (token) {
      const array = token.accessToken.split(".");
      const buff = Buffer.from(array[1], "base64");
      return new Promise((resolve) => {
        resolve(JSON.parse(buff.toString("utf-8")));
      });
    } else {
      return new Promise((resolve) => {
        resolve(undefined);
      });
    }
  }

  /**
   * signout from Azure
   */
  async signout(): Promise<boolean> {
    const userConfirmation: boolean = await this.doesUserConfirmSignout();
    if (!userConfirmation) {
      // throw user cancel error
      throw new UserError(
        "SignOut",
        ExtensionErrors.UserCancel,
        getDefaultString("teamstoolkit.common.userCancel"),
        localize("teamstoolkit.common.userCancel")
      );
    }
    try {
      await vscode.commands.executeCommand("azure-account.logout");
      AzureAccountManager.tenantId = undefined;
      AzureAccountManager.subscriptionId = undefined;
      this.clearSubscription();
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.SignOut, {
        [TelemetryProperty.AccountType]: AccountType.Azure,
        [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      });
      return new Promise((resolve) => {
        resolve(true);
      });
    } catch (e) {
      VsCodeLogInstance.error("[Logout Azure] " + e.message);
      ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.SignOut, e, {
        [TelemetryProperty.AccountType]: AccountType.Azure,
        [TelemetryProperty.Success]: TelemetrySuccess.No,
        [TelemetryProperty.ErrorType]:
          e instanceof UserError ? TelemetryErrorType.UserError : TelemetryErrorType.SystemError,
        [TelemetryProperty.ErrorCode]: `${e.source}.${e.name}`,
        [TelemetryProperty.ErrorMessage]: `${e.message}`,
      });
      return Promise.resolve(false);
    }
  }

  /**
   * list all subscriptions
   */
  async listSubscriptions(): Promise<SubscriptionInfo[]> {
    await this.getAccountCredentialAsync();
    const azureAccount: AzureAccount =
      vscode.extensions.getExtension<AzureAccount>("ms-vscode.azure-account")!.exports;
    const arr: SubscriptionInfo[] = [];
    if (azureAccount.status === "LoggedIn") {
      if (azureAccount.subscriptions.length > 0) {
        for (let i = 0; i < azureAccount.subscriptions.length; ++i) {
          const item = azureAccount.subscriptions[i];
          arr.push({
            subscriptionId: item.subscription.subscriptionId!,
            subscriptionName: item.subscription.displayName!,
            tenantId: item.session.tenantId!,
          });
        }
      }
    }
    return arr;
  }

  /**
   * set tenantId and subscriptionId
   */
  async setSubscription(subscriptionId: string): Promise<void> {
    if (this.isUserLogin()) {
      const azureAccount: AzureAccount =
        vscode.extensions.getExtension<AzureAccount>("ms-vscode.azure-account")!.exports;
      for (let i = 0; i < azureAccount.subscriptions.length; ++i) {
        const item = azureAccount.subscriptions[i];
        if (item.subscription.subscriptionId == subscriptionId) {
          AzureAccountManager.tenantId = item.session.tenantId;
          AzureAccountManager.subscriptionId = subscriptionId;
          AzureAccountManager.subscriptionName = item.subscription.displayName;
          const subscriptionInfo = {
            subscriptionId: item.subscription.subscriptionId!,
            subscriptionName: item.subscription.displayName!,
            tenantId: item.session.tenantId,
          };
          await this.saveSubscription(subscriptionInfo);
          await accountTreeViewProviderInstance.azureAccountNode.setSubscription(subscriptionInfo);
          return;
        }
      }
    }
    throw new UserError(
      "Login",
      ExtensionErrors.UnknownSubscription,
      getDefaultString("teamstoolkit.azureLogin.unknownSubscription"),
      localize("teamstoolkit.azureLogin.unknownSubscription")
    );
  }

  getAzureAccount(): AzureAccount {
    const azureAccount: AzureAccount =
      vscode.extensions.getExtension<AzureAccount>("ms-vscode.azure-account")!.exports;
    return azureAccount;
  }

  async getStatus(): Promise<LoginStatus> {
    const azureAccount = this.getAzureAccount();
    if (this.isLegacyVersion()) {
      // add this to make sure Azure Account Extension has fully initialized
      // this will wait for login finish when version >= 0.10.0, so loggingIn status will be ignored
      await azureAccount.waitForSubscriptions();
    }
    if (azureAccount.status === loggedIn) {
      const credential = await this.doGetAccountCredentialAsync();
      const token = await credential?.getToken();
      const accountJson = await this.getJsonObject();
      return Promise.resolve({
        status: signedIn,
        token: token?.accessToken,
        accountInfo: accountJson,
      });
    } else if (azureAccount.status === loggingIn) {
      return Promise.resolve({ status: signingIn, token: undefined, accountInfo: undefined });
    } else {
      return Promise.resolve({ status: signedOut, token: undefined, accountInfo: undefined });
    }
  }

  async addStatusChangeEvent() {
    const azureAccount: AzureAccount =
      vscode.extensions.getExtension<AzureAccount>("ms-vscode.azure-account")!.exports;
    AzureAccountManager.currentStatus = azureAccount.status;
    await this.updateSubscriptionInfo();
    azureAccount.onStatusChanged(async (event) => {
      if (this.isLegacyVersion()) {
        if (AzureAccountManager.currentStatus === "Initializing") {
          AzureAccountManager.currentStatus = event;
          if (AzureAccountManager.currentStatus === "LoggedIn") {
            const subscriptioninfo = await this.readSubscription();
            if (subscriptioninfo) {
              this.setSubscription(subscriptioninfo.subscriptionId);
            }
          }
          return;
        }
        AzureAccountManager.currentStatus = event;
        if (event === loggedOut) {
          if (AzureAccountManager.statusChange !== undefined) {
            await AzureAccountManager.statusChange(signedOut, undefined, undefined);
          }
          await this.notifyStatus();
        } else if (event === loggedIn) {
          await this.updateLoginStatus();
          await this.notifyStatus();
        } else if (event === loggingIn) {
          await this.notifyStatus();
        }
      } else {
        if (AzureAccountManager.currentStatus === initializing) {
          if (event === loggedIn) {
            const subscriptioninfo = await this.readSubscription();
            if (subscriptioninfo) {
              this.setSubscription(subscriptioninfo.subscriptionId);
            }
            AzureAccountManager.currentStatus = event;
          } else if (event === loggedOut) {
            AzureAccountManager.currentStatus = event;
          }
          return;
        }
        AzureAccountManager.currentStatus = event;
        if (event === loggedOut) {
          if (AzureAccountManager.statusChange !== undefined) {
            await AzureAccountManager.statusChange(signedOut, undefined, undefined);
          }
          await this.notifyStatus();
        } else if (event === loggedIn) {
          await this.updateLoginStatus();
          await this.notifyStatus();
        } else if (event === loggingIn) {
          await this.notifyStatus();
        }
      }
    });
  }

  getAccountInfo(): Record<string, string> | undefined {
    const azureAccount = this.getAzureAccount();
    if (azureAccount.status === loggedIn) {
      return this.getJsonObject() as unknown as Record<string, string>;
    } else {
      return undefined;
    }
  }

  async getSelectedSubscription(triggerUI = false): Promise<SubscriptionInfo | undefined> {
    const azureAccount = this.getAzureAccount();
    if (triggerUI) {
      if (azureAccount.status !== loggedIn) {
        await this.login(true);
      }
      if (azureAccount.status === loggedIn && !AzureAccountManager.subscriptionId) {
        await this.selectSubscription();
      }
    } else {
      if (azureAccount.status === loggedIn && !AzureAccountManager.subscriptionId) {
        const subscriptionList = await this.listSubscriptions();
        if (subscriptionList && subscriptionList.length == 1) {
          await this.setSubscription(subscriptionList[0].subscriptionId);
        }
      }
    }
    if (azureAccount.status === loggedIn && AzureAccountManager.subscriptionId) {
      const selectedSub: SubscriptionInfo = {
        subscriptionId: AzureAccountManager.subscriptionId,
        tenantId: AzureAccountManager.tenantId!,
        subscriptionName: AzureAccountManager.subscriptionName ?? "",
      };
      return selectedSub;
    } else {
      return undefined;
    }
  }

  async selectSubscription(): Promise<void> {
    const subscriptionList = await this.listSubscriptions();
    if (!subscriptionList || subscriptionList.length == 0) {
      throw new UserError(
        getDefaultString("teamstoolkit.codeFlowLogin.loginComponent"),
        getDefaultString("teamstoolkit.azureLogin.noSubscriptionFound"),
        getDefaultString("teamstoolkit.azureLogin.failToFindSubscription"),
        localize("teamstoolkit.azureLogin.failToFindSubscription")
      );
    }
    if (subscriptionList && subscriptionList.length == 1) {
      await this.setSubscription(subscriptionList[0].subscriptionId);
    } else if (subscriptionList.length > 1) {
      const options: OptionItem[] = subscriptionList.map((sub) => {
        return {
          id: sub.subscriptionId,
          label: sub.subscriptionName,
          data: sub.tenantId,
        } as OptionItem;
      });
      const config: SingleSelectConfig = {
        name: localize("teamstoolkit.azureLogin.subscription"),
        title: localize("teamstoolkit.azureLogin.selectSubscription"),
        options: options,
      };
      const result = await VS_CODE_UI.selectOption(config);
      if (result.isErr()) {
        throw result.error;
      } else {
        const subId = result.value.result as string;
        await this.setSubscription(subId);
      }
    }
  }

  async saveSubscription(subscriptionInfo: SubscriptionInfo): Promise<void> {
    const subscriptionFilePath = await this.getSubscriptionInfoPath();
    if (!subscriptionFilePath) {
      return;
    } else {
      await fs.writeFile(subscriptionFilePath, JSON.stringify(subscriptionInfo, null, 4));
    }
  }

  async clearSubscription(): Promise<void> {
    const subscriptionFilePath = await this.getSubscriptionInfoPath();
    if (!subscriptionFilePath) {
      return;
    } else {
      await fs.writeFile(subscriptionFilePath, "");
    }
  }

  async readSubscription(): Promise<SubscriptionInfo | undefined> {
    const subscriptionFilePath = await this.getSubscriptionInfoPath();
    if (!subscriptionFilePath || !fs.existsSync(subscriptionFilePath)) {
      const solutionSubscriptionInfo = await getSubscriptionInfoFromEnv(
        environmentManager.getDefaultEnvName()
      );
      if (solutionSubscriptionInfo) {
        await this.saveSubscription(solutionSubscriptionInfo);
        return solutionSubscriptionInfo;
      }
      return undefined;
    } else {
      const content = (await fs.readFile(subscriptionFilePath)).toString();
      if (content.length == 0) {
        return undefined;
      }
      const subcriptionJson = JSON.parse(content);
      return {
        subscriptionId: subcriptionJson.subscriptionId,
        tenantId: subcriptionJson.tenantId,
        subscriptionName: subcriptionJson.subscriptionName,
      };
    }
  }

  async getSubscriptionInfoPath(): Promise<string | undefined> {
    if (globalVariables.workspaceUri) {
      const workspacePath: string = globalVariables.workspaceUri.fsPath;
      if (!globalVariables.isTeamsFxProject) {
        return undefined;
      }
      const configRoot = await commonUtils.getProjectRoot(workspacePath, `.${ConfigFolderName}`);
      const subscriptionFile = path.join(configRoot!, subscriptionInfoFile);
      return subscriptionFile;
    } else {
      return undefined;
    }
  }

  isLegacyVersion(): boolean {
    try {
      const version: string =
        vscode.extensions.getExtension<AzureAccount>("ms-vscode.azure-account")!.packageJSON
          .version;
      const versionDetail = version.split(".");
      return parseInt(versionDetail[0]) === 0 && parseInt(versionDetail[1]) < 10;
    } catch (e) {
      VsCodeLogInstance.error("[Get Azure extension] " + e.message);
      return false;
    }
  }
}

export default AzureAccountManager.getInstance();
