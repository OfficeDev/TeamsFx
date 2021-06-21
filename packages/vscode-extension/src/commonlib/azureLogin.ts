/* eslint-disable @typescript-eslint/no-empty-function */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { TokenCredential } from "@azure/core-auth";
import { DeviceTokenCredentials, TokenCredentialsBase } from "@azure/ms-rest-nodeauth";
import { AzureAccountProvider, UserError, SubscriptionInfo } from "@microsoft/teamsfx-api";
import { ExtensionErrors } from "../error";
import { AzureAccount } from "./azure-account.api";
import { LoginFailureError } from "./codeFlowLogin";
import * as vscode from "vscode";
import * as identity from "@azure/identity";
import { loggedIn, loggedOut, loggingIn, signedIn, signedOut, signingIn } from "./common/constant";
import { login, LoginStatus } from "./common/login";
import * as StringResources from "../resources/Strings.json";
import * as util from "util";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import VsCodeLogInstance from "./log";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
  AccountType
} from "../telemetry/extTelemetryEvents";

export class AzureAccountManager extends login implements AzureAccountProvider {
  private static instance: AzureAccountManager;
  private static subscriptionId: string | undefined;
  private static tenantId: string | undefined;

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
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Login, {
        [TelemetryProperty.AccountType]: AccountType.Azure,
        [TelemetryProperty.Success]: TelemetrySuccess.No,
        [TelemetryProperty.UserId]: "",
        [TelemetryProperty.Internal]: "false"
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
      [TelemetryProperty.Internal]: internal ? "true" : "false"
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
          ExtensionErrors.UserCancel,
          StringResources.vsc.common.userCancel,
          "Login"
        );
      }
    }

    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.LoginStart, {
      [TelemetryProperty.AccountType]: AccountType.Azure
    });
    await vscode.commands.executeCommand("azure-account.login");
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
    const message = StringResources.vsc.azureLogin.message;
    const signin = StringResources.vsc.common.signin;
    const readMore = StringResources.vsc.common.readMore;
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
    const confirm = StringResources.vsc.common.signout;
    const userSelected: string | undefined = await vscode.window.showInformationMessage(
      util.format(StringResources.vsc.common.signOutOf, email),
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
        ExtensionErrors.UserCancel,
        StringResources.vsc.common.userCancel,
        "SignOut"
      );
    }
    try {
      await vscode.commands.executeCommand("azure-account.logout");
      AzureAccountManager.tenantId = undefined;
      AzureAccountManager.subscriptionId = undefined;
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.SignOut, {
        [TelemetryProperty.AccountType]: AccountType.Azure,
        [TelemetryProperty.Success]: TelemetrySuccess.Yes
      });
      return new Promise((resolve) => {
        resolve(true);
      });
    } catch (e) {
      VsCodeLogInstance.error(
        "[Logout Azure] " + e.message
      );
      ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.SignOut, e, {
        [TelemetryProperty.AccountType]: AccountType.Azure,
        [TelemetryProperty.Success]: TelemetrySuccess.No
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
          return;
        }
      }
    }
    throw new UserError(
      ExtensionErrors.UnknownSubscription,
      StringResources.vsc.azureLogin.unknownSubscription,
      "Login"
    );
  }

  getAzureAccount(): AzureAccount {
    const azureAccount: AzureAccount =
      vscode.extensions.getExtension<AzureAccount>("ms-vscode.azure-account")!.exports;
    return azureAccount;
  }

  async getStatus(): Promise<LoginStatus> {
    const azureAccount = this.getAzureAccount();
    if (azureAccount.status === loggedIn) {
      const credential = await this.doGetAccountCredentialAsync();
      const token = await credential?.getToken();
      const accountJson = await this.getJsonObject();
      return Promise.resolve({
        status: signedIn,
        token: token?.accessToken,
        accountInfo: accountJson
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
    // wait for Azure Account extension initialize
    await azureAccount.waitForSubscriptions();
    azureAccount.onStatusChanged(async (event) => {
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
    });
  }
}

export default AzureAccountManager.getInstance();
