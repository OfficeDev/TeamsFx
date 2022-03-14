/* eslint-disable @typescript-eslint/ban-ts-comment */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";
import { SharepointTokenProvider, UserError } from "@microsoft/teamsfx-api";
import { LogLevel } from "@azure/msal-node";
import { ExtensionErrors } from "../error";
import { checkIsOnline, CodeFlowLogin, ConvertTokenToJson } from "./codeFlowLogin";
import { login, LoginStatus } from "./common/login";
import { loggedIn, loggingIn, signedIn, signedOut, signingIn } from "./common/constant";
import VsCodeLogInstance from "./log";
import * as vscode from "vscode";
import { CryptoCachePlugin } from "./cacheAccess";
import axios from "axios";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  AccountType,
  TelemetryErrorType,
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/extTelemetryEvents";
import { localize } from "../utils/localizeUtils";

const accountName = "appStudio";
const cachePlugin = new CryptoCachePlugin(accountName);

const graphScopes = ["https://graph.microsoft.com/User.ReadBasic.All"];
const graphCachePlugin = new CryptoCachePlugin(accountName);

const SERVER_PORT = 0;

/**
 * use msal to implement sharepoint login
 */
export class SharepointLogin extends login implements SharepointTokenProvider {
  private static instance: SharepointLogin;

  private static codeFlowInstance: CodeFlowLogin;
  private graphCodeFlowInstance: CodeFlowLogin;

  private static statusChange?: (
    status: string,
    token?: string,
    accountInfo?: Record<string, unknown>
  ) => Promise<void>;

  private constructor() {
    super();
    this.graphCodeFlowInstance = new CodeFlowLogin(
      graphScopes,
      this.getConfig(graphCachePlugin),
      SERVER_PORT,
      accountName
    );
  }

  /**
   * Gets instance
   * @returns instance
   */
  public static getInstance(): SharepointLogin {
    if (!SharepointLogin.instance) {
      SharepointLogin.instance = new SharepointLogin();
    }

    return SharepointLogin.instance;
  }

  /**
   * Get sharepoint access token
   */
  async getAccessToken(showDialog = true): Promise<string | undefined> {
    let isFirstLogin = false;
    if (!SharepointLogin.codeFlowInstance) {
      isFirstLogin = true;
      try {
        const scopes = await this.getScopes(showDialog);
        if (!scopes) {
          return undefined;
        }
        SharepointLogin.codeFlowInstance = new CodeFlowLogin(
          scopes,
          this.getConfig(cachePlugin),
          SERVER_PORT,
          accountName
        );
      } catch (error) {
        throw error;
      }
    }

    await SharepointLogin.codeFlowInstance.reloadCache();
    if (!isFirstLogin) {
      try {
        const scopes = await this.getScopes(showDialog);
        if (!scopes) {
          return undefined;
        }
        SharepointLogin.codeFlowInstance.scopes = scopes;
      } catch (error) {
        throw error;
      }
    }

    const accessToken = SharepointLogin.codeFlowInstance.getToken();
    return accessToken;
  }

  private async getSPTenant(accessToken: string): Promise<string> {
    const GRAPH_TENANT_ENDPT = "https://graph.microsoft.com/v1.0/sites/root?$select=webUrl";

    if (accessToken.length > 0) {
      const response = await axios.get(GRAPH_TENANT_ENDPT, {
        headers: { Authorization: `Bearer ${accessToken}` },
      });
      return response.data.webUrl;
    }
    return "";
  }

  private async getScopes(showDialog: boolean): Promise<string[] | undefined> {
    await this.graphCodeFlowInstance.reloadCache();
    try {
      const graphToken = await this.getGraphAccessToken(showDialog);
      if (!graphToken) {
        return undefined;
      }

      const tenant = await this.getSPTenant(graphToken!);
      if (!tenant) {
        return undefined;
      }
      const scopes = [`${tenant}/Sites.FullControl.All`];
      return scopes;
    } catch (error) {
      throw error;
    }
  }

  /**
   * Get graph access token
   */
  async getGraphAccessToken(showDialog = true): Promise<string | undefined> {
    await this.graphCodeFlowInstance.reloadCache();
    if (!this.graphCodeFlowInstance.account) {
      if (showDialog) {
        const userConfirmation: boolean = await this.doesUserConfirmLogin();
        if (!userConfirmation) {
          // throw user cancel error
          ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Login, {
            [TelemetryProperty.AccountType]: AccountType.M365,
            [TelemetryProperty.Success]: TelemetrySuccess.No,
            [TelemetryProperty.UserId]: "",
            [TelemetryProperty.Internal]: "",
            [TelemetryProperty.ErrorType]: TelemetryErrorType.UserError,
            [TelemetryProperty.ErrorCode]: `${localize(
              "teamstoolkit.codeFlowLogin.loginComponent"
            )}.${ExtensionErrors.UserCancel}`,
            [TelemetryProperty.ErrorMessage]: `${localize("teamstoolkit.common.userCancel")}`,
          });
          throw new UserError(
            ExtensionErrors.UserCancel,
            localize("teamstoolkit.common.userCancel"),
            "Login"
          );
        }
        this.graphCodeFlowInstance.status = loggingIn;
        this.notifyStatus();
      }
      try {
        const loginToken = await this.graphCodeFlowInstance.getToken();
        if (loginToken && SharepointLogin.statusChange !== undefined) {
          const tokenJson = await this.getJsonObject();
          await SharepointLogin.statusChange(signedIn, loginToken, tokenJson);
        }
        await this.notifyStatus();
        return loginToken;
      } catch (error) {
        this.notifyStatus();
        throw error;
      }
    }

    return this.graphCodeFlowInstance.getToken();
  }

  private async doesUserConfirmLogin(): Promise<boolean> {
    const message = localize("teamstoolkit.appStudioLogin.message");
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
          vscode.Uri.parse("https://developer.microsoft.com/en-us/microsoft-365/dev-program")
        );
      }
    } while (userSelected === readMore);
    return Promise.resolve(userSelected === signin);
  }

  async getJsonObject(showDialog = true): Promise<Record<string, unknown> | undefined> {
    const token = await this.getAccessToken();
    if (token) {
      const array = token.split(".");
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

  async getStatus(): Promise<LoginStatus> {
    await this.graphCodeFlowInstance.reloadCache();
    if (this.graphCodeFlowInstance.status === loggedIn) {
      const loginToken = await this.graphCodeFlowInstance.getToken(false);
      if (loginToken) {
        const tokenJson = ConvertTokenToJson(loginToken);
        return Promise.resolve({
          status: signedIn,
          token: loginToken,
          accountInfo: tokenJson as any,
        });
      } else {
        if (await checkIsOnline()) {
          return Promise.resolve({ status: signedOut, token: undefined, accountInfo: undefined });
        } else {
          return Promise.resolve({
            status: signedIn,
            token: undefined,
            accountInfo: { upn: this.graphCodeFlowInstance.account?.username },
          });
        }
      }
    } else if (this.graphCodeFlowInstance.status === loggingIn) {
      return Promise.resolve({ status: signingIn, token: undefined, accountInfo: undefined });
    } else {
      return Promise.resolve({ status: signedOut, token: undefined, accountInfo: undefined });
    }
  }

  private getConfig(cachePlugin: CryptoCachePlugin) {
    return {
      auth: {
        clientId: "7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0",
        authority: "https://login.microsoftonline.com/common",
      },
      system: {
        loggerOptions: {
          // @ts-ignore
          loggerCallback(loglevel, message, containsPii) {
            if (loglevel <= LogLevel.Error) {
              VsCodeLogInstance.error(message);
            }
          },
          piiLoggingEnabled: false,
          logLevel: LogLevel.Error,
        },
      },
      cache: {
        cachePlugin,
      },
    };
  }
}

export default SharepointLogin.getInstance();
