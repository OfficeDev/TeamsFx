/* eslint-disable @typescript-eslint/ban-ts-comment */
/* eslint-disable @typescript-eslint/ban-types */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import {
  FxError,
  M365TokenProvider,
  ok,
  Result,
  TokenRequest,
  UserError,
  err,
  LoginStatus,
  BasicLogin,
} from "@microsoft/teamsfx-api";
import { AccountInfo, LogLevel } from "@azure/msal-node";
import { ExtensionErrors } from "../error";
import { CodeFlowLogin, ConvertTokenToJson, UserCancelError } from "./codeFlowLogin";
import VsCodeLogInstance from "./log";
import * as vscode from "vscode";
import { CryptoCachePlugin } from "./cacheAccess";
import {
  loggedIn,
  loggingIn,
  m365CacheName,
  signedIn,
  signedOut,
  signingIn,
} from "./common/constant";
import * as util from "util";
import * as ExtTelemetry from "../telemetry/telemetry";
import {
  AccountType,
  TelemetryErrorType,
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/extTelemetryEvents";
import { getDefaultString, localize } from "../utils/localizeUtils";
import { AppStudioScopes } from "@microsoft/teamsfx-core";

const SERVER_PORT = 0;
const cachePlugin = new CryptoCachePlugin(m365CacheName);

const config = {
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

export class M365Login extends BasicLogin implements M365TokenProvider {
  private static instance: M365Login;
  private static codeFlowInstance: CodeFlowLogin;

  private constructor() {
    super();
    M365Login.codeFlowInstance = new CodeFlowLogin([], config, SERVER_PORT, m365CacheName);
  }

  /**
   * Gets instance
   * @returns instance
   */
  public static getInstance(): M365Login {
    if (!M365Login.instance) {
      M365Login.instance = new M365Login();
    }

    return M365Login.instance;
  }

  /**
   * Get team access token
   */
  async getAccessToken(tokenRequest: TokenRequest): Promise<Result<string, FxError>> {
    await M365Login.codeFlowInstance.reloadCache();
    let tokenRes: Result<string, FxError>;
    if (!M365Login.codeFlowInstance.account) {
      if (tokenRequest.showDialog === undefined || tokenRequest.showDialog) {
        const userConfirmation: boolean = await this.doesUserConfirmLogin();
        if (!userConfirmation) {
          ExtTelemetry.sendTelemetryErrorEvent(
            TelemetryEvent.Login,
            UserCancelError(getDefaultString("teamstoolkit.codeFlowLogin.loginComponent")),
            {
              [TelemetryProperty.AccountType]: AccountType.M365,
              [TelemetryProperty.Success]: TelemetrySuccess.No,
              [TelemetryProperty.UserId]: "",
              [TelemetryProperty.Internal]: "",
              [TelemetryProperty.ErrorType]: TelemetryErrorType.UserError,
              [TelemetryProperty.ErrorCode]: `${getDefaultString(
                "teamstoolkit.codeFlowLogin.loginComponent"
              )}.${ExtensionErrors.UserCancel}`,
              [TelemetryProperty.ErrorMessage]: `${getDefaultString(
                "teamstoolkit.common.userCancel"
              )}`,
            }
          );
          return err(
            UserCancelError(getDefaultString("teamstoolkit.codeFlowLogin.loginComponent"))
          );
        }
        M365Login.codeFlowInstance.status = loggingIn;
        this.notifyStatus(tokenRequest);
      }
      tokenRes = await M365Login.codeFlowInstance.getTokenByScopes(tokenRequest.scopes);
      await this.notifyStatus(tokenRequest);
    } else {
      tokenRes = await M365Login.codeFlowInstance.getTokenByScopes(tokenRequest.scopes);
    }

    if (tokenRes.isOk()) {
      return ok(tokenRes.value);
    } else {
      return tokenRes;
    }
  }

  async getJsonObject(
    tokenRequest: TokenRequest
  ): Promise<Result<Record<string, unknown>, FxError>> {
    const tokenRes = await this.getAccessToken(tokenRequest);
    if (tokenRes.isOk()) {
      const tokenJson = ConvertTokenToJson(tokenRes.value);
      return ok(tokenJson as any);
    } else {
      return err(tokenRes.error);
    }
  }

  async signout(): Promise<boolean> {
    await M365Login.codeFlowInstance.reloadCache();
    const userConfirmation = await this.doesUserConfirmSignout();
    if (!userConfirmation) {
      ExtTelemetry.sendTelemetryErrorEvent(
        TelemetryEvent.SignOut,
        UserCancelError(getDefaultString("teamstoolkit.commands.signOut.title")),
        {
          [TelemetryProperty.AccountType]: AccountType.M365,
          [TelemetryProperty.Success]: TelemetrySuccess.No,
          [TelemetryProperty.UserId]: "",
          [TelemetryProperty.Internal]: "",
          [TelemetryProperty.ErrorType]: TelemetryErrorType.UserError,
          [TelemetryProperty.ErrorCode]: `${getDefaultString(
            "teamstoolkit.codeFlowLogin.loginComponent"
          )}.${ExtensionErrors.UserCancel}`,
          [TelemetryProperty.ErrorMessage]: `${getDefaultString("teamstoolkit.common.userCancel")}`,
        }
      );
      throw UserCancelError(getDefaultString("teamstoolkit.commands.signOut.title"));
    }
    await M365Login.codeFlowInstance.logout();
    await this.notifyStatus({ scopes: AppStudioScopes });
    return true;
  }

  private async doesUserConfirmLogin(): Promise<boolean> {
    const message = localize("teamstoolkit.appStudioLogin.message");
    const signin = localize("teamstoolkit.common.signin");
    const createTestingTenant = localize("teamstoolkit.appStudioLogin.createM365TestingTenant");
    let userSelected: string | undefined;
    do {
      userSelected = await vscode.window.showInformationMessage(
        message,
        { modal: true },
        signin,
        createTestingTenant
      );
      if (userSelected === createTestingTenant) {
        vscode.env.openExternal(
          vscode.Uri.parse("https://developer.microsoft.com/en-us/microsoft-365/dev-program")
        );
      }
    } while (userSelected === createTestingTenant);
    return Promise.resolve(userSelected === signin);
  }

  private async doesUserConfirmSignout(): Promise<boolean> {
    const accountInfo = M365Login.codeFlowInstance.account;
    const email = accountInfo?.username;
    const confirm = localize("teamstoolkit.common.signout");
    const userSelected = await vscode.window.showInformationMessage(
      util.format(localize("teamstoolkit.common.signOutOf"), email),
      { modal: true },
      confirm
    );
    return Promise.resolve(userSelected === confirm);
  }

  async getStatus(tokenRequest: TokenRequest): Promise<Result<LoginStatus, FxError>> {
    await M365Login.codeFlowInstance.reloadCache();
    if (M365Login.codeFlowInstance.status === loggedIn) {
      const tokenRes = await M365Login.codeFlowInstance.getTokenByScopes(
        tokenRequest.scopes,
        false
      );
      if (tokenRes.isOk()) {
        const tokenJson = ConvertTokenToJson(tokenRes.value);
        return ok({ status: signedIn, token: tokenRes.value, accountInfo: tokenJson as any });
      } else {
        if (
          tokenRes.error.name !==
          getDefaultString("teamstoolkit.codeFlowLogin.checkOnlineFailTitle")
        ) {
          return ok({ status: signedOut, token: undefined, accountInfo: undefined });
        } else {
          return ok({
            status: signedIn,
            token: undefined,
            accountInfo: { upn: M365Login.codeFlowInstance.account?.username },
          });
        }
      }
    } else if (M365Login.codeFlowInstance.status === loggingIn) {
      return ok({ status: signingIn, token: undefined, accountInfo: undefined });
    } else {
      return ok({ status: signedOut, token: undefined, accountInfo: undefined });
    }
  }

  getCachedAccountInfo(): AccountInfo | undefined {
    return M365Login.codeFlowInstance.account;
  }
}

export default M365Login.getInstance();
