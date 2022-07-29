// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Colors,
  FxError,
  Inputs,
  InputTextConfig,
  InputTextResult,
  LoginStatus,
  LogLevel,
  MultiSelectConfig,
  MultiSelectResult,
  Result,
  SelectFileConfig,
  SelectFileResult,
  SelectFilesConfig,
  SelectFilesResult,
  SelectFolderConfig,
  SelectFolderResult,
  SingleSelectConfig,
  SingleSelectResult,
  SubscriptionInfo,
  TokenRequest,
  Void,
} from "@microsoft/teamsfx-api";
import {
  CancellationToken,
  NotificationType2,
  NotificationType3,
  NotificationType4,
  RequestType0,
  RequestType1,
  RequestType4,
} from "vscode-jsonrpc";

export enum Namespaces {
  /**
   * server-side
   */
  Server = "server",

  /**
   * client-side
   */
  Logger = "logger",
  Azure = "azure",
  Graph = "graph",
  AppStudio = "appStudio",
  SharePoint = "sharepoint",
  M365 = "m365",
  UserInteraction = "ui",
  Telemetry = "telemetry",
}

export type CustomizeFuncType = "LocalFunc" | "ValidateFunc" | "OnSelectionChangeFunc";

export interface CustomizeFuncRequestType {
  type: CustomizeFuncType;
  id: number;
}

/**
 * server-side request / notification types which are called from client to the server.
 */
export interface IServerConnection {
  createProjectRequest: (
    inputs: Inputs,
    token: CancellationToken
  ) => Promise<Result<string, FxError>>;
  localDebugRequest: (inputs: Inputs, token: CancellationToken) => Promise<Result<Void, FxError>>;
  provisionResourcesRequest: (
    inputs: Inputs,
    token: CancellationToken
  ) => Promise<Result<Void, FxError>>;
  deployArtifactsRequest: (
    inputs: Inputs,
    token: CancellationToken
  ) => Promise<Result<Void, FxError>>;
  buildArtifactsRequest: (
    inputs: Inputs,
    token: CancellationToken
  ) => Promise<Result<Void, FxError>>;
  publishApplicationRequest: (
    inputs: Inputs,
    token: CancellationToken
  ) => Promise<Result<Void, FxError>>;
  deployTeamsAppManifestRequest: (
    inputs: Inputs,
    token: CancellationToken
  ) => Promise<Result<any, FxError>>;

  customizeLocalFuncRequest: (
    funcId: number,
    inputs: Inputs,
    token: CancellationToken
  ) => Promise<Result<Void, FxError>>;
  customizeValidateFuncRequest: (
    funcId: number,
    answer: any,
    previousAnswers: Inputs | undefined,
    token: CancellationToken
  ) => Promise<Result<Void, FxError>>;
  customizeOnSelectionChangeFuncRequest: (
    funcId: number,
    currentSelectedIds: Set<string>,
    previousSelectedIds: Set<string>,
    token: CancellationToken
  ) => Promise<Result<Void, FxError>>;

  addSsoRequest: (inputs: Inputs, token: CancellationToken) => Promise<Result<any, FxError>>;
}

/**
 * notification types which are called from the server to the client.
 */
export const NotificationTypes = {
  [Namespaces.Logger]: {
    show: new NotificationType2<LogLevel, string>(`${Namespaces.Logger}/show`),
  },
  [Namespaces.Telemetry]: {
    sendTelemetryEvent: new NotificationType3<
      string,
      { [key: string]: string } | undefined,
      { [key: string]: number } | undefined
    >(`${Namespaces.Telemetry}/sendTelemetryEvent`),
    sendTelemetryErrorEvent: new NotificationType4<
      string,
      { [key: string]: string } | undefined,
      { [key: string]: number } | undefined,
      string[] | undefined
    >(`${Namespaces.Telemetry}/sendTelemetryErrorEvent`),
    sendTelemetryException: new NotificationType3<
      Error,
      { [key: string]: string } | undefined,
      { [key: string]: number } | undefined
    >(`${Namespaces.Telemetry}/sendTelemetryException`),
  },
};

/**
 * request types which are called from the server to the client.
 */
export const RequestTypes = {
  /**
   * token part.
   * TODO: do more investigation and update the parameters.
   * also, add other requests, such as signout.
   */
  [Namespaces.Azure]: {
    getAccountCredential: new RequestType0<
      Result<{ accessToken: string; tokenJsonString: string }, FxError>,
      Error
    >(`${Namespaces.Azure}/getAccountCredentialRequest`),
    getJsonObject: new RequestType0<Result<string, FxError>, Error>(
      `${Namespaces.Azure}/getJsonObjectRequest`
    ),
    listSubscriptions: new RequestType0<Result<SubscriptionInfo[], FxError>, Error>(
      `${Namespaces.Azure}/listSubscriptionsRequest`
    ),
    setSubscription: new RequestType1<string, Result<undefined, FxError>, Error>(
      `${Namespaces.Azure}/setSubscriptionRequest`
    ),
    getSelectedSubscription: new RequestType0<Result<SubscriptionInfo | undefined, FxError>, Error>(
      `${Namespaces.Azure}/getSelectedSubscriptionRequest`
    ),
  },
  [Namespaces.M365]: {
    getAccessToken: new RequestType1<TokenRequest, Result<string, FxError>, Error>(
      `${Namespaces.M365}/getAccessTokenRequest`
    ),
    getJsonObject: new RequestType1<TokenRequest, Result<string, FxError>, Error>(
      `${Namespaces.M365}/getJsonObjectRequest`
    ),
    getStatus: new RequestType1<TokenRequest, Result<LoginStatus, FxError>, Error>(
      `${Namespaces.M365}/getStatusRequest`
    ),
  },

  /**
   * user interaction
   * TODO: add the progress part.
   */
  [Namespaces.UserInteraction]: {
    selectOption: new RequestType1<SingleSelectConfig, Result<SingleSelectResult, FxError>, Error>(
      `${Namespaces.UserInteraction}/selectOptionRequest`
    ),
    selectOptions: new RequestType1<MultiSelectConfig, Result<MultiSelectResult, FxError>, Error>(
      `${Namespaces.UserInteraction}/selectOptionsRequest`
    ),
    inputText: new RequestType1<InputTextConfig, Result<InputTextResult, FxError>, Error>(
      `${Namespaces.UserInteraction}/inputTextRequest`
    ),
    openUrl: new RequestType1<string, Result<boolean, FxError>, Error>(
      `${Namespaces.UserInteraction}/openUrlRequest`
    ),
    selectFile: new RequestType1<SelectFileConfig, Result<SelectFileResult, FxError>, Error>(
      `${Namespaces.UserInteraction}/selectFileRequest`
    ),
    selectFiles: new RequestType1<SelectFilesConfig, Result<SelectFilesResult, FxError>, Error>(
      `${Namespaces.UserInteraction}/selectFilesRequest`
    ),
    selectFolder: new RequestType1<SelectFolderConfig, Result<SelectFolderResult, FxError>, Error>(
      `${Namespaces.UserInteraction}/selectFolderRequest`
    ),
    showMessage: new RequestType4<
      "info" | "warn" | "error",
      string | Array<{ content: string; color: Colors }>,
      boolean,
      string[],
      Result<string | undefined, FxError>,
      Error
    >(`${Namespaces.UserInteraction}/showMessageRequest`),
  },
};

export interface IServerFxError {
  errorType: "UserError" | "SystemError";
  source: string;
  name: string;
  message: string;
  stack?: string;
  innerError?: any;
  userData?: any;
  timestamp: Date;
  helpLink?: string;
  issueLink?: string;
  displayMessage?: string;
}
