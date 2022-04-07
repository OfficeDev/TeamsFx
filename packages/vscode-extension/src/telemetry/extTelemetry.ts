/* eslint-disable @typescript-eslint/no-namespace */
/* eslint-disable @typescript-eslint/no-unused-vars */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as vscode from "vscode";
import { VSCodeTelemetryReporter } from "../commonlib/telemetry";
import {
  TelemetryProperty,
  TelemetryComponentType,
  TelemetrySuccess,
  TelemetryEvent,
  TelemetryErrorType,
} from "./extTelemetryEvents";
import * as extensionPackage from "../../package.json";
import { FxError, Stage, UserError } from "@microsoft/teamsfx-api";
import { getIsExistingUser, isSPFxProject } from "../utils/commonUtils";
import { ext } from "../extensionVariables";

export namespace ExtTelemetry {
  export let reporter: VSCodeTelemetryReporter;
  export let hasSentTelemetry = false;
  /* eslint-disable prefer-const */
  export let isFromSample: boolean | undefined = undefined;
  export let settingsVersion: string | undefined = undefined;

  export function setHasSentTelemetry(eventName: string) {
    if (eventName === "query-expfeature") return;
    hasSentTelemetry = true;
  }

  export function addSharedProperty(name: string, value: string): void {
    reporter.addSharedProperty(name, value);
  }

  export class Reporter extends vscode.Disposable {
    constructor(ctx: vscode.ExtensionContext) {
      super(() => reporter.dispose());

      reporter = new VSCodeTelemetryReporter(
        extensionPackage.aiKey,
        extensionPackage.version,
        extensionPackage.name
      );
    }
  }

  export function stageToEvent(stage: Stage): string | undefined {
    /* debug telemetry event is not handling here */
    switch (stage) {
      case Stage.create:
        return TelemetryEvent.CreateProject;
      case Stage.init:
        return TelemetryEvent.InitProject;
      case Stage.update:
        return TelemetryEvent.AddResource;
      case Stage.provision:
        return TelemetryEvent.Provision;
      case Stage.deploy:
        return TelemetryEvent.Deploy;
      case Stage.publish:
        return TelemetryEvent.Publish;
      case Stage.createEnv:
        return TelemetryEvent.CreateNewEnvironment;
      case Stage.grantPermission:
        return TelemetryEvent.GrantPermission;
      default:
        return undefined;
    }
  }

  export function sendTelemetryEvent(
    eventName: string,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {
    setHasSentTelemetry(eventName);
    if (!properties) {
      properties = {};
    }

    if (TelemetryProperty.Component in properties === false) {
      properties[TelemetryProperty.Component] = TelemetryComponentType;
    }

    const isExistingUser = getIsExistingUser();
    properties[TelemetryProperty.IsExistingUser] = isExistingUser ? isExistingUser : "";

    const isSPFx = isSPFxProject(ext.workspaceUri.fsPath);
    properties[TelemetryProperty.IsSpfx] = isSPFx.toString();

    if (isFromSample != undefined) {
      properties![TelemetryProperty.IsFromSample] = isFromSample.toString();
    }
    if (settingsVersion !== undefined) {
      properties![TelemetryProperty.SettingsVersion] = settingsVersion.toString();
    }

    reporter.sendTelemetryEvent(eventName, properties, measurements);
  }

  export function sendTelemetryErrorEvent(
    eventName: string,
    error: FxError,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number },
    errorProps?: string[]
  ): void {
    if (!properties) {
      properties = {};
    }

    if (TelemetryProperty.Component in properties === false) {
      properties[TelemetryProperty.Component] = TelemetryComponentType;
    }

    const isExistingUser = getIsExistingUser();
    properties[TelemetryProperty.IsExistingUser] = isExistingUser ? isExistingUser : "";

    properties[TelemetryProperty.Success] = TelemetrySuccess.No;
    if (error instanceof UserError) {
      properties[TelemetryProperty.ErrorType] = TelemetryErrorType.UserError;
    } else {
      properties[TelemetryProperty.ErrorType] = TelemetryErrorType.SystemError;
    }

    properties[TelemetryProperty.ErrorCode] = `${error.source}.${error.name}`;
    properties[TelemetryProperty.ErrorMessage] = `${error.message}${
      error.stack ? "\nstack:\n" + error.stack : ""
    }`;

    const isSPFx = isSPFxProject(ext.workspaceUri.fsPath);
    properties[TelemetryProperty.IsSpfx] = isSPFx.toString();

    if (isFromSample != undefined) {
      properties![TelemetryProperty.IsFromSample] = isFromSample.toString();
    }
    if (settingsVersion !== undefined) {
      properties![TelemetryProperty.SettingsVersion] = settingsVersion.toString();
    }

    reporter.sendTelemetryErrorEvent(eventName, properties, measurements, errorProps);
  }

  export function sendTelemetryException(
    error: Error,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {
    if (!properties) {
      properties = {};
    }

    if (TelemetryProperty.Component in properties === false) {
      properties[TelemetryProperty.Component] = TelemetryComponentType;
    }

    const isExistingUser = getIsExistingUser();
    properties[TelemetryProperty.IsExistingUser] = isExistingUser ? isExistingUser : "";

    const isSPFx = isSPFxProject(ext.workspaceUri.fsPath);
    properties[TelemetryProperty.IsSpfx] = isSPFx.toString();

    if (isFromSample != undefined) {
      properties![TelemetryProperty.IsFromSample] = isFromSample.toString();
    }
    if (settingsVersion !== undefined) {
      properties![TelemetryProperty.SettingsVersion] = settingsVersion.toString();
    }

    reporter.sendTelemetryException(error, properties, measurements);
  }

  export async function dispose() {
    await reporter.dispose();
  }
}
