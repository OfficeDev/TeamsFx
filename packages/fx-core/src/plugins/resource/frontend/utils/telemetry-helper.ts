// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  DependentPluginInfo,
  FrontendPluginInfo,
  TelemetryEvent,
  TelemetryKey,
  TelemetryValue,
} from "../constants";
import { PluginContext, SystemError, UserError } from "@microsoft/teamsfx-api";

export class telemetryHelper {
  private static fillCommonProperty(
    ctx: PluginContext,
    properties: { [key: string]: string }
  ): void {
    properties[TelemetryKey.Component] = FrontendPluginInfo.PluginName;
    properties[TelemetryKey.AppId] =
      (ctx.configOfOtherPlugins
        .get(DependentPluginInfo.SolutionPluginName)
        ?.get(DependentPluginInfo.RemoteTeamsAppId) as string) || "";
  }

  static sendStartEvent(
    ctx: PluginContext,
    eventName: string,
    properties: { [key: string]: string } = {},
    measurements: { [key: string]: number } = {}
  ): void {
    telemetryHelper.fillCommonProperty(ctx, properties);

    ctx.telemetryReporter?.sendTelemetryEvent(
      eventName + TelemetryEvent.startSuffix,
      properties,
      measurements
    );
  }

  static sendSuccessEvent(
    ctx: PluginContext,
    eventName: string,
    properties: { [key: string]: string } = {},
    measurements: { [key: string]: number } = {}
  ): void {
    telemetryHelper.fillCommonProperty(ctx, properties);
    properties[TelemetryKey.Success] = TelemetryValue.Success;

    ctx.telemetryReporter?.sendTelemetryEvent(eventName, properties, measurements);
  }

  static sendErrorEvent(
    ctx: PluginContext,
    eventName: string,
    e: SystemError | UserError,
    properties: { [key: string]: string } = {},
    measurements: { [key: string]: number } = {}
  ): void {
    telemetryHelper.fillCommonProperty(ctx, properties);
    properties[TelemetryKey.Success] = TelemetryValue.Fail;

    if (e instanceof SystemError) {
      properties[TelemetryKey.ErrorType] = TelemetryValue.SystemError;
    } else if (e instanceof UserError) {
      properties[TelemetryKey.ErrorType] = TelemetryValue.UserError;
    }
    properties[TelemetryKey.ErrorMessage] = e.message;
    properties[TelemetryKey.ErrorCode] = e.name;

    ctx.telemetryReporter?.sendTelemetryEvent(eventName, properties, measurements);
  }
}
