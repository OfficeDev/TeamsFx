// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as os from "os";
import { performance } from "perf_hooks";
import { SystemError, UserError } from "@microsoft/teamsfx-api";
import { TelemetryProperty } from "../../telemetry/extTelemetryEvents";
import * as ExtTelemetry from "../../telemetry/telemetry";
import { DepsTelemetry, DepsCheckerEvent, TelemetryMessurement } from "@microsoft/teamsfx-core";

export class VSCodeTelemetry implements DepsTelemetry {
  private readonly _telemetryComponentType = "extension:debug:envchecker";

  public sendEvent(
    eventName: DepsCheckerEvent,
    properties: { [key: string]: string } = {},
    timecost?: number
  ): void {
    this.addCommonProps(properties);
    const measurements: { [p: string]: number } = {};

    if (timecost) {
      measurements[TelemetryMessurement.completionTime] = timecost;
    }

    ExtTelemetry.sendTelemetryEvent(eventName, properties, measurements);
  }

  public async sendEventWithDuration(
    eventName: DepsCheckerEvent,
    action: () => Promise<void>
  ): Promise<void> {
    const start = performance.now();
    await action();
    // use seconds instead of milliseconds
    const timecost = Number(((performance.now() - start) / 1000).toFixed(2));
    this.sendEvent(eventName, {}, timecost);
  }

  public sendUserErrorEvent(eventName: DepsCheckerEvent, errorMessage: string): void {
    const error = new UserError(this._telemetryComponentType, eventName, errorMessage);
    ExtTelemetry.sendTelemetryErrorEvent(eventName, error, this.addCommonProps());
  }

  public sendSystemErrorEvent(
    eventName: DepsCheckerEvent,
    errorMessage: string,
    errorStack: string
  ): void {
    const error = new SystemError(
      this._telemetryComponentType,
      eventName,
      `errorMsg=${errorMessage},errorStack=${errorStack}`
    );
    error.stack = errorStack;
    ExtTelemetry.sendTelemetryErrorEvent(eventName, error, this.addCommonProps());
  }

  private addCommonProps(properties: { [key: string]: string } = {}): { [key: string]: string } {
    properties[TelemetryProperty.OSArch] = os.arch();
    properties[TelemetryProperty.OSRelease] = os.release();
    properties[TelemetryProperty.Component] = this._telemetryComponentType;
    return properties;
  }
}

export const vscodeTelemetry = new VSCodeTelemetry();
