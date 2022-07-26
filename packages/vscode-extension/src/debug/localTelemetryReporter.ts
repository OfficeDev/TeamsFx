// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError } from "@microsoft/teamsfx-api";
import { LocalTelemetryReporter } from "@microsoft/teamsfx-core";
import { performance } from "perf_hooks";
import * as ExtTelemetry from "../telemetry/telemetry";
import {
  TelemetryEvent,
  TelemetryMeasurements,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/extTelemetryEvents";
import { getLocalDebugSession, getProjectComponents } from "./commonUtils";

function saveEventTime(eventName: string, time: number) {
  const session = getLocalDebugSession();
  // Assuming the event is only sent once in one local debug session,
  // because we only use the "high-level" events like debug-prerequisites, debug-precheck, etc..
  // And these events are indeed sent once.
  session.eventTimes[eventName] = time;
}

export const localTelemetryReporter = new LocalTelemetryReporter(
  {
    // Cannot directly refer to a global function because of import dependency cycle in ../telemetry/extTelemetry.ts.
    sendTelemetryEvent: (
      eventName: string,
      properties?: { [p: string]: string },
      measurements?: { [p: string]: number }
    ) => ExtTelemetry.sendTelemetryEvent(eventName, properties, measurements),

    sendTelemetryErrorEvent: (
      eventName: string,
      error: FxError,
      properties?: { [p: string]: string },
      measurements?: { [p: string]: number },
      errorProps?: string[]
    ) =>
      ExtTelemetry.sendTelemetryErrorEvent(eventName, error, properties, measurements, errorProps),
  },
  saveEventTime
);

export async function sendDebugAllStartEvent(additionalProperties: {
  [key: string]: string;
}): Promise<void> {
  const session = getLocalDebugSession();
  const components = await getProjectComponents();
  session.properties[TelemetryProperty.DebugProjectComponents] = components + "";
  Object.assign(session.properties, additionalProperties);

  const properties = Object.assign(
    { [TelemetryProperty.CorrelationId]: session.id },
    session.properties
  );
  localTelemetryReporter.sendTelemetryEvent(TelemetryEvent.DebugAllStart, properties);
}

export async function sendDebugAllEvent(
  error?: FxError,
  additionalProperties?: { [key: string]: string }
): Promise<void> {
  const session = getLocalDebugSession();
  const now = performance.now();

  let duration = -1;
  const startTime = session.eventTimes[TelemetryEvent.DebugAllStart];
  if (startTime !== undefined) {
    duration = (now - startTime) / 1000;
  }

  // Calculate the time between two events
  // Event must be only once in one debug session.
  function durationBetween(eventStart: string, eventEnd: string): number {
    const t0 = session.eventTimes[eventStart];
    const t1 = session.eventTimes[eventEnd];
    if (t0 !== undefined && t1 !== undefined) {
      return t1 - t0;
    } else {
      return -1;
    }
  }

  // Calculate the "time gap" in a local debug session.
  // In current local debug implementation, there is some time that vscode is in control and extension has no control.
  // For example, between "debug-precheck" (task finishes) and "debug-all" (browser starts), vscode is starting the services.
  // However, we don't know when the services successfully start because we use problem matcher to determine the service start or fail.
  // And vscode does not provide a callback for that.
  // Estimating from the current data, this "time gap" can be up to 1 minute, so not neglectable.
  const precheckGap =
    durationBetween(TelemetryEvent.DebugPrerequisites, TelemetryEvent.DebugPreCheckStart) / 1000;
  const precheckTime = session.eventTimes[TelemetryEvent.DebugPreCheck];
  const servicesGap = precheckTime === undefined ? -1 : (performance.now() - precheckTime) / 1000;

  const properties = {
    [TelemetryProperty.CorrelationId]: session.id,
    [TelemetryProperty.Success]: error === undefined ? TelemetrySuccess.Yes : TelemetrySuccess.No,
    ...session.properties,
    ...additionalProperties,
  };

  const measurements = {
    [LocalTelemetryReporter.PropertyDuration]: duration,
    [TelemetryMeasurements.DebugPrecheckGapDuration]: precheckGap,
    [TelemetryMeasurements.DebugServicesGapDuration]: servicesGap,
  };

  if (error === undefined) {
    localTelemetryReporter.sendTelemetryEvent(TelemetryEvent.DebugAll, properties, measurements);
  } else {
    localTelemetryReporter.sendTelemetryErrorEvent(
      TelemetryEvent.DebugAll,
      error,
      properties,
      measurements
    );
  }
}
