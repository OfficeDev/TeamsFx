// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Mutex } from "async-mutex";

import { globalStateGet, globalStateUpdate } from "@microsoft/teamsfx-core";
/* eslint-disable-next-line import/default */
import TelemetryReporter from "@vscode/extension-telemetry";

import { TelemetryEventCache, TelemetryProperty } from "./extTelemetryEvents";

const TelemetryCacheKey = "TelemetryEvents";
const CacheSize = 100;
const CacheLimit = 10;
const SuccessTimeSpan = 10 * 1000; // 10 seconds
const FlushCacheDelay = 3 * 1000; // 3 seconds

export class TelemetryCache {
  private cachedEvents = new Array<TelemetryEventCache>(CacheSize);
  private insertPos = 0;
  private endPos = 0;
  private timeout: NodeJS.Timeout | undefined;
  private mutex: Mutex;

  constructor(private reporter: TelemetryReporter) {
    this.mutex = new Mutex();
  }

  public async addEvent(event: TelemetryEventCache): Promise<void> {
    await this.mutex.runExclusive(() => {
      this.cachedEvents[this.insertPos] = event;
      this.insertPos = (this.insertPos + 1) % CacheSize;
      const size = (this.insertPos + CacheSize - this.endPos) % CacheSize;
      if (this.timeout) {
        clearTimeout(this.timeout);
        this.timeout = undefined;
      }
      if (size >= CacheLimit) {
        this.sendEventsInCache();
      } else {
        this.timeout = setTimeout(
          async () => await this.mutex.runExclusive(() => this.sendEventsInCache()),
          FlushCacheDelay
        );
      }
    });
  }

  public sendEventsInCache(): void {
    for (let i = this.endPos; i !== this.insertPos; i = (i + 1) % CacheSize) {
      const event = this.cachedEvents[i];
      if (event && event.sendTime === undefined) {
        const properties = {
          [TelemetryProperty.Timestamp]: event.occurTime.toISOString(),
          ...event.properties,
        };
        switch (event.type) {
          case "normal":
            this.reporter.sendTelemetryEvent(event.eventName, properties, event.measurements);
            break;
          case "error":
            this.reporter.sendTelemetryErrorEvent(event.eventName, properties, event.measurements);
            break;
        }
        event.sendTime = new Date();
      }
    }
    this.endPos = this.insertPos;
    this.timeout = undefined;
  }

  public async persistUncertainEventsToDiskAsync(
    deactivateEvent: TelemetryEventCache
  ): Promise<void> {
    const events: TelemetryEventCache[] = [];
    const now = new Date();
    for (let i = 0; i < CacheSize; i += 1) {
      const event = this.cachedEvents[i];
      if (!event) {
        continue;
      }
      if (
        !event.sendTime ||
        (event.sendTime && now.getTime() - event.sendTime.getTime() < SuccessTimeSpan)
      ) {
        events.push(event);
      }
    }
    events.push(deactivateEvent);
    const newValue = JSON.stringify(events);
    await globalStateUpdate(TelemetryCacheKey, newValue);
  }

  public async recoverUnsentEventsFromDiskAsync(): Promise<void> {
    const value = await globalStateGet(TelemetryCacheKey);
    if (value) {
      const events = JSON.parse(value);
      for (const event of events) {
        event.occurTime = new Date(event.occurTime);
        event.sendTime = undefined;
        this.addEvent(event);
      }
    }
  }
}
