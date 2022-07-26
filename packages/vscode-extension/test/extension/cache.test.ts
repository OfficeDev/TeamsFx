// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as chai from "chai";
import * as sinon from "sinon";

import * as globalState from "@microsoft/teamsfx-core/build/common/globalState";
/* eslint-disable-next-line import/default */
import TelemetryReporter from "@vscode/extension-telemetry";

import { TelemetryCache } from "../../src/telemetry/cache";
import { TelemetryEventCache } from "../../src/telemetry/extTelemetryEvents";

describe("Telemetry Cache", () => {
  it("addEvent - excess limit number", async () => {
    const mockReporter = sinon.createStubInstance(TelemetryReporter);
    const cache = new TelemetryCache(mockReporter);
    for (let i = 0; i < 10; i += 1) {
      await cache.addEvent({
        type: "normal",
        eventName: "test",
        occurTime: new Date(),
        measurements: {},
      });
    }

    const calls = mockReporter.sendTelemetryEvent.getCalls();
    chai.expect(calls.length).to.equal(10);

    sinon.restore();
  });

  it("addEvent - excess timeout", async () => {
    const clock = sinon.useFakeTimers();
    const mockReporter = sinon.createStubInstance(TelemetryReporter);
    const cache = new TelemetryCache(mockReporter);
    await cache.addEvent({
      type: "normal",
      eventName: "test",
      occurTime: new Date(),
      measurements: {},
    });
    await clock.tickAsync(10 * 1000);
    await clock.nextAsync();
    clock.restore();
    const calls = mockReporter.sendTelemetryEvent.getCalls();
    chai.expect(calls.length).to.equal(1);
    sinon.restore();
  });

  it("persistUncertainEventsToDiskAsync", async () => {
    const mockReporter = sinon.createStubInstance(TelemetryReporter);
    const cache = new TelemetryCache(mockReporter);
    let state = "";
    const globalStateUpdateStub = sinon
      .stub(globalState, "globalStateUpdate")
      .callsFake(async (key, value) => (state = value));
    const clock = sinon.useFakeTimers();
    await cache.addEvent({
      type: "normal",
      eventName: "test1",
      occurTime: new clock.Date(),
    });
    await clock.tickAsync(4000);
    await cache.addEvent({
      type: "normal",
      eventName: "test2",
      occurTime: new clock.Date(),
    });
    const telemetryEvents: any[] = [
      {
        type: "normal",
        eventName: "test2",
        occurTime: new clock.Date(),
      },
    ];
    await clock.tickAsync(3000);
    telemetryEvents[0].sendTime = new clock.Date();
    await clock.tickAsync(9500);
    await clock.nextAsync();

    const time = new clock.Date();

    telemetryEvents.push({
      type: "normal",
      eventName: "deactivate",
      occurTime: time,
    });
    const expectedValue = JSON.stringify(telemetryEvents);

    await cache.persistUncertainEventsToDiskAsync({
      type: "normal",
      eventName: "deactivate",
      occurTime: time,
    });
    await clock.nextAsync();
    clock.restore();

    sinon.assert.calledOnce(globalStateUpdateStub);
    chai.expect(state).equals(expectedValue);

    sinon.restore();
  });

  it("recoverUnsentEventsFromDiskAsync", async () => {
    const mockReporter = sinon.createStubInstance(TelemetryReporter);
    const cache = new TelemetryCache(mockReporter);
    const time = new Date();
    const eventsInState = [
      {
        type: "normal",
        eventName: "test1",
        occurTime: time,
        sendTime: time,
      },
    ];
    const state = JSON.stringify(eventsInState);
    console.log(state);
    sinon.stub(globalState, "globalStateGet").callsFake(async (key) => {
      return state;
    });
    const events: TelemetryEventCache[] = [];
    sinon.stub(cache, "addEvent").callsFake(async (event) => {
      events.push(event);
    });

    await cache.recoverUnsentEventsFromDiskAsync();

    chai.expect(events.length).equals(1);
    chai.expect(events[0]["eventName"]).equals("test1");
    chai.expect(events[0]["occurTime"].toISOString()).equals(time.toISOString());
    chai.expect(events[0]["sendTime"]).equals(undefined);

    sinon.restore();
  });
});
