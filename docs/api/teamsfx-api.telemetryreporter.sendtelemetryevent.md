<!-- Do not edit this file. It is automatically generated by API Documenter. -->

[Home](./index.md) &gt; [@microsoft/teamsfx-api](./teamsfx-api.md) &gt; [TelemetryReporter](./teamsfx-api.telemetryreporter.md) &gt; [sendTelemetryEvent](./teamsfx-api.telemetryreporter.sendtelemetryevent.md)

## TelemetryReporter.sendTelemetryEvent() method

Send general events to App Insights

<b>Signature:</b>

```typescript
sendTelemetryEvent(eventName: string, properties?: {
        [key: string]: string;
    }, measurements?: {
        [key: string]: number;
    }): void;
```

## Parameters

|  Parameter | Type | Description |
|  --- | --- | --- |
|  eventName | string | Event name. Max length: 512 characters. To allow proper grouping and useful metrics, restrict your application so that it generates a small number of separate event names. |
|  properties | { \[key: string\]: string; } | Name-value collection of custom properties. Max key length: 150, Max value length: 8192. this collection is used to extend standard telemetry with the custom dimensions. |
|  measurements | { \[key: string\]: number; } | Collection of custom measurements. Use this collection to report named measurement associated with the telemetry item. |

<b>Returns:</b>

void

