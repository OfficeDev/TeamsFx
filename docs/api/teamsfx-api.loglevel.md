<!-- Do not edit this file. It is automatically generated by API Documenter. -->

[Home](./index.md) &gt; [@microsoft/teamsfx-api](./teamsfx-api.md) &gt; [LogLevel](./teamsfx-api.loglevel.md)

## LogLevel enum

<b>Signature:</b>

```typescript
export declare enum LogLevel 
```

## Enumeration Members

|  Member | Value | Description |
|  --- | --- | --- |
|  Debug | <code>1</code> | For debugging and development. |
|  Error | <code>4</code> | For errors and exceptions that cannot be handled. These messages indicate a failure in the current operation or request, not an app-wide failure. |
|  Fatal | <code>5</code> | For failures that require immediate attention. Examples: data loss scenarios. |
|  Info | <code>2</code> | Tracks the general flow of the app. May have long-term value. |
|  Trace | <code>0</code> | Contain the most detailed messages. |
|  Warning | <code>3</code> | For abnormal or unexpected events. Typically includes errors or conditions that don't cause the app to fail. |

