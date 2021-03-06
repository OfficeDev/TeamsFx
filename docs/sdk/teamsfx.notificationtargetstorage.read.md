<!-- Do not edit this file. It is automatically generated by API Documenter. -->

[Home](./index.md) &gt; [@microsoft/teamsfx](./teamsfx.md) &gt; [NotificationTargetStorage](./teamsfx.notificationtargetstorage.md) &gt; [read](./teamsfx.notificationtargetstorage.read.md)

## NotificationTargetStorage.read() method

Read one notification target by its key.

<b>Signature:</b>

```typescript
read(key: string): Promise<{
        [key: string]: unknown;
    } | undefined>;
```

## Parameters

|  Parameter | Type | Description |
|  --- | --- | --- |
|  key | string | the key of a notification target. |

<b>Returns:</b>

Promise&lt;{ \[key: string\]: unknown; } \| undefined&gt;

- the notification target. Or undefined if not found.

