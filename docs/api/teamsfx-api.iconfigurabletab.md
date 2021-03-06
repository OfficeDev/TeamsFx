<!-- Do not edit this file. It is automatically generated by API Documenter. -->

[Home](./index.md) &gt; [@microsoft/teamsfx-api](./teamsfx-api.md) &gt; [IConfigurableTab](./teamsfx-api.iconfigurabletab.md)

## IConfigurableTab interface

<b>Signature:</b>

```typescript
export interface IConfigurableTab 
```

## Properties

|  Property | Type | Description |
|  --- | --- | --- |
|  [canUpdateConfiguration?](./teamsfx-api.iconfigurabletab.canupdateconfiguration.md) | boolean | <i>(Optional)</i> A value indicating whether an instance of the tab's configuration can be updated by the user after creation. |
|  [configurationUrl](./teamsfx-api.iconfigurabletab.configurationurl.md) | string | The url to use when configuring the tab. |
|  [context?](./teamsfx-api.iconfigurabletab.context.md) | ("channelTab" \| "privateChatTab" \| "meetingChatTab" \| "meetingDetailsTab" \| "meetingSidePanel" \| "meetingStage")\[\] | <i>(Optional)</i> The set of contextItem scopes that a tab belong to |
|  [objectId?](./teamsfx-api.iconfigurabletab.objectid.md) | string | <i>(Optional)</i> |
|  [scopes](./teamsfx-api.iconfigurabletab.scopes.md) | ("team" \| "groupchat")\[\] | Specifies whether the tab offers an experience in the context of a channel in a team, in a 1:1 or group chat, or in an experience scoped to an individual user alone. These options are non-exclusive. Currently, configurable tabs are only supported in the teams and groupchats scopes. |
|  [sharePointPreviewImage?](./teamsfx-api.iconfigurabletab.sharepointpreviewimage.md) | string | <i>(Optional)</i> A relative file path to a tab preview image for use in SharePoint. Size 1024x768. |
|  [supportedSharePointHosts?](./teamsfx-api.iconfigurabletab.supportedsharepointhosts.md) | ("sharePointFullPage" \| "sharePointWebPart")\[\] | <i>(Optional)</i> Defines how your tab will be made available in SharePoint. |

