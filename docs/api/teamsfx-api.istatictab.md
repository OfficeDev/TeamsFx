<!-- Do not edit this file. It is automatically generated by API Documenter. -->

[Home](./index.md) &gt; [@microsoft/teamsfx-api](./teamsfx-api.md) &gt; [IStaticTab](./teamsfx-api.istatictab.md)

## IStaticTab interface

<b>Signature:</b>

```typescript
export interface IStaticTab 
```

## Properties

|  Property | Type | Description |
|  --- | --- | --- |
|  [contentUrl?](./teamsfx-api.istatictab.contenturl.md) | string | <i>(Optional)</i> The url which points to the entity UI to be displayed in the Teams canvas. |
|  [context?](./teamsfx-api.istatictab.context.md) | ("personalTab" \| "channelTab")\[\] | <i>(Optional)</i> The set of contextItem scopes that a tab belong to |
|  [entityId](./teamsfx-api.istatictab.entityid.md) | string | A unique identifier for the entity which the tab displays. |
|  [name?](./teamsfx-api.istatictab.name.md) | string | <i>(Optional)</i> The display name of the tab. |
|  [objectId?](./teamsfx-api.istatictab.objectid.md) | string | <i>(Optional)</i> |
|  [scopes](./teamsfx-api.istatictab.scopes.md) | ("team" \| "personal")\[\] | Specifies whether the tab offers an experience in the context of a channel in a team, or an experience scoped to an individual user alone. These options are non-exclusive. Currently static tabs are only supported in the 'personal' scope. |
|  [searchUrl?](./teamsfx-api.istatictab.searchurl.md) | string | <i>(Optional)</i> The url to direct a user's search queries. |
|  [websiteUrl?](./teamsfx-api.istatictab.websiteurl.md) | string | <i>(Optional)</i> The url to point at if a user opts to view in a browser. |

