<!-- Do not edit this file. It is automatically generated by API Documenter. -->

[Home](./index.md) &gt; [@microsoft/teamsfx-api](./teamsfx-api.md) &gt; [v3](./teamsfx-api.v3.md) &gt; [ResourcePlugin](./teamsfx-api.v3.resourceplugin.md) &gt; [addResource](./teamsfx-api.v3.resourceplugin.addresource.md)

## v3.ResourcePlugin.addResource property

add resource is a new lifecycle task for resource plugin, which will do some extra work after project settings is updated, for example, APIM will scaffold the openapi folder with files

<b>Signature:</b>

```typescript
addResource?: (ctx: Context, inputs: InputsWithProjectPath) => Promise<Result<Void, FxError>>;
```
