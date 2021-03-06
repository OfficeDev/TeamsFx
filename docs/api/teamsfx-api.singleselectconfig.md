<!-- Do not edit this file. It is automatically generated by API Documenter. -->

[Home](./index.md) &gt; [@microsoft/teamsfx-api](./teamsfx-api.md) &gt; [SingleSelectConfig](./teamsfx-api.singleselectconfig.md)

## SingleSelectConfig interface

single selection UI config

<b>Signature:</b>

```typescript
export interface SingleSelectConfig extends UIConfig<string> 
```
<b>Extends:</b> [UIConfig](./teamsfx-api.uiconfig.md)<!-- -->&lt;string&gt;

## Properties

|  Property | Type | Description |
|  --- | --- | --- |
|  [options](./teamsfx-api.singleselectconfig.options.md) | [StaticOptions](./teamsfx-api.staticoptions.md) | option array |
|  [returnObject?](./teamsfx-api.singleselectconfig.returnobject.md) | boolean | <i>(Optional)</i> This config only works for option items with <code>OptionItem[]</code> type. If <code>returnObject</code> is true, the answer value is an <code>OptionItem</code> object; otherwise, the answer value is the <code>id</code> string of the <code>OptionItem</code>. In case of option items with <code>string[]</code> type, whether <code>returnObject</code> is true or false, the returned answer value is always a string. |

