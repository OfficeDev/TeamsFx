<!-- Do not edit this file. It is automatically generated by API Documenter. -->

[Home](./index.md) &gt; [@microsoft/teamsfx-api](./teamsfx-api.md) &gt; [InputResult](./teamsfx-api.inputresult.md) &gt; [type](./teamsfx-api.inputresult.type.md)

## InputResult.type property

`success`<!-- -->: the returned answer value is successfully collected when user click predefined confirm button/key, user will continue to answer the next question if available `skip`<!-- -->: the answer value is automatically selected when `skipSingleOption` is true for single/multiple selection list, user will continue to answer the next question if available `back`<!-- -->: the returned answer is undefined because user click the go-back button/key and will go back to re-answer the previous question in the question flow

<b>Signature:</b>

```typescript
type: "success" | "skip" | "back";
```
