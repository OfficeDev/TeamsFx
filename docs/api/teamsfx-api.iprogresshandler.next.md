<!-- Do not edit this file. It is automatically generated by API Documenter. -->

[Home](./index.md) &gt; [@microsoft/teamsfx-api](./teamsfx-api.md) &gt; [IProgressHandler](./teamsfx-api.iprogresshandler.md) &gt; [next](./teamsfx-api.iprogresshandler.next.md)

## IProgressHandler.next property

Update the progress bar's message. After calling it, the progress bar will be seen to users with $<!-- -->{<!-- -->currentStep<!-- -->}<!-- -->++ and $<!-- -->{<!-- -->detail<!-- -->} = detail. This func must be called after calling start().

<b>Signature:</b>

```typescript
next: (detail?: string) => Promise<void>;
```
