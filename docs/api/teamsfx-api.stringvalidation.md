<!-- Do not edit this file. It is automatically generated by API Documenter. -->

[Home](./index.md) &gt; [@microsoft/teamsfx-api](./teamsfx-api.md) &gt; [StringValidation](./teamsfx-api.stringvalidation.md)

## StringValidation interface

Validation for Strings

<b>Signature:</b>

```typescript
export interface StringValidation extends StaticValidation 
```
<b>Extends:</b> [StaticValidation](./teamsfx-api.staticvalidation.md)

## Properties

|  Property | Type | Description |
|  --- | --- | --- |
|  [endsWith?](./teamsfx-api.stringvalidation.endswith.md) | string | <i>(Optional)</i> A string instance is valid against this keyword if the string ends with the value of this keyword. |
|  [enum?](./teamsfx-api.stringvalidation.enum.md) | string\[\] | <i>(Optional)</i> A string instance validates successfully against this keyword if its value is equal to one of the elements in this keyword's array value. |
|  [equals?](./teamsfx-api.stringvalidation.equals.md) | string | <i>(Optional)</i> An instance validates successfully against this keyword if its value is equal to the value of the keyword. |
|  [includes?](./teamsfx-api.stringvalidation.includes.md) | string | <i>(Optional)</i> A string instance is valid against this keyword if the string contains the value of this keyword. |
|  [maxLength?](./teamsfx-api.stringvalidation.maxlength.md) | number | <i>(Optional)</i> A string instance is valid against this keyword if its length is less than, or equal to, the value of this keyword. |
|  [minLength?](./teamsfx-api.stringvalidation.minlength.md) | number | <i>(Optional)</i> A string instance is valid against this keyword if its length is greater than, or equal to, the value of this keyword. |
|  [pattern?](./teamsfx-api.stringvalidation.pattern.md) | string | <i>(Optional)</i> A string instance is considered valid if the regular expression matches the instance successfully. |
|  [startsWith?](./teamsfx-api.stringvalidation.startswith.md) | string | <i>(Optional)</i> A string instance is valid against this keyword if the string starts with the value of this keyword. |

