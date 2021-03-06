<!-- Do not edit this file. It is automatically generated by API Documenter. -->

[Home](./index.md) &gt; [@microsoft/teamsfx-api](./teamsfx-api.md) &gt; [StringArrayValidation](./teamsfx-api.stringarrayvalidation.md)

## StringArrayValidation interface

Validation for String Arrays

<b>Signature:</b>

```typescript
export interface StringArrayValidation extends StaticValidation 
```
<b>Extends:</b> [StaticValidation](./teamsfx-api.staticvalidation.md)

## Properties

|  Property | Type | Description |
|  --- | --- | --- |
|  [contains?](./teamsfx-api.stringarrayvalidation.contains.md) | string | <i>(Optional)</i> An array instance is valid against "contains" if it contains the value of <code>contains</code> |
|  [containsAll?](./teamsfx-api.stringarrayvalidation.containsall.md) | string\[\] | <i>(Optional)</i> An array instance is valid against "containsAll" array if it contains all of the elements of <code>containsAll</code> array. |
|  [containsAny?](./teamsfx-api.stringarrayvalidation.containsany.md) | string\[\] | <i>(Optional)</i> An array instance is valid against "containsAny" array if it contains any one of the elements of <code>containsAny</code> array. |
|  [enum?](./teamsfx-api.stringarrayvalidation.enum.md) | string\[\] | <i>(Optional)</i> An array instance is valid against "enum" array if all of the elements of the array is contained in the <code>enum</code> array. |
|  [equals?](./teamsfx-api.stringarrayvalidation.equals.md) | string\[\] | <i>(Optional)</i> An instance validates successfully against this string array if they have the exactly the same elements. |
|  [maxItems?](./teamsfx-api.stringarrayvalidation.maxitems.md) | number | <i>(Optional)</i> The value of this keyword MUST be a non-negative integer. An array instance is valid against "maxItems" if its size is less than, or equal to, the value of this keyword. |
|  [minItems?](./teamsfx-api.stringarrayvalidation.minitems.md) | number | <i>(Optional)</i> The value of this keyword MUST be a non-negative integer. An array instance is valid against "minItems" if its size is greater than, or equal to, the value of this keyword. |
|  [uniqueItems?](./teamsfx-api.stringarrayvalidation.uniqueitems.md) | boolean | <i>(Optional)</i> If this keyword has boolean value false, the instance validates successfully. If it has boolean value true, the instance validates successfully if all of its elements are unique. |

