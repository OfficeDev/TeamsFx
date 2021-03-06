<!-- Do not edit this file. It is automatically generated by API Documenter. -->

[Home](./index.md) &gt; [@microsoft/teamsfx](./teamsfx.md) &gt; [MessageBuilder](./teamsfx.messagebuilder.md) &gt; [attachHeroCard](./teamsfx.messagebuilder.attachherocard.md)

## MessageBuilder.attachHeroCard() method

Build a bot message activity attached with an hero card.

<b>Signature:</b>

```typescript
static attachHeroCard(title: string, images?: (CardImage | string)[], buttons?: (CardAction | string)[], other?: Partial<HeroCard>): Partial<Activity>;
```

## Parameters

|  Parameter | Type | Description |
|  --- | --- | --- |
|  title | string | The card title. |
|  images | (CardImage \| string)\[\] | Optional. The array of images to include on the card. |
|  buttons | (CardAction \| string)\[\] | Optional. The array of buttons to include on the card. Each <code>string</code> in the array is converted to an <code>imBack</code> button with a title and value set to the value of the string. |
|  other | Partial&lt;HeroCard&gt; | Optional. Any additional properties to include on the card. |

<b>Returns:</b>

Partial&lt;Activity&gt;

A bot message activity attached with a hero card.

## Example


```javascript
const message = MessageBuilder.attachHeroCard(
     'sample title',
     ['https://example.com/sample.jpg'],
     ['action']
);
```

