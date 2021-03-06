<!-- Do not edit this file. It is automatically generated by API Documenter. -->

[Home](./index.md) &gt; [@microsoft/teamsfx](./teamsfx.md) &gt; [BearerTokenAuthProvider](./teamsfx.bearertokenauthprovider.md) &gt; [AddAuthenticationInfo](./teamsfx.bearertokenauthprovider.addauthenticationinfo.md)

## BearerTokenAuthProvider.AddAuthenticationInfo() method

Adds authentication info to http requests

<b>Signature:</b>

```typescript
AddAuthenticationInfo(config: AxiosRequestConfig): Promise<AxiosRequestConfig>;
```

## Parameters

|  Parameter | Type | Description |
|  --- | --- | --- |
|  config | AxiosRequestConfig | Contains all the request information and can be updated to include extra authentication info. Refer https://axios-http.com/docs/req\_config for detailed document. |

<b>Returns:</b>

Promise&lt;AxiosRequestConfig&gt;

Updated axios request config.

## Exceptions

[AuthorizationInfoAlreadyExists](./teamsfx.errorcode.md) - when Authorization header already exists in request configuration.

