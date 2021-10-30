// Resources for Simple Auth
module simpleAuthProvision './simpleAuthProvision.result.v2.bicep' = {
  name: 'simpleAuthProvision'
  params: {
    provisionParameters: provisionParameters
  }
}

output simpleAuthOutput object = {
  teamsFxPluginId: 'fx-resource-simple-auth'
  endpoint: simpleAuthProvision.outputs.endpoint
  webAppResourceId: simpleAuthProvision.outputs.webAppResourceId
}
