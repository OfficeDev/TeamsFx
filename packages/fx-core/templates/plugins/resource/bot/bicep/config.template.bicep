var botCurrentAppSettings = list('${provisionOutputs.botOutput.value.botWebAppResourceId}/config/appsettings', '2021-01-15').properties

module teamsFxBotConfig '{{PluginOutput.fx-resource-bot.Configuration.bot.ConfigPath}}' = {
  name: 'addTeamsFxBotConfiguration'
  params: {
    provisionParameters: provisionParameters
    provisionOutputs: provisionOutputs
    currentAppSettings: botCurrentAppSettings
  }
}
