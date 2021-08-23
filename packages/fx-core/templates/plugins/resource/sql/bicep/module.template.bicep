
module azureSqlProvision '{{PluginOutput.fx-resource-azure-sql.Modules.azureSqlProvision.Path}}' = {
  name: 'azureSqlProvision'
  params: {
    sqlServerName: azureSql_serverName
    sqlDatabaseName: azureSql_databaseName
    AADUser: azureSql_aadAdmin
    AADObjectId: azureSql_aadAdminObjectId
    AADTenantId: azureSql_tenantId
    administratorLogin: azureSql_admin
    administratorLoginPassword: azureSql_adminPassword
  }
}
