# yaml-language-server: $schema=https://developer.microsoft.com/json-schemas/teams-toolkit/teamsapp-yaml/1.0.0/yaml.schema.json
# Visit https://aka.ms/teamsfx-v5.0-guide for details on this file
# Visit https://aka.ms/teamsfx-actions for details on actions
version: 1.0.0

environmentFolderPath: ./env

# Triggered when 'teamsfx provision' is executed
provision:
  # Creates a Teams app
  - uses: teamsApp/create
    with:
      # Teams app name
      name: {{appName}}-${{TEAMSFX_ENV}}
    # Write the information of created resources into environment file for
    # the specified environment variable(s).
    writeToEnvironmentFile: 
      teamsAppId: TEAMS_APP_ID

  - uses: botAadApp/create # Creates a new or reuses an existing Azure Active Directory application for bot.
    with:
      name: {{appName}}-${{TEAMSFX_ENV}} # The Azure Active Directory application's display name
    writeToEnvironmentFile:
      botId: BOT_ID # The Azure Active Directory application's client id created for bot.
      botPassword: SECRET_BOT_PASSWORD # The Azure Active Directory application's client secret created for bot. 

  - uses: arm/deploy # Deploy given ARM templates parallelly.
    with:
      subscriptionId: ${{AZURE_SUBSCRIPTION_ID}} # The AZURE_SUBSCRIPTION_ID is a built-in environment variable. TeamsFx will ask you select one subscription if its value is empty. You're free to reference other environment varialbe here, but TeamsFx will not ask you to select subscription if it's empty in this case.
      resourceGroupName: ${{AZURE_RESOURCE_GROUP_NAME}} # The AZURE_RESOURCE_GROUP_NAME is a built-in environment variable. TeamsFx will ask you to select or create one resource group if its value is empty. You're free to reference other environment varialbe here, but TeamsFx will not ask you to select or create resource grouop if it's empty in this case.
      templates:
      - path: ./infra/azure.bicep
        parameters: ./infra/azure.parameters.json
        deploymentName: Create-resources-for-bot
      bicepCliVersion: v0.9.1 # Teams Toolkit will download this bicep CLI version from github for you, will use bicep CLI in PATH if you remove this config.

  # Validate using manifest schema
  - uses: teamsApp/validateManifest
    with:
      # Path to manifest template
      manifestPath: ./appPackage/manifest.json
  # Build Teams app package with latest env value
  - uses: teamsApp/zipAppPackage
    with:
      # Path to manifest template
      manifestPath: ./appPackage/manifest.json
      outputZipPath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
      outputJsonPath: ./appPackage/build/manifest.${{TEAMSFX_ENV}}.json
  # Validate app package using validation rules
  - uses: teamsApp/validateAppPackage
    with:
  # Build Teams app package with latest env value
  - uses: teamsApp/zipAppPackage
    with:
      # Path to manifest template
      manifestPath: ./appPackage/manifest.json
      outputZipPath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
      outputJsonPath: ./appPackage/build/manifest.${{TEAMSFX_ENV}}.json
  # Apply the Teams app manifest to an existing Teams app in Teams Developer Portal.
  # Will use the app id in manifest file to determine which Teams app to update.
  - uses: teamsApp/update
    with:
      # Relative path to this file. This is the path for built zip file.
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip

# Triggered when 'teamsfx deploy' is executed
deploy:
  - uses: cli/runNpmCommand # Run npm command
    with:
      args: install --production
  # Deploy your application to Azure App Service using the zip deploy feature.
  # For additional details, please refer to https://aka.ms/zip-deploy-to-app-services.
  - uses: azureAppService/zipDeploy
    with:
      # Deploy base folder
      artifactFolder: .
      # Can be changed to any ignore file location, leave blank will ignore nothing
      ignoreFile: .appserviceIgnore
      # The resource id of the cloud resource to be deployed to. This key will be generated by arm/deploy action automatically. You can replace it with your existing Azure Resource id or add it to your environment variable file.
      resourceId: ${{BOT_AZURE_APP_SERVICE_RESOURCE_ID}}

# Triggered when 'teamsfx publish' is executed
publish:
  # Validate using manifest schema
  - uses: teamsApp/validateManifest
    with:
      # Path to manifest template
      manifestPath: ./appPackage/manifest.json
  - uses: teamsApp/zipAppPackage
    with:
      # Path to manifest template
      manifestPath: ./appPackage/manifest.json
      outputZipPath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
      outputJsonPath: ./appPackage/build/manifest.${{TEAMSFX_ENV}}.json
  # Validate app package using validation rules
  - uses: teamsApp/validateAppPackage
    with:
      # Relative path to this file. This is the path for built zip file.
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
  # Apply the Teams app manifest to an existing Teams app in Teams Developer Portal.
  # Will use the app id in manifest file to determine which Teams app to update.
  - uses: teamsApp/update
    with:
      # Relative path to this file. This is the path for built zip file.
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
  # Publish the app to
  # Teams Admin Center (https://admin.teams.microsoft.com/policies/manage-apps)
  # for review and approval
  - uses: teamsApp/publishAppPackage
    with:
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
    # Write the information of created resources into environment file for
    # the specified environment variable(s).
    writeToEnvironmentFile:
      publishedAppId: TEAMS_APP_PUBLISHED_APP_ID