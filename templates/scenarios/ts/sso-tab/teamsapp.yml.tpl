# yaml-language-server: $schema=https://developer.microsoft.com/json-schemas/teams-toolkit/teamsapp-yaml/1.0.0/yaml.schema.json
# Visit https://aka.ms/teamsfx-v5.0-guide for details on this file
# Visit https://aka.ms/teamsfx-actions for details on actions
version: 1.0.0

environmentFolderPath: ./env

# Triggered when 'teamsfx provision' is executed
provision:
  - uses: aadApp/create # Creates a new Azure Active Directory (AAD) app to authenticate users if the environment variable that stores clientId is empty
    with:
      name: {{appName}}-aad # Note: when you run aadApp/update, the AAD app name will be updated based on the definition in manifest. If you don't want to change the name, make sure the name in AAD manifest is the same with the name defined here.
      generateClientSecret: true # If the value is false, the action will not generate client secret for you
      signInAudience: "AzureADMyOrg" # Authenticate users with a Microsoft work or school account in your organization's Azure AD tenant (for example, single tenant).
    writeToEnvironmentFile: # Write the information of created resources into environment file for the specified environment variable(s).
      clientId: AAD_APP_CLIENT_ID
      clientSecret: SECRET_AAD_APP_CLIENT_SECRET # Environment variable that starts with `SECRET_` will be stored to the .env.{envName}.user environment file
      objectId: AAD_APP_OBJECT_ID
      tenantId: AAD_APP_TENANT_ID
      authority: AAD_APP_OAUTH_AUTHORITY
      authorityHost: AAD_APP_OAUTH_AUTHORITY_HOST

  - uses: teamsApp/create # Creates a Teams app
    with:
      name: {{appName}}-${{TEAMSFX_ENV}} # Teams app name
    writeToEnvironmentFile: # Write the information of created resources into environment file for the specified environment variable(s).
      teamsAppId: TEAMS_APP_ID

  - uses: arm/deploy # Deploy given ARM templates parallelly.
    with:
      subscriptionId: ${{AZURE_SUBSCRIPTION_ID}} # The AZURE_SUBSCRIPTION_ID is a built-in environment variable. TeamsFx will ask you select one subscription if its value is empty. You're free to reference other environment varialbe here, but TeamsFx will not ask you to select subscription if it's empty in this case.
      resourceGroupName: ${{AZURE_RESOURCE_GROUP_NAME}} # The AZURE_RESOURCE_GROUP_NAME is a built-in environment variable. TeamsFx will ask you to select or create one resource group if its value is empty. You're free to reference other environment varialbe here, but TeamsFx will not ask you to select or create resource grouop if it's empty in this case.
      templates:
       - path: ./infra/azure.bicep # Relative path to this file
         parameters: ./infra/azure.parameters.json # Relative path to this file. Placeholders will be replaced with corresponding environment variable before ARM deployment.
         deploymentName: Create-resources-for-tab # Required when deploy ARM template
      bicepCliVersion: v0.9.1 # Teams Toolkit will download this bicep CLI version from github for you, will use bicep CLI in PATH if you remove this config.

  - uses: azureStorage/enableStaticWebsite
    with:
      storageResourceId: ${{TAB_AZURE_STORAGE_RESOURCE_ID}}
      indexPage: index.html
      errorPage: error.html

  - uses: aadApp/update # Apply the AAD manifest to an existing AAD app. Will use the object id in manifest file to determine which AAD app to update.
    with:
      manifestPath: ./aad.manifest.json # Relative path to this file. Environment variables in manifest will be replaced before apply to AAD app
      outputFilePath : ./build/aad.manifest.${{TEAMSFX_ENV}}.json

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
      # Relative path to this file. This is the path for built zip file.
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
  # Apply the Teams app manifest to an existing Teams app in Teams Developer Portal.
  # Will use the app id in manifest file to determine which Teams app to update.
  - uses: teamsApp/update
    with:
      # Relative path to this file. This is the path for built zip file.
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip

# Triggered when 'teamsfx deploy' is executed
deploy:
  - uses: cli/runNpmCommand # Run npm command
    name: install dependencies
    with:
      args: install
  - uses: cli/runNpmCommand # Run npm command
    name: build app
    env:
      REACT_APP_CLIENT_ID: ${{AAD_APP_CLIENT_ID}}
      REACT_APP_START_LOGIN_PAGE_URL: ${{TAB_ENDPOINT}}/auth-start.html
    with:
      args: run build --if-present
  # Deploy bits to Azure Storage Static Website
  - uses: azureStorage/deploy
    with:
      # Deploy base folder. This folder includes manifest files for AAD app and Teams app that should be ignored using the ignoreFile.
      artifactFolder: build
      ignoreFile: .storageignore
      # The resource id of the cloud resource to be deployed to. This key will be generated by arm/deploy action automatically. You can replace it with your existing Azure Resource id or add it to your environment variable file.
      resourceId: ${{TAB_AZURE_STORAGE_RESOURCE_ID}}

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