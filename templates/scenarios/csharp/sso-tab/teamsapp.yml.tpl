# yaml-language-server: $schema=https://developer.microsoft.com/json-schemas/teams-toolkit/teamsapp-yaml/1.0.0/yaml.schema.json
# Visit https://aka.ms/teamsfx-v5.0-guide for details on this file
# Visit https://aka.ms/teamsfx-actions for details on actions
version: 1.0.0

environmentFolderPath: ./env

# Triggered when 'teamsfx provision' is executed
provision:
  # Creates a new Azure Active Directory (AAD) app to authenticate users if
  # the environment variable that stores clientId is empty
  - uses: aadApp/create
    with:
      # Note: when you run aadApp/update, the AAD app name will be updated
      # based on the definition in manifest. If you don't want to change the
      # name, make sure the name in AAD manifest is the same with the name
      # defined here.
      name: {{appName}}-aad
      # If the value is false, the action will not generate client secret for you
      generateClientSecret: true
      # Authenticate users with a Microsoft work or school account in your
      # organization's Azure AD tenant (for example, single tenant).
      signInAudience: "AzureADMyOrg"
    # Write the information of created resources into environment file for the
    # specified environment variable(s).
    writeToEnvironmentFile:
      clientId: AAD_APP_CLIENT_ID
      # Environment variable that starts with `SECRET_` will be stored to the
      # .env.{envName}.user environment file
      clientSecret: SECRET_AAD_APP_CLIENT_SECRET
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

  # Apply the AAD manifest to an existing AAD app. Will use the object id in
  # manifest file to determine which AAD app to update.
  - uses: aadApp/update
    with:
      # Relative path to this file. Environment variables in manifest will
      # be replaced before apply to AAD app
      manifestPath: ./aad.manifest.json
      outputFilePath : ./build/aad.manifest.${{TEAMSFX_ENV}}.json

  - uses: teamsApp/validateManifest # Validate using manifest schema
    with:
      manifestPath: ./appPackage/manifest.json # Path to manifest template
  - uses: teamsApp/zipAppPackage # Build Teams app package with latest env value
    with:
      manifestPath: ./appPackage/manifest.json # Path to manifest template
      outputZipPath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
      outputJsonPath: ./appPackage/build/manifest.${{TEAMSFX_ENV}}.json
  - uses: teamsApp/validateAppPackage # Validate app package using validation rules
    with:
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip # Relative path to this file. This is the path for built zip file.
  - uses: teamsApp/update # Apply the Teams app manifest to an existing Teams app in Teams Developer Portal. Will use the app id in manifest file to determine which Teams app to update.
    with:
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip # Relative path to this file. This is the path for built zip file.

# Triggered when 'teamsfx deploy' is executed
deploy:
  - uses: cli/runDotnetCommand
    with:
      args: publish --configuration Release --runtime win-x86 --self-contained
  # Deploy your application to Azure App Service using the zip deploy feature.
  # For additional details, please refer to https://aka.ms/zip-deploy-to-app-services.
  - uses: azureAppService/zipDeploy
    with:
      # deploy base folder
      artifactFolder: bin/Release/net6.0/win-x86/publish
      # The resource id of the cloud resource to be deployed to. This key will be generated by arm/deploy action automatically. You can replace it with your existing Azure Resource id or add it to your environment variable file.
      resourceId: ${{TAB_AZURE_APP_SERVICE_RESOURCE_ID}}
