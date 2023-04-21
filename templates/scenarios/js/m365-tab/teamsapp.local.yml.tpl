# yaml-language-server: $schema=https://developer.microsoft.com/json-schemas/teams-toolkit/teamsapp-yaml/1.0.0/yaml.schema.json
# Visit https://aka.ms/teamsfx-v5.0-guide for details on this file
# Visit https://aka.ms/teamsfx-actions for details on actions
version: 1.0.0

provision:
  # Creates a new Azure Active Directory (AAD) app to authenticate users if
  # the environment variable that stores clientId is empty
  - uses: aadApp/create
    with:
      # Note: when you run aadApp/update, the AAD app name will be updated
      # based on the definition in manifest. If you don't want to change the
      # name, make sure the name in AAD manifest is the same with the name
      # defined here.
      name: {{appName}}
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
      
  # Set TAB_DOMAIN and TAB_ENDPOINT for local launch
  - uses: script 
    with:
      run:
        echo "::set-teamsfx-env TAB_DOMAIN=localhost:53000";
        echo "::set-teamsfx-env TAB_ENDPOINT=https://localhost:53000";

  # Apply the AAD manifest to an existing AAD app. Will use the object id in
  # manifest file to determine which AAD app to update.
  - uses: aadApp/update
    with:
      # Relative path to this file. Environment variables in manifest will
      # be replaced before apply to AAD app
      manifestPath: ./aad.manifest.json
      outputFilePath: ./build/aad.manifest.${{TEAMSFX_ENV}}.json

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

  # Extend your Teams app to Outlook and the Microsoft 365 app
  - uses: teamsApp/extendToM365
    with:
      # Relative path to the build app package.
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
    # Write the information of created resources into environment file for
    # the specified environment variable(s).
    writeToEnvironmentFile:
      titleId: M365_TITLE_ID
      appId: M365_APP_ID

deploy:
  - uses: devTool/install # Install development tool(s)
    with:
      devCert:
        trust: true
    writeToEnvironmentFile: # Write the information of installed development tool(s) into environment file for the specified environment variable(s).
      sslCertFile: SSL_CRT_FILE
      sslKeyFile: SSL_KEY_FILE

  - uses: cli/runNpmCommand # Run npm command
    with:
      workingDirectory: .
      args: install --no-audit

  - uses: file/createOrUpdateEnvironmentFile # Generate runtime environment variables
    with:
      target: ./.localConfigs
      envs:
        BROWSER: none
        HTTPS: true
        PORT: 53000
        SSL_CRT_FILE: ${{SSL_CRT_FILE}}
        SSL_KEY_FILE: ${{SSL_KEY_FILE}}
        REACT_APP_CLIENT_ID: ${{AAD_APP_CLIENT_ID}}
        REACT_APP_START_LOGIN_PAGE_URL: ${{TAB_ENDPOINT}}/auth-start.html
