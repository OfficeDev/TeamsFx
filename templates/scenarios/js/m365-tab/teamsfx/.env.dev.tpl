# Built-in environment variables
TEAMSFX_ENV=dev
TEAMS_APP_NAME={%appName%}
TEAMS_MANIFEST_PATH=./appPackage/manifest.template.json

# Updating AZURE_SUBSCRIPTION_ID or AZURE_RESOURCE_GROUP_NAME after provision may also require an update to RESOURCE_SUFFIX, because some services require a globally unique name across subscriptions/resource groups.
AZURE_SUBSCRIPTION_ID=
AZURE_RESOURCE_GROUP_NAME=
RESOURCE_SUFFIX=

# Generated during provision, you can also add your own variables. If you're adding a secret value, add SECRET_ prefix to the name so Teams Toolkit can handle them properly
TEAMS_APP_ID=
AAD_APP_CLIENT_ID=
SECRET_AAD_APP_CLIENT_SECRET=
AAD_APP_OBJECT_ID=
AAD_APP_OAUTH2_PERMISSION_ID=
AAD_APP_TENANT_ID=
AAD_APP_OAUTH_AUTHORITY_HOST=
AAD_APP_OAUTH_AUTHORITY=
TAB_AZURE_STORAGE_RESOURCE_ID=
TAB_ENDPOINT=
M365_CLIENT_ID=
M365_CLIENT_SECRET=
M365_TENANT_ID=
M365_OAUTH_AUTHORITY_HOST=