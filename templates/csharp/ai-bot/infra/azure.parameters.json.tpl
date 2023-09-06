{
    "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentParameters.json#",
    "contentVersion": "1.0.0.0",
    "parameters": {
      "resourceBaseName": {
        "value": "linkunfurling${{RESOURCE_SUFFIX}}"
      },
      "botAadAppClientId": {
        "value": "${{BOT_ID}}"
      },
      "botAadAppClientSecret": {
        "value": "${{SECRET_BOT_PASSWORD}}"
      },
      "openAIKey": {
      "value": "${{SECRET_OPENAI_API_KEY}}"
      },
      "azureOpenAIKey": {
        "value": "${{SECRET_AZURE_OPENAI_API_KEY}}"
      },
      "azureOpenAIEndpoint": {
        "value": "${{SECRET_AZURE_OPENAI_ENDPOINT}}"
      },
      "webAppSKU": {
        "value": "B1"
      },
      "botDisplayName": {
        "value": "{{appName}}"
      }
    }
  }