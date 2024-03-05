# This file includes environment variables that will not be committed to git by default. You can set these environment variables in your CI/CD system for your project.

# If you're adding a secret value, add SECRET_ prefix to the name so Teams Toolkit can handle them properly
# Secrets. Keys prefixed with `SECRET_` will be masked in Teams Toolkit logs.
SECRET_BOT_PASSWORD=
{{#userOpenAI}}
SECRET_OPENAI_API_KEY=<openai-api-key>
SECRET_OPENAI_MODEL_DEPLOYMENT_NAME=<openai-model-development-name>
{{/userOpenAI}}
{{#userAzureOpenAI}}
SECRET_AZURE_OPENAI_API_KEY=<azure-openai-api-key>
SECRET_AZURE_OPENAI_MODEL_DEPLOYMENT_NAME=<azure-openai-model-development-name>
SECRET_AZURE_OPENAI_ENDPOINT=<azure-openai-endpoint>
{{/userAzureOpenAI}}