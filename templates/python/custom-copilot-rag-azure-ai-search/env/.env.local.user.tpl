# This file includes environment variables that will not be committed to git by default. You can set these environment variables in your CI/CD system for your project.

# If you're adding a secret value, add SECRET_ prefix to the name so Teams Toolkit can handle them properly
# Secrets. Keys prefixed with `SECRET_` will be masked in Teams Toolkit logs.
SECRET_BOT_PASSWORD=
{{#useOpenAI}}
{{#openAIKey}}
SECRET_OPENAI_API_KEY='{{{openAIKey}}}'
{{/openAIKey}}
{{^openAIKey}}
SECRET_OPENAI_API_KEY=
{{/openAIKey}}
{{/useOpenAI}}
{{#useAzureOpenAI}}
{{#azureOpenAIKey}}
SECRET_AZURE_OPENAI_API_KEY='{{{azureOpenAIKey}}}'
{{/azureOpenAIKey}}
{{^azureOpenAIKey}}
SECRET_AZURE_OPENAI_API_KEY=
{{/azureOpenAIKey}}
AZURE_OPENAI_MODEL_DEPLOYMENT_NAME=
{{#azureOpenAIEndpoint}}
AZURE_OPENAI_ENDPOINT='{{{azureOpenAIEndpoint}}}'
{{/azureOpenAIEndpoint}}
{{^azureOpenAIEndpoint}}
AZURE_OPENAI_ENDPOINT=
{{/azureOpenAIEndpoint}}
AZURE_OPENAI_EMBEDDING_DEPLOYMENT=
{{/useAzureOpenAI}}

SECRET_AZURE_SEARCH_KEY=
AZURE_SEARCH_ENDPOINT=