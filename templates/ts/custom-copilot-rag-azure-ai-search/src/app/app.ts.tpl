import { MemoryStorage } from "botbuilder";
import * as path from "path";
import config from "../config";

// See https://aka.ms/teams-ai-library to learn more about the Teams AI library.
import { Application, ActionPlanner, OpenAIModel, PromptManager, TurnState } from "@microsoft/teams-ai";
import { AzureAISearchDataSource } from "./AzureAISearchDataSource";

// Create AI components
const model = new OpenAIModel({
  {{#useOpenAI}}
  apiKey: config.openAIKey,
  defaultModel: config.openAIModelName,
  {{/useOpenAI}}
  {{#useAzureOpenAI}}
  azureApiKey: config.azureOpenAIKey,
  azureDefaultDeployment: config.azureOpenAIDeploymentName,
  azureEndpoint: config.azureOpenAIEndpoint,
  {{/useAzureOpenAI}}

  useSystemMessages: true,
  logRequests: true,
});
const prompts = new PromptManager({
  promptsFolder: path.join(__dirname, "../prompts"),
});
const planner = new ActionPlanner<TurnState>({
  model,
  prompts,
  defaultPrompt: "chat",
});

// Register your data source with planner
planner.prompts.addDataSource(
  new AzureAISearchDataSource({
    name: 'azure-ai-search',
    indexName: 'my-documents',
    azureAISearchApiKey: config.azureSearchKey!,
    azureAISearchEndpoint: config.azureSearchEndpoint!,
    {{#useOpenAI}}
    apiKey: config.openAIKey!,
    openAIEmbeddingModelName: config.openAIEmbeddingModelName!,
    {{/useOpenAI}}
    {{#useAzureOpenAI}}
    azureOpenAIApiKey: config.azureOpenAIKey!,
    azureOpenAIEndpoint: config.azureOpenAIEndpoint!,
    azureOpenAIEmbeddingDeploymentName: config.azureOpenAIEmbeddingDeploymentName!,
    {{/useAzureOpenAI}}
  })
);

// Define storage and application
const storage = new MemoryStorage();
const app = new Application<TurnState>({
  storage,
  ai: {
    planner,
  },
});

export default app;
