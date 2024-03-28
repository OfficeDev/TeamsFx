import asyncio
from dataclasses import dataclass
import json
import os
import sys
import traceback
from typing import Generic, TypeVar

from botbuilder.core import MemoryStorage, TurnContext
from teams import Application, ApplicationOptions, TeamsAdapter
from teams.ai import AIOptions
from teams.ai.models import AzureOpenAIModelOptions, OpenAIModel, OpenAIModelOptions
from teams.ai.planners import ActionPlanner, ActionPlannerOptions
from teams.ai.prompts import PromptManager, PromptManagerOptions
from teams.ai.actions import ActionTypes
from teams.state import TurnState

from azure_ai_search_data_source import AzureAISearchDataSource, AzureAISearchDataSourceOptions
from config import Config

config = Config()

# Create AI components
model: OpenAIModel

{{#useAzureOpenAI}}
model = OpenAIModel(
    AzureOpenAIModelOptions(
        api_key=config.AZURE_OPENAI_API_KEY,
        default_model=config.AZURE_OPENAI_MODEL_DEPLOYMENT_NAME,
        endpoint=config.AZURE_OPENAI_ENDPOINT,
    )
)
{{/useAzureOpenAI}}    
{{#useOpenAI}}
model = OpenAIModel(
    OpenAIModelOptions(
        api_key=config.OPENAI_API_KEY,
        default_model=config.OPENAI_MODEL_NAME,
    )
)
{{/useOpenAI}}
    
prompts = PromptManager(PromptManagerOptions(prompts_folder=f"{os.getcwd()}/src/prompts"))

prompts.add_data_source(
    AzureAISearchDataSource(
        AzureAISearchDataSourceOptions(
            name='azure-ai-search',
            indexName='contoso-electronics',
            azureAISearchApiKey=config.AZURE_SEARCH_KEY,
            azureAISearchEndpoint=config.AZURE_SEARCH_ENDPOINT,
        )
    )
)

planner = ActionPlanner(
    ActionPlannerOptions(model=model, prompts=prompts, default_prompt="chat")
)

# Define storage and application
storage = MemoryStorage()
bot_app = Application[TurnState](
    ApplicationOptions(
        bot_app_id=config.APP_ID,
        storage=storage,
        adapter=TeamsAdapter(config),
        ai=AIOptions(planner=planner),
    )
)

@bot_app.conversation_update("membersAdded")
async def on_members_added(context: TurnContext, state: TurnState):
    await context.send_activity("How can I help you today?")

@bot_app.error
async def on_error(context: TurnContext, error: Exception):
    # This check writes out errors to console log .vs. app insights.
    # NOTE: In production environment, you should consider logging this to Azure
    #       application insights.
    print(f"\n [on_turn_error] unhandled error: {error}", file=sys.stderr)
    traceback.print_exc()

    # Send a message to the user
    await context.send_activity("The bot encountered an error or bug.")