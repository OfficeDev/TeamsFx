"""
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
"""

from typing import Any, Dict, Optional

from botbuilder.core import Storage, TurnContext
from teams.state import TurnState, ConversationState, UserState, TempState


class AppConversationState(ConversationState):
    tasks: Dict[str, Any]=None

    @classmethod
    async def load(cls, context: TurnContext, storage: Optional[Storage] = None) -> "AppConversationState":
        state = await super().load(context, storage)
        return cls(**state)


class AppTurnState(TurnState[AppConversationState, UserState, TempState]):
    conversation: AppConversationState

    @classmethod
    async def load(cls, context: TurnContext, storage: Optional[Storage] = None) -> "AppTurnState":
        return cls(
            conversation=await AppConversationState.load(context, storage),
            user=await UserState.load(context, storage),
            temp=await TempState.load(context, storage),
        )