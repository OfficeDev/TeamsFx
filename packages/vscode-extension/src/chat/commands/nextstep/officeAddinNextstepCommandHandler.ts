// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  ChatRequest,
  ChatContext,
  ChatResponseStream,
  CancellationToken,
  LanguageModelChatUserMessage,
} from "vscode";
import { ICopilotChatResult } from "../../types";
import { OfficeAddinChatCommand } from "../../consts";
import { Correlator } from "@microsoft/teamsfx-core";
import { getCopilotResponseAsString } from "../../utils";
import { defaultSystemPrompt } from "../../prompts";
import { ChatTelemetryData } from "../../telemetry";
import { ExtTelemetry } from "../../../telemetry/extTelemetry";
import { TelemetryEvent } from "../../../telemetry/extTelemetryEvents";

//TODO: Implement the function.
export default async function officeAddinNextStepCommandHandler(
  request: ChatRequest,
  context: ChatContext,
  response: ChatResponseStream,
  token: CancellationToken
): Promise<ICopilotChatResult> {
  const chatTelemetryData = ChatTelemetryData.createByCommand(OfficeAddinChatCommand.NextStep);
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.CopilotChatStart, chatTelemetryData.properties);

  const messages = [defaultSystemPrompt(), new LanguageModelChatUserMessage(request.prompt)];
  await getCopilotResponseAsString("copilot-gpt-3.5-turbo", messages, token);
  return {
    metadata: {
      command: OfficeAddinChatCommand.NextStep,
      requestId: chatTelemetryData.requestId,
    },
  };
}
