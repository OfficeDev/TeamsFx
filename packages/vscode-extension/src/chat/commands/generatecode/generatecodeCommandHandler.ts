// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  ChatRequest,
  ChatContext,
  ChatResponseStream,
  CancellationToken,
  LanguageModelChatUserMessage,
} from "vscode";
import { OfficeAddinChatCommand } from "../../consts";
import { ICopilotChatResult } from "../../types";
import { Planner } from "../../officeCommon/planner";

// TODO: Implement the function.
export default async function generatecodeCommandHandler(
  request: ChatRequest,
  context: ChatContext,
  response: ChatResponseStream,
  token: CancellationToken
): Promise<ICopilotChatResult> {
  return await Planner.getInstance().processRequest(
    new LanguageModelChatUserMessage(request.prompt),
    request,
    response,
    token,
    OfficeAddinChatCommand.GenerateCode
  );
}
