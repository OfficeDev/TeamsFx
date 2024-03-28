// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  CancellationToken,
  ChatRequest,
  ChatResponseStream,
  LanguageModelChatMessage,
  LanguageModelChatUserMessage,
} from "vscode";
import { ISkill } from "./iSkill"; // Add the missing import statement
import { Spec } from "./spec";
import { getCopilotResponseAsString } from "../../utils";
import { ExecutionResultEnum } from "./executionResultEnum";
import { CHAT_CREATE_OFFICEADDIN_SAMPLE_COMMAND_ID } from "../../consts";
import { localize } from "../../../utils/localizeUtils";

export class CodeMerger implements ISkill {
  name: string | undefined;
  capability: string | undefined;

  constructor() {
    this.name = "Code Merger";
    this.capability = "Merge code snippets into the generated template";
  }

  public canInvoke(request: ChatRequest, spec: Spec): boolean {
    return (
      !!spec.userInput &&
      !!spec.appendix.codeSnippet &&
      !!spec.appendix.codeTaskBreakdown &&
      spec.appendix.codeTaskBreakdown.length > 0 &&
      !!spec.appendix.tempAppLocation &&
      spec.appendix.tempAppLocation.length > 0
    );
  }

  // eslint-disable-next-line @typescript-eslint/require-await
  public async invoke(
    languageModel: LanguageModelChatUserMessage,
    request: ChatRequest,
    response: ChatResponseStream,
    token: CancellationToken,
    spec: Spec
  ): Promise<{ result: ExecutionResultEnum; spec: Spec }> {
    const sampleTitle = localize("teamstoolkit.chatParticipants.create.sample");
    response.button({
      command: CHAT_CREATE_OFFICEADDIN_SAMPLE_COMMAND_ID,
      arguments: [spec.appendix.tempAppLocation],
      title: sampleTitle,
    });
    return { result: ExecutionResultEnum.Success, spec: spec };
  }
}
