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

export class Explainer implements ISkill {
  name: string | undefined;
  capability: string | undefined;

  constructor() {
    this.name = "Explainer";
    this.capability = "Explain code snippet";
  }
  public canInvoke(request: ChatRequest, spec: Spec): boolean {
    return (
      !!spec.userInput &&
      !!spec.appendix.codeSnippet &&
      !!spec.appendix.codeTaskBreakdown &&
      spec.appendix.codeTaskBreakdown.length > 0
    );
  }

  public async invoke(
    languageModel: LanguageModelChatUserMessage,
    request: ChatRequest,
    response: ChatResponseStream,
    token: CancellationToken,
    spec: Spec
  ): Promise<{ result: ExecutionResultEnum; spec: Spec }> {
    const systemPrompt = `
Based on the user's input ${spec.userInput}, and the breakdown of the ask:
- ${spec.appendix.codeTaskBreakdown.join("\n- ")}

As output, you shou'd only provide the explanation for the code snippet, not the code snippet itself. The output should be in the format of Markdown.
`;

    const userPrompt = `
Please explain the code snippet below:
\`\`\`typescript
${spec.appendix.codeSnippet}
\`\`\`

Let's think it step by step.
    `;

    // Perform the desired operation
    const messages: LanguageModelChatMessage[] = [
      new LanguageModelChatUserMessage(systemPrompt),
      new LanguageModelChatUserMessage(userPrompt),
    ];
    const copilotResponse = await getCopilotResponseAsString(
      "copilot-gpt-3.5-turbo",
      messages,
      token
    );

    if (!copilotResponse) {
      // something wrong with the LLM output
      // however it's not a hard stop, still ok produce the output without explanation
      return { result: ExecutionResultEnum.Failure, spec: spec };
    }

    spec.appendix.codeExplanation = copilotResponse;

    return { result: ExecutionResultEnum.Success, spec: spec };
  }
}
