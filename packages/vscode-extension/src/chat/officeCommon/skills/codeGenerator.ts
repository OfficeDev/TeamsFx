// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import ts = require("typescript");
import {
  CancellationToken,
  ChatRequest,
  ChatResponseStream,
  LanguageModelChatMessage,
  LanguageModelChatSystemMessage,
  LanguageModelChatUserMessage,
} from "vscode";
import { compressCode } from "../Utils";
import { SampleProvider } from "../samples/sampleProvider";
import { getCodeGenerateGuidance } from "./codeGuidance";
import { ISkill } from "./iSkill"; // Add the missing import statement
import { Spec } from "./spec";
import { getCopilotResponseAsString } from "../../utils";
import { ExecutionResultEnum } from "./executionResultEnum";
import {
  MeasurementCodeGenAttemptCount,
  MeasurementCodeGenExecutionTimeInTotalSec,
  MeasurementScenarioBasedSampleMatchedCount,
  PropertySystemCodeGenIsCustomFunction,
  PropertySystemCodeGenResult,
  PropertySystemCodeGenTargetedOfficeHostApplication,
  MeasurementSystemCodegenTaskBreakdownAttemptFailedCount,
} from "../telemetryConsts";

const excelSystemPrompt = `
The following content written using Markdown syntax, using "Bold" style to highlight the key information.
`;
const cfSystemPrompt = `
The following content written using Markdown syntax, using "Bold" style to highlight the key information.

There're some references help you to understand some key concepts, read it and repeat by yourself, before start to generate code.
# References:
## Understanding the difference between a Custom Functions and the normal TypeScript/JavaScript function:
In the context of Office Excel Custom Functions, there are several differences compared to normal JavaScript/TypeScript functions:
## Metadata 
Custom Functions require metadata that specifies the function name, parameters, return value, etc. This metadata is used by Excel to properly use the function.

## Async Pattern
Custom Functions can be asynchronous, but they must follow a specific pattern. They should return a Promise object, and Excel will wait for the Promise to resolve to get the result.

## Streaming Pattern
For streaming Custom Functions, they must follow a specific pattern. They should take a handler parameter (typically the last parameter), and call the handler.setResult method to update the cell value.

## Error Handling
To return an error from a Custom Function, you should throw an OfficeExtension.Error object with a specific error code.

## Limited API Access
Custom Functions can only call a subset of the Office JavaScript API that is specifically designed for Custom Functions.

## Stateless
Custom Functions are stateless, meaning they don't retain information between function calls. Each call to a function has separate memory and computation.

## Cancellation
Custom Functions should handle cancellation requests from Excel. When Excel cancels a function call, it rejects the Promise with an "OfficeExtension.Error" object that has the error code "OfficeExtension.ErrorCodes.generalException".

## Example of a Custom Function:
\`\`\`typescript
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 */
function secondHighest(values) {
  let highest = values[0][0],
    secondHighest = values[0][0];
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      if (values[i][j] >= highest) {
        secondHighest = highest;
        highest = values[i][j];
      } else if (values[i][j] >= secondHighest) {
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
\`\`\`
The @customfunction tag in the JSDoc comment is used to indicate that this is a Custom Function. The @param and @returns tags are used to specify the parameters and return value. It's important to follow this pattern when creating Custom Functions in Excel.

## Invocation parameter
Every custom function is automatically passed an invocation argument as the last input parameter, even if it's not explicitly declared. This invocation parameter corresponds to the Invocation object. The Invocation object can be used to retrieve additional context, such as the address of the cell that invoked your custom function. To access the Invocation object, you must declare invocation as the last parameter in your custom function.
The following sample shows how to use the invocation parameter to return the address of the cell that invoked your custom function. This sample uses the address property of the Invocation object. To access the Invocation object, first declare CustomFunctions.Invocation as a parameter in your JSDoc. Next, declare @requiresAddress in your JSDoc to access the address property of the Invocation object. Finally, within the function, retrieve and then return the address property.
\`\`\`typescript
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function getAddress(first, second, invocation) {
  const address = invocation.address;
  return address;
}
\`\`\`

So once you understand the concept of Custom Functions, you should make sure:
- The JSDoc comment is correctly added to the function.
- The function must return a value.
- The invocation parameter is correctly added to the function.
- The function follows the asynchronous pattern if necessary.
- The function follows the streaming pattern if necessary.
- Although that is not forbidden, but you should explicitly state in your code that the function must avoid using the Office JavaScript API.

Let's think step by step.
`;

export class CodeGenerator implements ISkill {
  name: string;
  capability: string;

  constructor() {
    this.name = "Code Generator";
    this.capability = "Generate code";
  }

  public canInvoke(request: ChatRequest, spec: Spec): boolean {
    return !!request.prompt && request.prompt.length > 0 && !!spec;
  }

  public async invoke(
    languageModel: LanguageModelChatUserMessage,
    request: ChatRequest,
    response: ChatResponseStream,
    token: CancellationToken,
    spec: Spec
  ): Promise<{ result: ExecutionResultEnum; spec: Spec }> {
    if (
      !!spec.appendix.host ||
      !!spec.appendix.codeTaskBreakdown ||
      (spec.appendix.codeTaskBreakdown as string[]).length == 0
    ) {
      const breakdownResult = await this.userInputBreakdownTaskAsync(request, token);

      if (!breakdownResult) {
        if (
          !spec.appendix.telemetryData.measurements[
            MeasurementSystemCodegenTaskBreakdownAttemptFailedCount
          ]
        ) {
          spec.appendix.telemetryData.measurements[
            MeasurementSystemCodegenTaskBreakdownAttemptFailedCount
          ] = 0;
        }
        spec.appendix.telemetryData.measurements[
          MeasurementSystemCodegenTaskBreakdownAttemptFailedCount
        ] += 1;
        return { result: ExecutionResultEnum.Failure, spec: spec };
      }
      if (!breakdownResult.shouldContinue) {
        // Reject will make the whole request rejected
        spec.sections = breakdownResult.data;
        return { result: ExecutionResultEnum.Rejected, spec: spec };
      }

      spec.appendix.host = breakdownResult.host;
      spec.appendix.codeTaskBreakdown = breakdownResult.data;
      spec.appendix.isCustomFunction = breakdownResult.customFunctions;
    }

    if (!spec.appendix.telemetryData.measurements[MeasurementCodeGenAttemptCount]) {
      spec.appendix.telemetryData.measurements[MeasurementCodeGenAttemptCount] = 0;
    }
    spec.appendix.telemetryData.measurements[MeasurementCodeGenAttemptCount] += 1;
    let codeSnippet: string | null = "";
    const t0 = performance.now();
    codeSnippet = await this.generateCode(
      request,
      token,
      spec.appendix.host,
      spec.appendix.isCustomFunction,
      spec.appendix.codeTaskBreakdown,
      spec
    );
    const t1 = performance.now();
    const duration = (t1 - t0) / 1000;
    if (!spec.appendix.telemetryData.measurements[MeasurementCodeGenExecutionTimeInTotalSec]) {
      spec.appendix.telemetryData.measurements[MeasurementCodeGenExecutionTimeInTotalSec] =
        duration;
    } else {
      spec.appendix.telemetryData.measurements[MeasurementCodeGenExecutionTimeInTotalSec] +=
        duration;
    }
    console.log(`Code generation took ${duration} seconds.`);
    if (!codeSnippet) {
      spec.appendix.telemetryData.properties[PropertySystemCodeGenResult] = "false";
      return { result: ExecutionResultEnum.Failure, spec: spec };
    }

    spec.appendix.telemetryData.properties[PropertySystemCodeGenResult] = "true";
    spec.appendix.codeSnippet = codeSnippet;
    return { result: ExecutionResultEnum.Success, spec: spec };
  }

  async userInputBreakdownTaskAsync(
    request: ChatRequest,
    token: CancellationToken
  ): Promise<null | {
    host: string;
    shouldContinue: boolean;
    customFunctions: boolean;
    data: string[];
  }> {
    const userPrompt = `
    Assume this is a ask: "${request.prompt}". I need you help to analyze it, and give me your suggestion. Follow the guidance below:
    - If the ask is not relevant to Microsoft Excel, Microsoft Word, or Microsoft PowerPoint, you should reject it because today this agent only support offer assistant to those Office host applications. And give the reason to reject the ask.
    - If the ask is not about automating a certain process or accomplishing a certain task using Office JavaScript Add-ins, you should reject it. And give the reason to reject the ask.
    - If the ask is **NOT JUST** asking for generate **TypeScript** or **JavaScript** code for Office Add-ins. You should reject it. And give the reason to reject the ask. For example, if part of the ask is about generating code of VBA, Python, HTML, CSS, or other languages, you should reject it. If that is not relevant to Office Add-ins, you should reject it. etc.
    - Otherwise, please think about if you can process the ask. 
      - If you cannot process the ask, you should reject it. And give me the reason to reject the ask.
      - If you can process the ask, you should break down the ask into sub steps that could be performed by Office Add-ins JavaScript APIs. Each step should be actions accomplished by using **code**. Emphasize the "Bold" part in the title.
    return the result in a JSON object.

    Think about that step by step.
    `;
    const defaultSystemPrompt = `
    The following content written using Markdown syntax, using "Bold" style to highlight the key information.

    #Role:
    You are an expert in Office JavaScript Add-ins, and you are familiar with scenario and the capabilities of Office JavaScript Add-ins. You need to offer the user a suggestion based on the user's ask.

    #Your tasks:
    Repeat the user's ask, and then give your suggestion based on the user's ask. Follow the guidance below:
    If you suggested to accept the ask. Put the list of sub tasks into the "data" field of the output JSON object. A "shouldContinue" field on that JSON object should be true.
    If you suggested to reject the ask, put the reason to reject into the "data" field of the output JSON object. A "shouldContinue" field on that JSON object should be false.
    You must strickly follow the format of output.

    #The format of output:
    The output should be just a **JSON object**. You should not add anything else to the output
    - The first key named "host", that value is a string to indicate which Office application is the most relevant to the user's ask. You can pick from "Excel", "Word", "PowerPoint". 
    - The second key is "shouldContinue", the value is a Boolean.
    - The third key named "data", the value of it is the list of sub tasks or rejection reason, and that is a string array.
    - The last key named "customFunctions", set value of it to be a Boolean true if the user's ask is about Office JavaScript Add-ins with custom functions on Excel. Otherwise, set it to be a Boolean false.
    If the value of "shouldContinue" is true, then the value of "data" should be the list of sub tasks; if the value of "shouldContinue" is false, then the value of "data" should be the list of missing information or reason to reject. **Beyond this JSON object, you should not add anything else to the output**.

    Think about that step by step.
    `;

    // Perform the desired operation
    const messages: LanguageModelChatMessage[] = [
      new LanguageModelChatSystemMessage(defaultSystemPrompt),
      new LanguageModelChatUserMessage(userPrompt),
    ];
    const copilotResponse = await getCopilotResponseAsString(
      "copilot-gpt-3.5-turbo",
      messages,
      token
    );
    let copilotRet = {
      host: "",
      shouldContinue: false,
      customFunctions: false,
      data: [],
    };

    try {
      copilotRet = JSON.parse(copilotResponse.trim());
    } catch (error) {
      console.error("[User task breakdown] Failed to parse the response from Copilot:", error);
      return null;
    }

    return copilotRet;
  }

  async generateCode(
    request: ChatRequest,
    token: CancellationToken,
    host: string,
    isCustomFunctions: boolean,
    subTasks: string[],
    spec: Spec
  ) {
    const userPrompt = `
The following content written using Markdown syntax, using "Bold" style to highlight the key information.

# Your role:
You're a professional and senior Office JavaScript Add-ins developer with a lot of experience and know all best practice on JavaScript, CSS, HTML, popular algorithm, and Office Add-ins API. You should help the user to automate a certain process or accomplish a certain task using Office JavaScript Add-ins.

# Context:
This is the ask need your help to generate the code for this request:
- ${request.prompt}. 
The request is about Office Add-ins, and it is relevant to the Office application "${host}".
It could be broken down into a few steps able to be accomplished by Office Add-ins JavaScript APIs. **Read through the those steps, repeat by yourself**. Make sure you understand that before go to the task. You have the list of steps.:
${subTasks.map((task, index) => `${index + 1}. ${task}`).join("\n")}

# Your tasks:
Implement **all** steps with **TypeScript code** and **Office JavaScript Add-ins API**, while **follow the coding rule**.

${getCodeGenerateGuidance(host)}

# Format of output:
**You must strickly follow the format of output**. The output will only contains code without any explanation on the code or generate process. Beyond that, nothing else should be included in the output.
- The code surrounded by a pair of triple backticks, and must follow with a string "typescript". For example:
\`\`\`typescript
// The code snippet
\`\`\`

Let's think step by step.
    `;
    spec.appendix.telemetryData.properties[PropertySystemCodeGenTargetedOfficeHostApplication] =
      host;
    spec.appendix.telemetryData.properties[PropertySystemCodeGenIsCustomFunction] =
      isCustomFunctions.toString();
    let defaultSystemPrompt = `
    The following content written using Markdown syntax, using "Bold" style to highlight the key information.

    # There're some samples relevant to the your's ask, you can read it and repeat by yourself, before start to generate code.
    `;
    let referenceUserPrompt = "";
    switch (host) {
      case "Excel":
        if (!isCustomFunctions) {
          referenceUserPrompt = excelSystemPrompt;
        } else {
          referenceUserPrompt = cfSystemPrompt;
        }
        break;
      default:
        defaultSystemPrompt = "";
        break;
    }

    // Then let's query if any code examples relevant to the user's ask that we can put as examples
    const scenarioSamples =
      await SampleProvider.getInstance().getTopKMostRelevantScenarioSampleCodes(
        request,
        token,
        host,
        request.prompt,
        1
      );
    if (scenarioSamples.size > 0) {
      const codeSnippets: string[] = [];
      scenarioSamples.forEach((sample, api) => {
        codeSnippets.push(`- ${sample.description}:
                              \`\`\`typescript
                              ${compressCode(sample.codeSample)}
                              \`\`\`\n`);
      });

      if (codeSnippets.length > 0) {
        defaultSystemPrompt = defaultSystemPrompt.concat(`\n${codeSnippets.join("\n")}\n\n`);
      }
    }
    if (!spec.appendix.telemetryData.measurements[MeasurementScenarioBasedSampleMatchedCount]) {
      spec.appendix.telemetryData.measurements[MeasurementScenarioBasedSampleMatchedCount] = 0;
    }
    spec.appendix.telemetryData.measurements[MeasurementScenarioBasedSampleMatchedCount] +=
      scenarioSamples.size > 0 ? 1 : 0;

    // Perform the desired operation
    const messages: LanguageModelChatMessage[] = [
      new LanguageModelChatSystemMessage(defaultSystemPrompt),
      new LanguageModelChatUserMessage(referenceUserPrompt),
      new LanguageModelChatUserMessage(userPrompt),
    ];
    // The "copilot-gpt-4" model is significantly slower than "copilot-gpt-3.5-turbo", but also significantly more accurate
    // In order to avoid waste more time on the correct, I believe using GPT-4 is a better choice
    const copilotResponse = await getCopilotResponseAsString("copilot-gpt-4", messages, token);

    // extract the code snippet and the api list out
    const codeSnippetRet = copilotResponse.match(/```typescript([\s\S]*?)```/);
    if (!codeSnippetRet) {
      // something wrong with the LLM output
      // TODO: Add handling for this case
      console.error(
        "[Code generation] Failed to extract the code snippet from the response:",
        copilotResponse
      );
      return null;
    }

    return codeSnippetRet[1].trim();
  }
}
