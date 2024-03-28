// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  CancellationToken,
  ChatRequest,
  ChatResponseStream,
  LanguageModelChatUserMessage,
} from "vscode";
import { ISkill } from "./iSkill";
import { Spec } from "./spec";
import { ExecutionResultEnum } from "./executionResultEnum";
import { deepClone } from "../Utils";

export class SkillSet implements ISkill {
  name: string | undefined;
  capability: string | undefined;
  skills: ISkill[] | undefined;
  retriableTimes: number;

  constructor(skills: ISkill[], retriableTimes?: number) {
    this.name = "skillSet";
    this.capability = "A container for muultiple skills";
    this.skills = skills;
    this.retriableTimes = retriableTimes ?? 1;
  }

  public canInvoke(request: ChatRequest, spec: Spec): boolean {
    if (!this.skills) {
      return false;
    }
    return true;
  }

  // eslint-disable-next-line @typescript-eslint/require-await
  public async invoke(
    languageModel: LanguageModelChatUserMessage,
    request: ChatRequest,
    response: ChatResponseStream,
    token: CancellationToken,
    spec: Spec
  ): Promise<{ result: ExecutionResultEnum; spec: Spec }> {
    if (!this.skills) {
      return { result: ExecutionResultEnum.Success, spec: spec };
    }
    const specCopy = new Spec("");
    let retried = 0;
    while (retried < this.retriableTimes) {
      retried++;
      specCopy.clone(spec);
      let isSuccessed = true;

      for (const skill of this.skills) {
        if (!skill.canInvoke(request, specCopy)) {
          isSuccessed = false;
          continue;
        }
        const { result: result, spec: newSpec }: { result: ExecutionResultEnum; spec: Spec } =
          await skill.invoke(languageModel, request, response, token, specCopy);
        if (result === ExecutionResultEnum.Rejected) {
          // We want to keep the telemetry data anyway
          return { result: result, spec: newSpec };
        }
        if (result === ExecutionResultEnum.Failure) {
          isSuccessed = false;
        }
        if (result === ExecutionResultEnum.Success) {
          isSuccessed = true;
          specCopy.clone(newSpec);
        }
      }

      if (isSuccessed) {
        return { result: ExecutionResultEnum.Success, spec: specCopy };
      }
    }
    return { result: ExecutionResultEnum.Failure, spec: specCopy };
  }
}
