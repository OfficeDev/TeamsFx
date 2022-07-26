import { Result } from "neverthrow";
import { FxError } from "./error";
import { InputsWithProjectPath } from "./v2/types";
import { Action } from "./action";
import { ContextV3, MaybePromise } from "./types";
export { InputsWithProjectPath };
export interface ResourceOutput {
  key: string;
  bicepVariable?: string;
}

export interface ResourceOutputs {
  [k: string]: ResourceOutput;
}

export interface CloudResource {
  readonly name: string;
  readonly description?: string;
  readonly outputs: ResourceOutputs;
  readonly finalOutputKeys: string[];
  readonly secretKeys?: string[];
  generateBicep?: (
    context: ContextV3,
    inputs: InputsWithProjectPath
  ) => MaybePromise<Result<Action | undefined, FxError>>; // MaybePromise<Result<Bicep, FxError>>
  provision?: (
    context: ContextV3,
    inputs: InputsWithProjectPath
  ) => MaybePromise<Result<Action | undefined, FxError>>; // MaybePromise<Result<undefined, FxError>>
  configure?: (
    context: ContextV3,
    inputs: InputsWithProjectPath
  ) => MaybePromise<Result<Action | undefined, FxError>>; // MaybePromise<Result<undefined, FxError>>
  deploy?: (
    context: ContextV3,
    inputs: InputsWithProjectPath
  ) => MaybePromise<Result<Action | undefined, FxError>>; // MaybePromise<Result<undefined, FxError>>
}

export interface SourceCodeProvider {
  readonly name: string;
  readonly description?: string;
  generate: (
    context: ContextV3,
    inputs: InputsWithProjectPath
  ) => MaybePromise<Result<Action | undefined, FxError>>; // MaybePromise<Result<undefined, FxError>>
  build?: (
    context: ContextV3,
    inputs: InputsWithProjectPath
  ) => MaybePromise<Result<Action | undefined, FxError>>; // MaybePromise<Result<undefined, FxError>>
}
