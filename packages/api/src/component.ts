import { Result } from "neverthrow";
import { FxError } from "./error";
import { InputsWithProjectPath } from "./v2/types";
import { ContextV3, ResourceContextV3 } from "./types";
import { Bicep } from "./bicep";
import { IProgressHandler } from "./qm";
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
    inputs: InputsWithProjectPath,
    actionContext?: ActionContext
  ) => Promise<Result<Bicep[], FxError>>;
  provision?: (
    context: ResourceContextV3,
    inputs: InputsWithProjectPath,
    actionContext?: ActionContext
  ) => Promise<Result<undefined, FxError>>;
  configure?: (
    context: ResourceContextV3,
    inputs: InputsWithProjectPath,
    actionContext?: ActionContext
  ) => Promise<Result<undefined, FxError>>;
  deploy?: (
    context: ResourceContextV3,
    inputs: InputsWithProjectPath,
    actionContext?: ActionContext
  ) => Promise<Result<undefined, FxError>>;
}

export interface ActionContext {
  progressBar?: IProgressHandler;
  telemetryProps?: Record<string, string>;
}
