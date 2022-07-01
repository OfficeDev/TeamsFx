import {
  Func,
  FxError,
  Inputs,
  Json,
  M365TokenProvider,
  QTreeNode,
  Result,
  TokenProvider,
  v2,
  Void,
} from "@microsoft/teamsfx-api";
import { DeepReadonly, EnvInfoV2 } from "@microsoft/teamsfx-api/build/v2";
import { Service } from "typedi";
import { isPreviewFeaturesEnabled } from "../../../../common";
import { PluginDisplayName } from "../../../../common/constants";
import { SolutionPluginsV2 } from "../../../../core/SolutionPluginContainer";
import { checkPermission } from "./checkPermission";
import { TeamsAppSolutionNameV2 } from "./constants";
import { createEnv } from "./createEnv";
import { deploy } from "./deploy";
import { executeUserTask } from "./executeUserTask";
import { generateResourceTemplate } from "./generateResourceTemplate";
import {
  getQuestions,
  getQuestionsForScaffolding,
  getQuestionsForScaffoldingPreview,
  getQuestionsForUserTask,
} from "./getQuestions";
import { grantPermission } from "./grantPermission";
import { listCollaborator } from "./listCollaborator";
import { provisionResource } from "./provision";
import { provisionLocalResource } from "./provisionLocal";
import { publishApplication } from "./publish";
import { scaffoldSourceCode } from "./scaffolding";
import "../../../resource/officeaddin";

@Service(SolutionPluginsV2.AzureTeamsSolutionV2)
export class TeamsAppSolutionV2 implements v2.SolutionPlugin {
  name = TeamsAppSolutionNameV2;
  displayName: string = PluginDisplayName.Solution;

  scaffoldSourceCode: (ctx: v2.Context, inputs: Inputs) => Promise<Result<Void, FxError>> =
    scaffoldSourceCode;
  generateResourceTemplate: (ctx: v2.Context, inputs: Inputs) => Promise<Result<Json, FxError>> =
    generateResourceTemplate;

  provisionResources = provisionResource;

  deploy?: (
    ctx: v2.Context,
    inputs: Inputs,
    envInfo: EnvInfoV2,
    tokenProvider: TokenProvider
  ) => Promise<Result<Void, FxError>> = deploy;

  publishApplication: (
    ctx: v2.Context,
    inputs: Inputs,
    envInfo: v2.EnvInfoV2,
    tokenProvider: M365TokenProvider
  ) => Promise<Result<Void, FxError>> = publishApplication;
  provisionLocalResource?: (
    ctx: v2.Context,
    inputs: Inputs,
    localSettings: Json,
    tokenProvider: TokenProvider,
    envInfo?: EnvInfoV2
  ) => Promise<v2.FxResult<Json, FxError>> = provisionLocalResource;

  getQuestionsForScaffolding?: (
    ctx: v2.Context,
    inputs: Inputs
  ) => Promise<Result<QTreeNode | undefined, FxError>> = isPreviewFeaturesEnabled()
    ? getQuestionsForScaffoldingPreview
    : getQuestionsForScaffolding;

  executeUserTask?: (
    ctx: v2.Context,
    inputs: Inputs,
    func: Func,
    localSettings: Json,
    envInfo: EnvInfoV2,
    tokenProvider: TokenProvider
  ) => Promise<Result<unknown, FxError>> = executeUserTask;

  createEnv?: (ctx: v2.Context, inputs: Inputs) => Promise<Result<Void, FxError>> = createEnv;

  getQuestions?: (
    ctx: v2.Context,
    inputs: Inputs,
    envInfo: DeepReadonly<EnvInfoV2>,
    tokenProvider: TokenProvider
  ) => Promise<Result<QTreeNode | undefined, FxError>> = getQuestions;

  getQuestionsForUserTask?: (
    ctx: v2.Context,
    inputs: Inputs,
    func: Func,
    envInfo: DeepReadonly<EnvInfoV2>,
    tokenProvider: TokenProvider
  ) => Promise<Result<QTreeNode | undefined, FxError>> = getQuestionsForUserTask;

  grantPermission?: (
    ctx: v2.Context,
    inputs: v2.InputsWithProjectPath,
    envInfo: DeepReadonly<EnvInfoV2>,
    tokenProvider: TokenProvider
  ) => Promise<Result<Json, FxError>> = (
    ctx: v2.Context,
    inputs: v2.InputsWithProjectPath,
    envInfo: DeepReadonly<EnvInfoV2>,
    tokenProvider: TokenProvider
  ) => grantPermission({ apiVersion: 2, ctx, inputs, envInfo, tokenProvider });

  checkPermission?: (
    ctx: v2.Context,
    inputs: v2.InputsWithProjectPath,
    envInfo: DeepReadonly<EnvInfoV2>,
    tokenProvider: TokenProvider
  ) => Promise<Result<Json, FxError>> = (
    ctx: v2.Context,
    inputs: v2.InputsWithProjectPath,
    envInfo: DeepReadonly<EnvInfoV2>,
    tokenProvider: TokenProvider
  ) => checkPermission({ apiVersion: 2, ctx, inputs, envInfo, tokenProvider });

  listCollaborator?: (
    ctx: v2.Context,
    inputs: v2.InputsWithProjectPath,
    envInfo: DeepReadonly<EnvInfoV2>,
    tokenProvider: TokenProvider
  ) => Promise<Result<Json, FxError>> = (
    ctx: v2.Context,
    inputs: v2.InputsWithProjectPath,
    envInfo: DeepReadonly<EnvInfoV2>,
    tokenProvider: TokenProvider
  ) => listCollaborator({ apiVersion: 2, ctx, inputs, envInfo, tokenProvider });
}
