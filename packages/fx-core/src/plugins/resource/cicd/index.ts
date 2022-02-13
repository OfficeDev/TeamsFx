// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  err,
  UserError,
  SystemError,
  v2,
  TokenProvider,
  FxError,
  Result,
  Inputs,
  Json,
  Func,
  ok,
  QTreeNode,
  ProjectSettings,
} from "@microsoft/teamsfx-api";

import { FxResult, FxCICDPluginResultFactory as ResultFactory } from "./result";
import { CICDImpl } from "./plugin";
import { ErrorType, PluginError } from "./errors";
import { LifecycleFuncNames } from "./constants";
import { Service } from "typedi";
import { ResourcePluginsV2 } from "../../solution/fx-solution/ResourcePluginContainer";
import { ResourcePlugin, Context, DeepReadonly, EnvInfoV2 } from "@microsoft/teamsfx-api/build/v2";
import {
  githubOption,
  azdoOption,
  jenkinsOption,
  ciOption,
  cdOption,
  provisionOption,
  publishOption,
} from "./questions";

@Service(ResourcePluginsV2.CICDPlugin)
export class CICDPluginV2 implements ResourcePlugin {
  name = "fx-resource-cicd";
  displayName = "CICD";
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  activate(projectSettings: ProjectSettings): boolean {
    return true;
  }

  public cicdImpl: CICDImpl = new CICDImpl();

  public async addCICDWorkflows(
    context: Context,
    projectPath: string,
    template: string,
    envName: string
  ): Promise<FxResult> {
    const result = await this.runWithExceptionCatching(
      context,
      () => this.cicdImpl.addCICDWorkflows(context, projectPath, template, envName),
      true,
      LifecycleFuncNames.ADD_CICD_WORKFLOWS
    );

    return result;
  }

  public async getQuestionsForUserTask(
    ctx: Context,
    inputs: Inputs,
    func: Func,
    envInfo: DeepReadonly<EnvInfoV2>,
    tokenProvider: TokenProvider
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    const cicdWorkflowQuestions = new QTreeNode({
      type: "group",
    });

    const whichPlatform = new QTreeNode({
      name: "which_platform",
      type: "singleSelect",
      staticOptions: [githubOption, azdoOption, jenkinsOption],
      title: "Select a CI/CD Platform",
      default: githubOption.id,
    });

    const whichTemplate = new QTreeNode({
      name: "template",
      type: "multiSelect",
      staticOptions: [ciOption, cdOption, provisionOption, publishOption],
      title: "Select template(s)",
      default: [ciOption.id],
    });

    cicdWorkflowQuestions.addChild(whichPlatform);
    cicdWorkflowQuestions.addChild(whichTemplate);

    return ok(cicdWorkflowQuestions);
  }

  public async executeUserTask(
    ctx: Context,
    inputs: Inputs,
    func: Func,
    localSettings: Json,
    envInfo: v2.EnvInfoV2,
    tokenProvider: TokenProvider
  ): Promise<Result<unknown, FxError>> {
    if (func.method === "addCICDWorkflows") {
      return await this.runWithExceptionCatching(
        ctx,
        () => this.addCICDWorkflows(ctx, inputs.projectPath!, inputs["template"], envInfo.envName),
        true,
        LifecycleFuncNames.ADD_CICD_WORKFLOWS
      );
    }
    return ok(undefined);
  }

  private async runWithExceptionCatching(
    context: Context,
    fn: () => Promise<FxResult>,
    sendTelemetry: boolean,
    name: string
  ): Promise<FxResult> {
    try {
      const res: FxResult = await fn();
      return res;
    } catch (e) {
      if (e instanceof UserError || e instanceof SystemError) {
        const res = err(e);
        return res;
      }

      if (e instanceof PluginError) {
        const result =
          e.errorType === ErrorType.System
            ? ResultFactory.SystemError(e.name, e.genMessage(), e.innerError)
            : ResultFactory.UserError(e.name, e.genMessage(), e.showHelpLink, e.innerError);
        return result;
      } else {
        // Unrecognized Exception.
        const UnhandledErrorCode = "UnhandledError";
        return ResultFactory.SystemError(UnhandledErrorCode, e.message, e);
      }
    }
  }
}

export default new CICDPluginV2();
