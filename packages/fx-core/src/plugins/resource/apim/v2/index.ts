// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  AzureAccountProvider,
  AzureSolutionSettings, err,
  Func,
  FxError,
  Inputs,
  Json,
  QTreeNode,
  Result,
  Stage,
  TokenProvider,
  traverse,
  Void
} from "@microsoft/teamsfx-api";
import {
  Context,
  DeploymentInputs, ProvisionInputs,
  ResourcePlugin,
  ResourceProvisionOutput
} from "@microsoft/teamsfx-api/build/v2";
import { Inject, Service } from "typedi";
import { ApimPlugin } from "..";
import {
  ResourcePlugins,
  ResourcePluginsV2
} from "../../../solution/fx-solution/ResourcePluginContainer";
import {
  configureResourceAdapter,
  convert2PluginContext,
  deployAdapter,
  executeUserTaskAdapter,
  getQuestionsForScaffoldingAdapter,
  provisionResourceAdapter,
  scaffoldSourceCodeAdapter
} from "../../utils4v2";

@Service(ResourcePluginsV2.ApimPlugin)
export class ApimPluginV2 implements ResourcePlugin {
  name = "fx-resource-apim";
  displayName = "API Management";
  @Inject(ResourcePlugins.ApimPlugin)
  plugin!: ApimPlugin;

  activate(solutionSettings: AzureSolutionSettings): boolean {
    return this.plugin.activate(solutionSettings);
  }
  async getQuestionsForScaffolding(
    ctx: Context,
    inputs: Inputs
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    return await getQuestionsForScaffoldingAdapter(ctx, inputs, this.plugin);
  }
  async scaffoldSourceCode(
    ctx: Context,
    inputs: Inputs
  ): Promise<Result<Void, FxError>> {
    return await scaffoldSourceCodeAdapter(ctx, inputs, this.plugin);
  }

  async provisionResource(
    ctx: Context,
    inputs: ProvisionInputs,
    provisionInputConfig: Json,
    tokenProvider: TokenProvider
  ): Promise<Result<ResourceProvisionOutput, FxError>> {
    return await provisionResourceAdapter(ctx, inputs, provisionInputConfig, tokenProvider, this.plugin);
  }

  async configureResource(
    ctx: Context,
    inputs: ProvisionInputs,
    provisionInputConfig: Json,
    provisionOutputs: Json,
    tokenProvider: TokenProvider
  ): Promise<Result<Json, FxError>> {
    return await configureResourceAdapter(
      ctx,
      inputs,
      provisionInputConfig,
      provisionOutputs,
      tokenProvider,
      this.plugin
    );
  }

  async deploy(
    ctx: Context,
    inputs: DeploymentInputs,
    provisionOutput: Json,
    tokenProvider: AzureAccountProvider
  ): Promise<Result<Json, FxError>> {
    const questionRes = await this.plugin.getQuestions(
      Stage.deploy,
      convert2PluginContext(ctx, inputs)
    );
    if (questionRes.isOk()) {
      const node = questionRes.value;
      if (node) {
        const res = await traverse(node, inputs, ctx.userInteraction);
        if (res.isErr()) {
          return err(res.error);
        }
      }
    }
    return await deployAdapter(ctx, inputs, provisionOutput, tokenProvider, this.plugin);
  }

  //addResource
  //TODO apim plugin implement executeUserTask() for addResource (preScaffold + scaffold)
  async executeUserTask(
    ctx: Context,
    inputs: Inputs,
    func: Func
  ): Promise<Result<unknown, FxError>> {
    const questionRes = await this.plugin.getQuestionsForUserTask(
      func,
      convert2PluginContext(ctx, inputs)
    );
    if (questionRes.isOk()) {
      const node = questionRes.value;
      if (node) {
        const res = await traverse(node, inputs, ctx.userInteraction);
        if (res.isErr()) {
          return err(res.error);
        }
      }
    }
    return await executeUserTaskAdapter(ctx, inputs, func, this.plugin);
  }
}
