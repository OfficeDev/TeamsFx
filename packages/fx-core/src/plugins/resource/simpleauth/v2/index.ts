// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  AzureSolutionSettings,
  FxError,
  Inputs,
  Json,
  ProjectSettings,
  Result,
  TokenProvider,
  v2,
  Void,
} from "@microsoft/teamsfx-api";
import { EnvInfoV2 } from "@microsoft/teamsfx-api/build/v2";
import { Inject, Service } from "typedi";
import { SimpleAuthPlugin } from "../..";
import {
  ResourcePlugins,
  ResourcePluginsV2,
} from "../../../solution/fx-solution/ResourcePluginContainer";
import {
  configureLocalResourceAdapter,
  configureResourceAdapter,
  generateResourceTemplateAdapter,
  provisionLocalResourceAdapter,
  provisionResourceAdapter,
  updateResourceTemplateAdapter,
} from "../../utils4v2";

@Service(ResourcePluginsV2.SimpleAuthPlugin)
export class SimpleAuthPluginV2 implements v2.ResourcePlugin {
  name = "fx-resource-simple-auth";
  displayName = "Simple Auth";
  @Inject(ResourcePlugins.SimpleAuthPlugin)
  plugin!: SimpleAuthPlugin;
  activate(projectSettings: ProjectSettings): boolean {
    const solutionSettings = projectSettings.solutionSettings as AzureSolutionSettings;
    return this.plugin.activate(solutionSettings);
  }

  async generateResourceTemplate(
    ctx: v2.Context,
    inputs: Inputs
  ): Promise<Result<v2.ResourceTemplate, FxError>> {
    return await generateResourceTemplateAdapter(ctx, inputs, this.plugin);
  }
  async updateResourceTemplate(
    ctx: v2.Context,
    inputs: Inputs
  ): Promise<Result<v2.ResourceTemplate, FxError>> {
    return await updateResourceTemplateAdapter(ctx, inputs, this.plugin);
  }
  async provisionLocalResource(
    ctx: v2.Context,
    inputs: Inputs,
    localSettings: Json,
    tokenProvider: TokenProvider,
    envInfo?: EnvInfoV2
  ): Promise<Result<Void, FxError>> {
    return await provisionLocalResourceAdapter(
      ctx,
      inputs,
      localSettings,
      tokenProvider,
      this.plugin,
      envInfo
    );
  }

  async configureLocalResource(
    ctx: v2.Context,
    inputs: Inputs,
    localSettings: Json,
    tokenProvider: TokenProvider,
    envInfo?: EnvInfoV2
  ): Promise<Result<Void, FxError>> {
    return await configureLocalResourceAdapter(
      ctx,
      inputs,
      localSettings,
      tokenProvider,
      this.plugin,
      envInfo
    );
  }

  async provisionResource(
    ctx: v2.Context,
    inputs: v2.ProvisionInputs,
    envInfo: v2.EnvInfoV2,
    tokenProvider: TokenProvider
  ): Promise<Result<Void, FxError>> {
    return await provisionResourceAdapter(ctx, inputs, envInfo, tokenProvider, this.plugin);
  }

  async configureResource(
    ctx: v2.Context,
    inputs: v2.ProvisionInputs,
    envInfo: v2.EnvInfoV2,
    tokenProvider: TokenProvider
  ): Promise<Result<Void, FxError>> {
    return await configureResourceAdapter(ctx, inputs, envInfo, tokenProvider, this.plugin);
  }
}
