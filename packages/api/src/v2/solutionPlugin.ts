// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { Result } from "neverthrow";
import {
  FunctionRouter,
  FxError,
  Inputs,
  QTreeNode,
  Stage,
  TokenProvider,
  Func,
  Json,
  Void,
  AzureAccountProvider,
  AppStudioTokenProvider,
} from "../index";
import { LocalProvisionOutput, ProvisionOutput } from "./resourcePlugin";
import { Context, PluginName } from "./types";

// Will check this with chaoyi.
export type ResourceTempalte = unknown;

export interface SolutionPlugin {
  name: string;

  displayName: string;

  /**
   * Called by Toolkit when creating a new project or adding a new resource.
   * Scaffolds source code on disk, relative to context.projectPath.
   *
   * @example
   * ```
   * scaffoldSourceCode(ctx: Context, inputs: Inputs) {
   *   const fs = require("fs-extra");
   *   let content = "let x = 1;"
   *   let path = path.join(ctx.projectPath, "myFolder");
   *   let sourcePath = "somePathhere";
   *   let result = await fs.copy(sourcePath, content);
   *   // no output values
   *   return { "output": {} };
   * }
   * ```
   *
   * @param {Context} ctx - plugin's runtime context shared by all lifecycles.
   * @param {Inputs} inputs - User answers to quesions defined in {@link getQuestionsForLifecycleTask}
   * for {@link Stage.create} along with some system inputs.
   *
   * @returns scaffold output values for each plugin, which will be persisted by the Toolkit and available to other plugins for other lifecyles.
   */
  scaffoldSourceCode?: (
    ctx: Context,
    inputs: Inputs
  ) => Promise<Result<Record<PluginName, { output: Record<string, string> }>, FxError>>;

  /**
   * Called when creating a new project or adding a new resource.
   * Returns resource templates (e.g. Bicep templates/plain JSON) for provisioning.
   *
   * @param {Context} ctx - plugin's runtime context shared by all lifecycles.
   * @param {Inputs} inputs - User's answers to quesions defined in {@link getQuestionsForLifecycleTask}
   * for {@link Stage.create} along with some system inputs.
   *
   * @return {@link ResourceTemplate} for provisioning and deployment.
   */
  generateResourceTemplate: (
    ctx: Context,
    inputs: Inputs
  ) => Promise<Result<ResourceTempalte, FxError>>;

  /**
   * This method is called by the Toolkit when users run "Provision in the Cloud" command.
   * The implementation of solution is expected to do these operations in order:
   * 1) Call resource plugins' provisionResource.
   * 2) Run Bicep/ARM deployment returned by {@link generateResourceTemplate}.
   * 3) Call resource plugins' configureResource.
   *
   * @param {Context} ctx - plugin's runtime context shared by all lifecycles.
   * @param {Json} provisionTemplate - provision template
   * @param {TokenProvider} tokenProvider - Tokens for Azure and AppStudio
   *
   * @returns the config, project state, secrect values for the current environment. Toolkit will persist them
   *          and pass them to {@link configureResource}.
   */
  provisionResources: (
    ctx: Context,
    inputs: Inputs,
    tokenProvider: TokenProvider
  ) => Promise<Result<Record<PluginName, ProvisionOutput>, FxError>>;

  /**
   * Depends on the values returned by {@link provisionResources}.
   * Expected behavior is to deploy code to cloud using credentials provided by {@link AzureAccountProvider}.
   *
   * @param {Context} ctx - plugin's runtime context shared by all lifecycles.
   * @param {Readonly<ProvisionOutput>} provisionTemplate - output generated during provision
   * @param {AzureAccountProvider} tokenProvider - Tokens for Azure and AppStudio
   *
   * @returns deployment output values for each plugin, which will be persisted by the Toolkit and available to other plugins for other lifecyles.
   */
  deploy?: (
    ctx: Context,
    provisionOutput: Readonly<ProvisionOutput>,
    tokenProvider: AzureAccountProvider
  ) => Promise<Result<Record<PluginName, { output: Record<string, string> }>, FxError>>;

  /**
   * Depends on the output of {@link package}. Uploads Teams package to AppStudio
   * @param {Context} ctx - plugin's runtime context shared by all lifecycles.
   * @param {AppStudioTokenProvider} tokenProvider - Token for AppStudio
   * @param {Inputs} inputs - User answers to quesions defined in {@link getQuestionsForLifecycleTask}
   * for {@link Stage.publish} along with some system inputs.
   *
   * @returns Void because side effect is expected.
   */
  publishApplication?: (
    ctx: Context,
    tokenProvider: AppStudioTokenProvider,
    inputs: Inputs
  ) => Promise<Result<Void, FxError>>;

  /**
   * Generates a Teams manifest package for the current project,
   * and stores it on disk.
   *
   * @param {Context} ctx - plugin's runtime context shared by all lifecycles.
   * @param {Inputs} inputs - User answers to quesions defined in {@link getQuestionsForLifecycleTask}
   * for {@link Stage.package} along with some system inputs.
   *
   * @returns Void because side effect is expected.
   */
  package?: (ctx: Context, inputs: Inputs) => Promise<Result<Void, FxError>>;

  /**
   * provisionLocalResource is a special lifecycle, called when users press F5 in vscode.
   * It works like provision, but only creates necessary cloud resources for local debugging like AAD and AppStudio App.
   * Implementation of this lifecycle is expected to call each resource plugins' provisionLocalResource, and after all of
   * them finishes, call configureLocalResource of each plugin.
   *
   * @param {Context} ctx - plugin's runtime context shared by all lifecycles.
   * @param {TokenProvider} tokenProvider - Tokens for Azure and AppStudio
   *
   * @returns the output values, project state, secrect values for each plugin
   */
  provisionLocalResource?: (
    ctx: Context,
    tokenProvider: TokenProvider
  ) => Promise<Result<Record<PluginName, LocalProvisionOutput>, FxError>>;

  /**
   * get question model for lifecycle {@link Stage} (create, provision, deploy, publish), Questions are organized as a tree. Please check {@link QTreeNode}.
   */
  getQuestionsForLifecycleTask: (
    task: Stage,
    inputs: Inputs,
    ctx?: Context
  ) => Promise<Result<QTreeNode | undefined, FxError>>;

  /**
   * get question model for plugin customized {@link Task}, Questions are organized as a tree. Please check {@link QTreeNode}.
   */
  getQuestionsForUserTask?: (
    router: FunctionRouter,
    inputs: Inputs,
    ctx?: Context
  ) => Promise<Result<QTreeNode | undefined, FxError>>;
  /**
   * execute user customized task, for example `Add Resource`, `Add Capabilities`, etc
   */
  executeUserTask?: (
    func: Func,
    inputs: Inputs,
    ctx?: Context
  ) => Promise<Result<unknown, FxError>>;
}
