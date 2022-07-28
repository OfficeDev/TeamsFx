// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Bicep,
  ContextV3,
  FxError,
  InputsWithProjectPath,
  ok,
  ResourceContextV3,
  ResourceOutputs,
  Result,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import * as path from "path";
import { azureWebSiteDeploy } from "../../../common/azure-hosting/utils";
import * as utils from "../../../plugins/resource/bot/utils/common";
import { getLanguage, getRuntime } from "../../../plugins/resource/bot/v2/mapping";
import {
  CheckThrowSomethingMissing,
  PackDirectoryExistenceError,
  PreconditionError,
} from "./errors";
import { AzureResource } from "./../azureResource";
import { Messages } from "./messages";
import { ProgressMessages, ProgressTitles } from "../../messages";

export abstract class AzureAppService extends AzureResource {
  abstract readonly name: string;
  abstract readonly alias: string;
  abstract readonly displayName: string;
  abstract readonly bicepModuleName: string;
  abstract readonly outputs: ResourceOutputs;
  abstract readonly finalOutputKeys: string[];
  async generateBicep(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<Bicep[], FxError>> {
    this.getTemplateContext = (context, inputs) => {
      const configs: string[] = [];
      configs.push(getRuntime(getLanguage(context.projectSetting.programmingLanguage)));
      this.templateContext.configs = configs;
      return this.templateContext;
    };
    return super.generateBicep(context, inputs);
  }

  async deploy(
    context: ResourceContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<undefined, FxError>> {
    const progressBar = context.userInteraction.createProgressBar(
      ProgressTitles.deploying(this.displayName, inputs.scenario),
      2
    );
    try {
      const ctx = context as ResourceContextV3;
      // Preconditions checking.
      if (!inputs.projectPath || !inputs.artifactFolder) {
        throw new PreconditionError(this.alias, Messages.WorkingDirIsMissing, []);
      }
      const publishDir = path.resolve(inputs.projectPath, inputs.artifactFolder);
      const packDirExisted = await fs.pathExists(publishDir);
      if (!packDirExisted) {
        throw new PackDirectoryExistenceError(this.alias);
      }

      const state = ctx.envInfo.state[inputs.componentId];
      const resourceId = CheckThrowSomethingMissing(
        this.alias,
        this.outputs.resourceId.key,
        state[this.outputs.resourceId.key]
      );
      await progressBar.next(ProgressMessages.packingCode);
      const zipBuffer = await utils.zipFolderAsync(publishDir, "");

      await azureWebSiteDeploy(
        resourceId,
        ctx.tokenProvider,
        zipBuffer,
        context.logProvider,
        progressBar
      );
    } finally {
      progressBar.end(true);
    }
    return ok(undefined);
  }
}
