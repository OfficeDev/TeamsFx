// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { DeployArgs, DeployContext, DeployStepArgs } from "../../interface/buildAndDeployArgs";
import { BaseComponentInnerError } from "../../../error/componentError";
import ignore, { Ignore } from "ignore";
import { DeployConstant } from "../../../constant/deployConstant";
import * as path from "path";
import * as fs from "fs-extra";
import { zipFolderAsync } from "../../../utils/fileOperation";
import { asBoolean, asFactory, asOptional, asString } from "../../../utils/common";
import { BaseDeployStepDriver } from "../../interface/baseDeployStepDriver";
import { ExecutionResult } from "../../interface/stepDriver";
import { ok, err, UserError, SystemError } from "@microsoft/teamsfx-api";

export abstract class BaseDeployDriver extends BaseDeployStepDriver {
  protected static readonly emptyMap = new Map<string, string>();
  protected helpLink: string | undefined = undefined;
  protected abstract summaries: string[];
  protected abstract summaryPrepare: string[];

  protected static asDeployArgs = asFactory<DeployArgs>({
    workingDirectory: asOptional(asString),
    distributionPath: asString,
    ignoreFile: asOptional(asString),
    resourceId: asString,
    dryRun: asOptional(asBoolean),
  });

  async run(): Promise<ExecutionResult> {
    await this.context.logProvider.debug("start deploy process");

    return await this.wrapErrorHandler(async () => {
      const deployArgs = BaseDeployDriver.asDeployArgs(this.args, this.helpLink);
      // if working directory not set, use current working directory
      deployArgs.workingDirectory = deployArgs.workingDirectory ?? "./";
      // if working dir is not absolute path, then join the path with project path
      this.workingDirectory = path.isAbsolute(deployArgs.workingDirectory)
        ? deployArgs.workingDirectory
        : path.join(this.workingDirectory, deployArgs.workingDirectory);
      // if distribution path is not absolute path, then join the path with project path
      this.distDirectory = path.isAbsolute(deployArgs.distributionPath)
        ? deployArgs.distributionPath
        : path.join(this.workingDirectory, deployArgs.distributionPath);
      this.dryRun = deployArgs.dryRun ?? false;
      // call real deploy
      return await this.deploy(deployArgs);
    });
  }

  /**
   * pack dist folder into zip
   * @param args dist folder and ignore files
   * @param context log provider etc..
   * @protected
   */
  protected async packageToZip(args: DeployStepArgs, context: DeployContext): Promise<Buffer> {
    const ig = await this.handleIgnore(args, context);
    const zipFilePath = path.join(
      this.workingDirectory,
      DeployConstant.DEPLOYMENT_TMP_FOLDER,
      DeployConstant.DEPLOYMENT_ZIP_CACHE_FILE
    );
    await this.context.logProvider?.debug(`start zip dist folder ${this.distDirectory}`);
    const res = await zipFolderAsync(this.distDirectory, zipFilePath, ig);
    await this.context.logProvider?.debug(
      `zip dist folder ${this.distDirectory} to ${zipFilePath} complete`
    );
    return res;
  }

  protected async handleIgnore(args: DeployStepArgs, context: DeployContext): Promise<Ignore> {
    // always add deploy temp folder into ignore list
    const ig = ignore().add(DeployConstant.DEPLOYMENT_TMP_FOLDER);
    if (args.ignoreFile) {
      const ignoreFilePath = path.join(this.workingDirectory, args.ignoreFile);
      if (await fs.pathExists(ignoreFilePath)) {
        const ignoreFileContent = await fs.readFile(ignoreFilePath);
        ignoreFileContent
          .toString()
          .split("\n")
          .map((line) => line.trim())
          .forEach((it) => {
            ig.add(it);
          });
      } else {
        await context.logProvider.warning(
          `already set deploy ignore file ${args.ignoreFile} but file not exists in ${this.workingDirectory}, skip ignore!`
        );
      }
    }
    return ig;
  }

  protected async wrapErrorHandler(fn: () => boolean | Promise<boolean>): Promise<ExecutionResult> {
    try {
      return (await fn())
        ? { result: ok(BaseDeployDriver.emptyMap), summaries: this.summaries }
        : { result: ok(BaseDeployDriver.emptyMap), summaries: this.summaryPrepare };
    } catch (e) {
      await this.context.progressBar?.end(false);
      if (e instanceof BaseComponentInnerError) {
        const errorDetail = e.detail ? `Detail: ${e.detail}` : "";
        await this.context.logProvider.error(`${e.message} ${errorDetail}`);
        return { result: err(e.toFxError()), summaries: [] };
      } else if (e instanceof UserError || e instanceof SystemError) {
        await this.context.logProvider.error(`Error occurred: ${e.message}`);
        return { result: err(e), summaries: [] };
      } else {
        await this.context.logProvider.error(`Unknown error: ${e}`);
        return {
          result: err(BaseComponentInnerError.unknownError("Deploy", e).toFxError()),
          summaries: [],
        };
      }
    }
  }

  /**
   * real deploy process
   * @param args deploy arguments
   */
  abstract deploy(args: DeployArgs): Promise<boolean>;
}
