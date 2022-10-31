// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { DeployArgs, DeployContext, DeployStepArgs } from "../../interface/buildAndDeployArgs";
import { BaseComponentInnerError } from "../../../error/componentError";
import ignore, { Ignore } from "ignore";
import { DeployConstant } from "../../../constant/deployConstant";
import * as path from "path";
import * as fs from "fs-extra";
import { zipFolderAsync } from "../../../utils/fileOperation";
import { asFactory, asOptional, asString } from "../../../utils/common";
import { BaseDeployStepDriver } from "../../interface/baseDeployStepDriver";

export abstract class BaseDeployDriver extends BaseDeployStepDriver {
  protected static readonly emptyMap = new Map<string, string>();

  protected static asDeployArgs = asFactory<DeployArgs>({
    workingDirectory: asString,
    distributionPath: asString,
    ignoreFile: asOptional(asString),
    resourceId: asString,
  });

  async run(): Promise<Map<string, string>> {
    await this.context.logProvider.debug("start deploy process");

    const deployArgs = BaseDeployDriver.asDeployArgs(this.args);
    this.workingDirectory = path.join(this.workingDirectory, deployArgs.workingDirectory);
    // call real deploy
    await this.wrapErrorHandler(async () => {
      await this.deploy(deployArgs);
    });
    return BaseDeployDriver.emptyMap;
  }

  /**
   * pack dist folder into zip
   * @param args dist folder and ignore files
   * @param context log provider etc..
   * @protected
   */
  protected async packageToZip(args: DeployStepArgs, context: DeployContext): Promise<Buffer> {
    const ig = await this.handleIgnore(args, context);
    const source = path.join(args.workingDirectory, args.distributionPath);
    const zipFilePath = path.join(
      source,
      DeployConstant.DEPLOYMENT_TMP_FOLDER,
      DeployConstant.DEPLOYMENT_ZIP_CACHE_FILE
    );
    await this.context.logProvider?.debug(`start zip dist folder ${source}`);
    return await zipFolderAsync(source, zipFilePath, ig);
  }

  protected async handleIgnore(args: DeployStepArgs, context: DeployContext): Promise<Ignore> {
    // always add deploy temp folder into ignore list
    const ig = ignore().add(DeployConstant.DEPLOYMENT_TMP_FOLDER);
    if (args.ignoreFile) {
      const ignoreFilePath = path.join(args.workingDirectory, args.ignoreFile);
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
          `already set deploy ignore file ${args.ignoreFile} but file not exists in ${args.workingDirectory}, skip ignore!`
        );
      }
    }
    return ig;
  }

  protected async wrapErrorHandler<T>(fn: () => T | Promise<T>): Promise<T> {
    try {
      return Promise.resolve(fn());
    } catch (e) {
      await this.context.progressBar?.end(false);
      if (e instanceof BaseComponentInnerError) {
        const errorDetail = e.detail ? `Detail: ${e.detail}` : "";
        await this.context.logProvider.error(`${e.message} ${errorDetail}`);
        throw e.toFxError();
      } else {
        await this.context.logProvider.error(`Unknown error: ${e}`);
        throw e;
      }
    }
  }

  /**
   * real deploy process
   * @param args deploy arguments
   */
  abstract deploy(args: DeployArgs): Promise<void>;
}
