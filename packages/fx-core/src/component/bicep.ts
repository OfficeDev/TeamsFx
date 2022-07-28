// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  FxError,
  ok,
  Result,
  InputsWithProjectPath,
  ResourceContextV3,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import * as path from "path";
import "reflect-metadata";
import { Service } from "typedi";
import { getProjectTemplatesFolderPath } from "../common/utils";
import { getTemplatesFolder } from "../folder";
import arm from "../plugins/solution/fx-solution/arm";
@Service("bicep")
export class BicepComponent {
  readonly name = "bicep";
  async init(projectPath: string): Promise<Result<undefined, FxError>> {
    const sourceFolder = path.join(getTemplatesFolder(), "bicep");
    const targetFolder = path.join(await getProjectTemplatesFolderPath(projectPath), "azure");
    if (
      (await fs.pathExists(path.join(targetFolder, "main.bicep"))) &&
      (await fs.pathExists(path.join(targetFolder, "provision.bicep"))) &&
      (await fs.pathExists(path.join(targetFolder, "config.bicep")))
    )
      return ok(undefined);
    await fs.ensureDir(targetFolder);
    await fs.ensureDir(path.join(targetFolder, "provision"));
    await fs.ensureDir(path.join(targetFolder, "teamsFx"));
    if (!(await fs.pathExists(path.join(targetFolder, "main.bicep")))) {
      await fs.copyFile(
        path.join(sourceFolder, "main.bicep"),
        path.join(targetFolder, "main.bicep")
      );
    }
    if (!(await fs.pathExists(path.join(targetFolder, "provision.bicep")))) {
      await fs.copyFile(
        path.join(sourceFolder, "provision.bicep"),
        path.join(targetFolder, "provision.bicep")
      );
    }
    if (!(await fs.pathExists(path.join(targetFolder, "config.bicep")))) {
      await fs.copyFile(
        path.join(sourceFolder, "config.bicep"),
        path.join(targetFolder, "config.bicep")
      );
    }
    return ok(undefined);
  }
  async deploy(
    context: ResourceContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<undefined, FxError>> {
    const ctx = context as ResourceContextV3;
    return await arm.deployArmTemplates(
      ctx,
      inputs,
      ctx.envInfo,
      ctx.tokenProvider.azureAccountProvider
    );
  }
}
