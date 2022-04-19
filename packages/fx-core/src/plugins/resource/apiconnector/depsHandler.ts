// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { Json } from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import * as path from "path";
import semver from "semver";
import { Constants } from "./constants";
import { ResultFactory, FileChange, FileChangeType } from "./result";
import { ErrorMessage } from "./errors";
import { TelemetryUtils, Telemetry } from "./telemetry";
import { getTemplatesFolder } from "../../../folder";
export class DepsHandler {
  private readonly projectRoot: string;
  private readonly componentType: string;
  constructor(workspaceFolder: string, componentType: string) {
    this.projectRoot = workspaceFolder;
    this.componentType = componentType;
  }

  public async addPkgDeps(): Promise<FileChange | undefined> {
    const depsConfig: Json = await this.getDepsConfig();
    return await this.updateLocalPkgDepsVersion(depsConfig);
  }

  public async getDepsConfig(): Promise<Json> {
    const configPath = path.join(getTemplatesFolder(), "plugins", "resource", "apiconnector");
    const sdkConfigPath = path.join(configPath, Constants.pkgJsonFile);
    const sdkContent: Json = await fs.readJson(sdkConfigPath);
    return sdkContent.dependencies;
  }

  public async updateLocalPkgDepsVersion(pkgConfig: Json): Promise<FileChange | undefined> {
    const localPkgPath = path.join(this.projectRoot, this.componentType, Constants.pkgJsonFile);
    if (!(await fs.pathExists(localPkgPath))) {
      throw ResultFactory.UserError(
        ErrorMessage.localPkgFileNotExistError.name,
        ErrorMessage.localPkgFileNotExistError.message(this.componentType)
      );
    }
    const pkgContent = await fs.readJson(localPkgPath);
    let needUpdate = false;
    for (const pkgItem in pkgConfig) {
      if (this.sdkVersionCheck(pkgContent.dependencies, pkgItem, pkgConfig[pkgItem])) {
        pkgContent.dependencies[pkgItem] = pkgConfig[pkgItem];
        needUpdate = true;
      }
    }
    if (needUpdate) {
      await fs.writeFile(localPkgPath, JSON.stringify(pkgContent, null, 4));
      const telemetryProperties = { component: this.componentType };

      TelemetryUtils.sendEvent(Telemetry.stage.updatePkg, undefined, telemetryProperties);
      return {
        changeType: FileChangeType.Update,
        filePath: localPkgPath,
      }; // return modified files
    }
    return undefined;
  }

  private sdkVersionCheck(deps: Json, sdkName: string, sdkVersion: string): boolean {
    // sdk not in dependencies.
    if (!deps[sdkName]) {
      return true;
    }
    // local sdk version intersect with sdk version in config.
    else if (semver.intersects(deps[sdkName], sdkVersion)) {
      return false;
    }
    // local sdk version lager than sdk version in config.
    else if (semver.gt(semver.minVersion(deps[sdkName])!, semver.minVersion(sdkVersion)!)) {
      return false;
    } else {
      throw ResultFactory.UserError(
        ErrorMessage.sdkVersionImcompatibleError.name,
        ErrorMessage.sdkVersionImcompatibleError.message(
          this.componentType,
          deps[sdkName],
          sdkVersion
        )
      );
    }
  }
}
