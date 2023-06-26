// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs-extra";
import * as path from "path";
import semver from "semver";
import {
  nodeNotFoundHelpLink,
  v3NodeNotFoundHelpLink,
  v3NodeNotLtsHelpLink,
  v3NodeNotSupportedHelpLink,
} from "../constant/helpLink";
import { Messages } from "../constant/message";
import { DependencyStatus, DepsChecker, DepsType, BaseInstallOptions } from "../depsChecker";
import {
  DepsCheckerError,
  NodeNotFoundError,
  NodeNotLtsError,
  V3NodeNotSupportedError,
} from "../depsError";
import { cpUtils } from "../util/cpUtils";

const NodeName = "Node.js";

class NodeVersion {
  public readonly version: string;
  public readonly majorVersion: string;

  constructor(version: string, majorVersion: string) {
    this.version = version;
    this.majorVersion = majorVersion;
  }
}

export abstract class NodeChecker implements DepsChecker {
  protected abstract readonly _nodeNotFoundHelpLink: string;
  protected abstract readonly _type: DepsType;
  protected abstract getSupportedVersions(projectPath?: string): Promise<string[]>;
  protected abstract isVersionSupported(supportedVersions: string[], version: NodeVersion): boolean;
  protected abstract getVersionNotSupportedError(
    supportedVersions: string[],
    version: NodeVersion
  ): DepsCheckerError;
  protected abstract readonly _minErrorVersion: number;
  protected abstract readonly _maxErrorVersion: number;

  public async getInstallationInfo(installOptions?: BaseInstallOptions): Promise<DependencyStatus> {
    let supportedVersions: string[] = [];
    try {
      supportedVersions = await this.getSupportedVersions(installOptions?.projectPath);

      const currentVersion = await NodeChecker.getInstalledNodeVersion();
      if (currentVersion === null) {
        const error = new NodeNotFoundError(Messages.NodeNotFound(), this._nodeNotFoundHelpLink);
        return await this.getDepsInfo(false, supportedVersions, undefined, error);
      }

      if (!this.isVersionSupported(supportedVersions, currentVersion)) {
        return await this.getDepsInfo(
          !NodeChecker.isVersionError(this._minErrorVersion, this._maxErrorVersion, currentVersion),
          supportedVersions,
          currentVersion.version,
          this.getVersionNotSupportedError(supportedVersions, currentVersion)
        );
      }
      return await this.getDepsInfo(true, supportedVersions, currentVersion.version);
    } catch (error) {
      return await this.getDepsInfo(
        false,
        supportedVersions,
        undefined,
        new DepsCheckerError(error.message, nodeNotFoundHelpLink)
      );
    }
  }

  public async resolve(installOptions?: BaseInstallOptions): Promise<DependencyStatus> {
    return await this.getInstallationInfo(installOptions);
  }

  public async install(): Promise<void> {
    return Promise.resolve();
  }

  public async getDepsInfo(
    isInstalled: boolean,
    supportedVersions: string[],
    installVersion?: string,
    error?: DepsCheckerError
  ): Promise<DependencyStatus> {
    return {
      name: NodeName,
      type: this._type,
      isInstalled: isInstalled,
      command: await this.command(),
      details: {
        isLinuxSupported: true,
        supportedVersions: supportedVersions,
        installVersion: installVersion,
      },
      error: error,
    };
  }

  private static isVersionError(
    minErrorVersion: number,
    maxErrorVersion: number,
    version: NodeVersion
  ): boolean {
    const majorVersion = Number.parseInt(version.majorVersion);

    return (
      !Number.isInteger(majorVersion) ||
      majorVersion <= minErrorVersion ||
      majorVersion >= maxErrorVersion
    );
  }

  public async command(): Promise<string> {
    return "node";
  }

  public static async getInstalledNodeVersion(): Promise<NodeVersion | null> {
    try {
      const output = await cpUtils.executeCommand(
        undefined,
        undefined,
        undefined,
        "node",
        "--version"
      );
      return getNodeVersion(output);
    } catch (error) {
      return null;
    }
  }
}

function getNodeVersion(output: string): NodeVersion | null {
  const regex = /v(?<major_version>\d+)\.(?<minor_version>\d+)\.(?<patch_version>\d+)/gm;
  const match = regex.exec(output);
  if (!match) {
    return null;
  }

  const majorVersion = match.groups?.major_version;
  if (!majorVersion) {
    return null;
  }

  return new NodeVersion(match[0], majorVersion);
}

export class LtsNodeChecker extends NodeChecker {
  protected readonly _nodeNotFoundHelpLink = v3NodeNotFoundHelpLink;
  protected readonly _type = DepsType.LtsNode;
  protected readonly _minErrorVersion = Number.MIN_SAFE_INTEGER;
  protected readonly _maxErrorVersion = Number.MAX_SAFE_INTEGER;

  protected async getSupportedVersions(): Promise<string[]> {
    return ["16", "18"];
  }

  protected isVersionSupported(supportedVersions: string[], version: NodeVersion): boolean {
    return supportedVersions.includes(version.majorVersion);
  }
  protected getVersionNotSupportedError(
    supportedVersions: string[],
    version: NodeVersion
  ): DepsCheckerError {
    const supportedVersionsString = supportedVersions.map((v) => "v" + v).join(", ");
    return new NodeNotLtsError(
      Messages.NodeNotLts(version.version, supportedVersionsString),
      v3NodeNotLtsHelpLink
    );
  }
}

export class ProjectNodeChecker extends NodeChecker {
  protected readonly _nodeNotFoundHelpLink = v3NodeNotFoundHelpLink;
  protected readonly _type = DepsType.ProjectNode;
  protected readonly _minErrorVersion = Number.MIN_SAFE_INTEGER;
  protected readonly _maxErrorVersion = Number.MAX_SAFE_INTEGER;

  protected async getNodeNotSupportedHelpLink(): Promise<string> {
    return v3NodeNotSupportedHelpLink;
  }

  protected async getSupportedVersions(projectPath?: string): Promise<string[]> {
    if (!projectPath) {
      return [];
    }
    const supportedVersion = await this.getSupportedVersion(projectPath);
    return supportedVersion ? [supportedVersion] : [];
  }

  private async getSupportedVersion(projectPath: string): Promise<string | undefined> {
    try {
      const packageJson = await fs.readJSON(path.join(projectPath, "package.json"));
      const node = packageJson?.engines?.node;
      if (typeof node !== "string") {
        return undefined;
      }
      return node;
    } catch {
      return undefined;
    }
  }

  protected isVersionSupported(supportedVersions: string[], version: NodeVersion): boolean {
    if (supportedVersions.length == 0) {
      return true;
    }
    const supportedVersion = supportedVersions[0];
    return semver.satisfies(version.version, supportedVersion);
  }

  protected getVersionNotSupportedError(
    supportedVersions: string[],
    version: NodeVersion
  ): DepsCheckerError {
    const supportedVersionsString = supportedVersions.join(", ");
    return new V3NodeNotSupportedError(
      Messages.V3NodeNotSupported(version.version, supportedVersionsString),
      v3NodeNotSupportedHelpLink
    );
  }
}
