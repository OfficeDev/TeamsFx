// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { DepsCheckerError } from "./depsError";

export interface DepsChecker {
  getInstallationInfo(installOptions?: InstallOptions): Promise<DependencyStatus>;

  resolve(installOptions?: InstallOptions): Promise<DependencyStatus>;
}

export type DependencyStatus = {
  name: string;
  type: DepsType;
  isInstalled: boolean;
  command: string;
  details: {
    isLinuxSupported: boolean;
    supportedVersions: string[];
    installVersion?: string;
    binFolders?: string[];
  };
  telemetryProperties?: { [key: string]: string };
  error?: DepsCheckerError;
};

export enum DepsType {
  LtsNode = "lts-node",
  ProjectNode = "project-node",
  Dotnet = "dotnet",
  FuncCoreTools = "func-core-tools",
  VxTestApp = "vx-test-app",
}

export interface BaseInstallOptions {
  projectPath?: string;
  version?: string;
}

export interface FuncInstallOptions {
  symlinkDir?: string;
  projectPath: string;
  version: string;
}

export type InstallOptions = BaseInstallOptions | FuncInstallOptions;
