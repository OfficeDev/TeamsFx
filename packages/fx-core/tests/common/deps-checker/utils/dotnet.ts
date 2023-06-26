// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs-extra";
import * as path from "path";
import * as os from "os";
import * as tmp from "tmp";

import { ConfigFolderName } from "@microsoft/teamsfx-api";
import { cpUtils } from "../../../../src/common/deps-checker/util/cpUtils";
import { isWindows } from "../../../../src/common/deps-checker/util/system";
import { createTmpDir } from "./common";
import {
  DotnetChecker,
  DotnetVersion,
} from "../../../../src/common/deps-checker/internal/dotnetChecker";

const find = require("find-process");

tmp.setGracefulCleanup();

export const dotnetConfigPath = path.join(os.homedir(), "." + ConfigFolderName, "dotnet.json");
export const dotnetPrivateInstallPath = path.join(
  os.homedir(),
  "." + ConfigFolderName,
  "bin",
  "dotnet"
);
export const dotnetCommand = "dotnet";
export const dotnetOldVersion = DotnetVersion.v21;
export const dotnetInstallVersion = DotnetVersion.v31;
export const dotnetSupportedVersions = [DotnetVersion.v31, DotnetVersion.v50];

export const testCsprojFileName = "extensions.csproj";
export const testOutputDirName = "bin";

export async function getDotnetExecPathFromConfig(
  dotnetConfigPath: string
): Promise<string | null> {
  try {
    const config = await fs.readJson(dotnetConfigPath, { encoding: "utf-8" });
    if (typeof config.dotnetExecutablePath === "string") {
      return config.dotnetExecutablePath;
    }
  } catch (error) {
    console.debug(`Failed to getDotnetConfig, error = '${error}'`);
  }
  return null;
}

export async function hasDotnetVersion(
  dotnetExecPath: string,
  versionString: string
): Promise<boolean> {
  return await hasAnyDotnetVersions(dotnetExecPath, [versionString]);
}

export async function hasAnyDotnetVersions(
  dotnetExecPath: string,
  versionStrings: string[]
): Promise<boolean> {
  try {
    const output = await cpUtils.executeCommand(
      undefined,
      undefined,
      undefined,
      dotnetExecPath,
      "--list-sdks"
    );
    return output.split(/\r?\n/).some((line: string) => {
      return versionStrings.some((versionString) => line.startsWith(versionString));
    });
  } catch (error) {
    console.debug(`Failed to run "${dotnetExecPath} --list-sdks", error = '${error}'`);
    return false;
  }
}

export async function cleanup() {
  // fs-extra.remove() does nothing if the file does not exist.
  await fs.remove(dotnetConfigPath);
  const processes = await find("name", "dotnet", true);
  processes.forEach((p: { pid: number }, index: number, array: any) =>
    process.kill(p.pid, "SIGKILL")
  );
  await fs.remove(dotnetPrivateInstallPath);
}

export async function withDotnet(
  dotnetChecker: DotnetChecker,
  version: DotnetVersion,
  addToPath: boolean,
  callback: (dotnetExecPath: string) => Promise<void>
): Promise<void> {
  let installDir: string;
  let cleanupCallback: () => void;

  try {
    [installDir, cleanupCallback] = await createTmpDir();
  } catch (error) {
    throw new Error(`Failed to create tmpdir for dotnet, error = '${error}'`);
  }

  const backupPath = process.env.PATH;

  try {
    // use private method as a helper method in test only
    await dotnetChecker["runDotnetInstallScript"](version, installDir);
    const dotnetExecPath = DotnetChecker["getDotnetExecPathFromDotnetInstallationDir"](installDir);

    if (!(await hasDotnetVersion(dotnetExecPath, version))) {
      throw new Error(
        `Failed to install .NET SDK version '${version}' for testing, dotnetExecPath = '${dotnetExecPath}'`
      );
    }

    if (addToPath) {
      process.env.PATH =
        path.resolve(dotnetExecPath, "..") + (isWindows() ? ";" : ":") + process.env.PATH;
    }

    await callback(dotnetExecPath);
  } finally {
    if (addToPath) {
      process.env.PATH = backupPath;
    }
    cleanupCallback();
  }
}

export async function createTmpBackendProjectDir(
  csprojFileName: string
): Promise<[string, () => void]> {
  const [dir, cleanupCallback] = await createTmpDir();

  const csprojPath = path.resolve(
    __dirname,
    "../../../../../../templates/function-base/ts/default/extensions.csproj"
  );
  const targetPath = path.join(dir, csprojFileName);
  await fs.copyFile(csprojPath, targetPath, fs.constants.COPYFILE_EXCL);

  return [dir, cleanupCallback];
}

export async function createMockResourceDir(dirName: string): Promise<[string, () => void]> {
  const [dir, cleanupCallback] = await createTmpDir();

  const resourceDir = path.resolve(__dirname, "../../../../resource/deps-checker");
  const targetDir = path.join(dir, dirName);

  await fs.ensureDir(targetDir);
  await fs.copy(resourceDir, targetDir);

  return [targetDir, cleanupCallback];
}
