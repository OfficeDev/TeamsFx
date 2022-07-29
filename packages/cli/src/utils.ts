// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import fs from "fs-extra";
import path from "path";
import { Options } from "yargs";
import chalk from "chalk";
import * as uuid from "uuid";
import * as dotenv from "dotenv";
import {
  OptionItem,
  Question,
  err,
  ok,
  Result,
  FxError,
  ConfigFolderName,
  getSingleOption,
  SingleSelectQuestion,
  MultiSelectQuestion,
  QTreeNode,
  Inputs,
  Platform,
  Colors,
  StatesFolderName,
  EnvNamePlaceholder,
  ProjectSettingsFileName,
  EnvStateFileNameTemplate,
  InputConfigsFolderName,
  SingleSelectConfig,
} from "@microsoft/teamsfx-api";

import { ConfigNotFoundError, UserdataNotFound, EnvUndefined, ReadFileError } from "./error";
import AzureAccountManager from "./commonlib/azureLogin";
import { FeatureFlags, SUPPORTED_SPFX_VERSION } from "./constants";
import {
  environmentManager,
  WriteFileError,
  localSettingsFileName,
  FxCore,
  isSPFxProject,
  PluginNames,
  isValidProject,
  ProjectSettingsHelper,
  LocalEnvManager,
} from "@microsoft/teamsfx-core";
import { WorkspaceNotSupported } from "./cmds/preview/errors";
import CLIUIInstance from "./userInteraction";
import { CliTelemetry } from "./telemetry/cliTelemetry";
import cliLogger from "./commonlib/log";

export type Json = { [_: string]: any };

export function getChoicesFromQTNodeQuestion(data: Question): string[] | undefined {
  const option = "staticOptions" in data ? data.staticOptions : undefined;
  if (option && option instanceof Array && option.length > 0) {
    if (typeof option[0] === "string") {
      return option as string[];
    } else {
      return (option as OptionItem[]).map((op) => op.cliName || toLocaleLowerCase(op.id));
    }
  } else {
    return undefined;
  }
}

export function getSingleOptionString(
  q: SingleSelectQuestion | MultiSelectQuestion
): string | string[] {
  const singleOption = getSingleOption(q);
  if (q.returnObject) {
    if (q.type === "singleSelect") {
      return typeof singleOption === "string" ? singleOption : singleOption.id;
    } else {
      return [singleOption[0].id];
    }
  } else {
    return singleOption;
  }
}

export function toYargsOptions(data: Question): Options {
  const choices = getChoicesFromQTNodeQuestion(data);

  let defaultValue;
  if (data.default && data.default instanceof Array && data.default.length > 0) {
    defaultValue = data.default.map((item) => item.toLocaleLowerCase());
  } else if (data.default && typeof data.default === "string") {
    defaultValue = data.default.toLocaleLowerCase();
  } else {
    defaultValue = undefined;
  }
  if (defaultValue === undefined) {
    return {
      array: data.type === "multiSelect",
      description: data.title || "",
      choices: choices,
      hidden: !!(data as any).hide,
      global: false,
      type: "string",
    };
  }
  return {
    array: data.type === "multiSelect",
    description: data.title || "",
    default: defaultValue,
    choices: choices,
    hidden: !!(data as any).hide,
    global: false,
    type: "string",
  };
}

export function toLocaleLowerCase(arg: any): any {
  if (typeof arg === "string") {
    return arg.toLocaleLowerCase();
  } else if (arg instanceof Array) {
    return arg.map((s: string) => s.toLocaleLowerCase());
  } else return arg;
}

export function flattenNodes(node: QTreeNode): QTreeNode[] {
  const nodeCopy = Object.assign({}, node);
  const children = (nodeCopy.children || []).concat([]);
  nodeCopy.children = undefined;
  return [nodeCopy].concat(...children.map((nd) => flattenNodes(nd)));
}

export async function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export function getConfigPath(projectFolder: string, filePath: string): string {
  return path.resolve(projectFolder, `.${ConfigFolderName}`, filePath);
}

// TODO: move config read/write utils to core
export function getEnvFilePath(
  projectFolder: string,
  env: string | undefined
): Result<string, FxError> {
  if (!env) {
    return err(new EnvUndefined());
  }
  return ok(
    path.join(
      projectFolder,
      `.${ConfigFolderName}`,
      StatesFolderName,
      EnvStateFileNameTemplate.replace(EnvNamePlaceholder, env)
    )
  );
}

export function getSettingsFilePath(projectFolder: string) {
  return path.join(
    projectFolder,
    `.${ConfigFolderName}`,
    InputConfigsFolderName,
    ProjectSettingsFileName
  );
}

export function getSecretFilePath(
  projectRoot: string,
  env: string | undefined
): Result<string, FxError> {
  return ok(path.join(projectRoot, `.${ConfigFolderName}`, StatesFolderName, `${env}.userdata`));
}

export async function readEnvJsonFile(
  projectFolder: string,
  env: string | undefined
): Promise<Result<Json, FxError>> {
  const filePathResult = getEnvFilePath(projectFolder, env);
  if (filePathResult.isErr()) {
    return err(filePathResult.error);
  }
  const filePath = filePathResult.value;
  if (!fs.existsSync(filePath)) {
    return err(ConfigNotFoundError(filePath));
  }
  try {
    const config = await fs.readJson(filePath);
    return ok(config);
  } catch (e) {
    return err(ReadFileError(e));
  }
}

export function readEnvJsonFileSync(
  projectFolder: string,
  env: string | undefined
): Result<Json, FxError> {
  const filePathResult = getEnvFilePath(projectFolder, env);
  if (filePathResult.isErr()) {
    return err(filePathResult.error);
  }
  const filePath = filePathResult.value;
  if (!fs.existsSync(filePath)) {
    return err(ConfigNotFoundError(filePath));
  }
  try {
    const config = fs.readJsonSync(filePath);
    return ok(config);
  } catch (e) {
    return err(ReadFileError(e));
  }
}

export function readLocalSettingsJsonFile(projectFolder: string): Result<Json, FxError> {
  const localSettingsPath = path.join(
    projectFolder,
    `.${ConfigFolderName}`,
    `${InputConfigsFolderName}`,
    localSettingsFileName
  );
  if (!fs.existsSync(localSettingsPath)) {
    return err(ConfigNotFoundError(localSettingsPath));
  }
  try {
    const config = fs.readJsonSync(localSettingsPath);
    return ok(config);
  } catch (e) {
    return err(ReadFileError(e));
  }
}

export function readLocalStateJsonFile(projectFolder: string): Result<Json, FxError> {
  const localStatePath = path.join(
    projectFolder,
    `.${ConfigFolderName}`,
    `${StatesFolderName}`,
    "state.local.json"
  );
  if (!fs.existsSync(localStatePath)) {
    return err(ConfigNotFoundError(localStatePath));
  }
  try {
    const config = fs.readJsonSync(localStatePath);
    return ok(config);
  } catch (e) {
    return err(ReadFileError(e));
  }
}

export function readSettingsFileSync(projectFolder: string): Result<Json, FxError> {
  const filePath = getSettingsFilePath(projectFolder);
  if (!fs.existsSync(filePath)) {
    return err(ConfigNotFoundError(filePath));
  }

  try {
    const settings = fs.readJsonSync(filePath);
    return ok(settings);
  } catch (e) {
    return err(ReadFileError(e));
  }
}

export async function readProjectSecrets(
  projectFolder: string,
  env: string | undefined
): Promise<Result<dotenv.DotenvParseOutput, FxError>> {
  const secretFileResult = getSecretFilePath(projectFolder, env);
  if (secretFileResult.isErr()) {
    return err(secretFileResult.error);
  }
  const secretFile = secretFileResult.value;
  if (!fs.existsSync(secretFile)) {
    return err(new UserdataNotFound(env!));
  }
  try {
    const secretData = await fs.readFile(secretFile);
    return ok(dotenv.parse(secretData));
  } catch (e) {
    return err(ReadFileError(e));
  }
}

export function writeSecretToFile(
  secrets: dotenv.DotenvParseOutput,
  rootFolder: string,
  env: string | undefined
): Result<null, FxError> {
  const secretFileResult = getSecretFilePath(rootFolder, env);
  if (secretFileResult.isErr()) {
    return err(secretFileResult.error);
  }
  const secretFile = secretFileResult.value;
  const array: string[] = [];
  for (const secretKey of Object.keys(secrets)) {
    const secretValue = secrets[secretKey];
    array.push(`${secretKey}=${secretValue}`);
  }
  if (!fs.existsSync(secretFile)) {
    return err(new UserdataNotFound(env!));
  }
  try {
    fs.writeFileSync(secretFile, array.join("\n"));
  } catch (e) {
    return err(WriteFileError(e));
  }
  return ok(null);
}

export async function setSubscriptionId(
  subscriptionId?: string,
  rootFolder = "./"
): Promise<Result<null, FxError>> {
  if (subscriptionId) {
    if (!isWorkspaceSupported(rootFolder)) {
      return err(WorkspaceNotSupported(rootFolder));
    }

    AzureAccountManager.setRootPath(rootFolder);
    if (subscriptionId) {
      await AzureAccountManager.setSubscription(subscriptionId);
    }
  }
  return ok(null);
}

export function isWorkspaceSupported(workspace: string): boolean {
  const p = workspace;

  const checklist: string[] = [
    p,
    `${p}/.${ConfigFolderName}`,
    path.join(p, `.${ConfigFolderName}`, InputConfigsFolderName, ProjectSettingsFileName),
  ];

  for (const fp of checklist) {
    if (!fs.existsSync(path.resolve(fp))) {
      return false;
    }
  }
  return true;
}

export interface TeamsAppTelemetryInfo {
  appId: string;
  tenantId: string;
}

export function getTeamsAppTelemetryInfoByEnv(
  projectDir: string,
  env: string
): TeamsAppTelemetryInfo | undefined {
  try {
    if (isWorkspaceSupported(projectDir)) {
      const result = environmentManager.getEnvStateFilesPath(env, projectDir);
      const envJson = JSON.parse(fs.readFileSync(result.envState, "utf8"));
      const appstudioState = envJson[PluginNames.APPST];
      return {
        appId: appstudioState.teamsAppId,
        tenantId: appstudioState.tenantId,
      };
    }
  } catch (e) {
    return undefined;
  }
}

/**
 * Ask user to select environment, local is included
 */
export async function askTargetEnvironment(projectDir: string): Promise<Result<string, FxError>> {
  const envProfilesResult = await environmentManager.listRemoteEnvConfigs(projectDir);
  if (envProfilesResult.isErr()) {
    return err(envProfilesResult.error);
  }
  const config: SingleSelectConfig = {
    name: "targetEnvName",
    title: "Select an environment",
    options: envProfilesResult.value.concat([environmentManager.getLocalEnvName()]),
  };
  const selectedEnv = await CLIUIInstance.selectOption(config);
  if (selectedEnv.isErr()) {
    return err(selectedEnv.error);
  } else {
    return ok(selectedEnv.value.result as string);
  }
}

// Only used for telemetry
export function getSettingsVersion(rootFolder: string | undefined): string | undefined {
  if (!rootFolder) {
    return undefined;
  }
  try {
    if (isWorkspaceSupported(rootFolder)) {
      const result = readSettingsFileSync(rootFolder);
      if (result.isOk()) {
        return result.value.version;
      }
    }
  } catch (e) {
    // ignore errors for telemetry
  }
  return undefined;
}

// Only used for telemetry
export function getIsM365(rootFolder: string | undefined): string | undefined {
  if (!rootFolder) {
    return undefined;
  }
  try {
    if (isWorkspaceSupported(rootFolder)) {
      const result = readSettingsFileSync(rootFolder);
      if (result.isOk() && result.value.isM365 !== undefined) {
        return `${result.value.isM365}`;
      }
    }
  } catch (e) {
    // ignore errors for telemetry
  }
  return undefined;
}

export function getLocalTeamsAppId(rootfolder: string | undefined): any {
  if (!rootfolder) {
    return undefined;
  }

  if (isWorkspaceSupported(rootfolder)) {
    const result = readLocalStateJsonFile(rootfolder);
    if (result.isErr()) {
      return undefined;
    }
    const localState = result.value;
    try {
      return localState[PluginNames.APPST].teamsAppId;
    } catch (error) {
      return undefined;
    }
  }

  return undefined;
}

export function getProjectId(rootfolder: string | undefined): any {
  if (!rootfolder) {
    return undefined;
  }

  // Do not check validity of project in multi-env.
  // Before migration, `isWorkspaceSupported()` is false, but we still need to send `project-id` telemetry property.
  const result = readSettingsFileSync(rootfolder);
  if (result.isOk()) {
    return result.value.projectId;
  }

  // Also try reading from the old project location to support `ProjectMigratorMW` telemetry.
  // While doing migration, sending telemetry will call this `getProjectId()` function.
  // But before migration done, the settings file is still in the old location.
  const settingsFilePathOld = getConfigPath(rootfolder, "settings.json");
  try {
    const settings = fs.readJsonSync(settingsFilePathOld);
    return settings.projectId;
  } catch (e) {
    return undefined;
  }
  return undefined;
}

export function getSystemInputs(projectPath?: string, env?: string): Inputs {
  const systemInputs: Inputs = {
    platform: Platform.CLI,
    projectPath: projectPath,
    correlationId: uuid.v4(),
    env: env,
  };
  return systemInputs;
}

export function argsToInputs(
  params: { [_: string]: Options },
  args: { [argName: string]: string | string[] }
): Inputs {
  const inputs = getSystemInputs();
  for (const name in params) {
    if (name.endsWith("folder") && args[name]) {
      inputs[name] = path.resolve(args[name] as string);
    } else {
      inputs[name] = args[name];
    }
  }
  const rootFolder = path.resolve((inputs["folder"] as string) || "./");
  delete inputs["folder"];
  inputs.projectPath = rootFolder;
  return inputs;
}

export function getColorizedString(message: Array<{ content: string; color: Colors }>): string {
  // Color support is automatically detected by chalk
  const colorizedMessage = message
    .map((item) => {
      switch (item.color) {
        case Colors.BRIGHT_WHITE:
          return chalk.whiteBright(item.content);
        case Colors.WHITE:
          return chalk.white(item.content);
        case Colors.BRIGHT_MAGENTA:
          return chalk.magentaBright(item.content);
        case Colors.BRIGHT_GREEN:
          return chalk.greenBright(item.content);
        case Colors.BRIGHT_RED:
          return chalk.redBright(item.content);
        case Colors.BRIGHT_YELLOW:
          return chalk.yellowBright(item.content);
        case Colors.BRIGHT_CYAN:
          return chalk.cyanBright.underline(item.content);
        default:
          return item.content;
      }
    })
    .join("");
  return colorizedMessage + (process.stdout.isTTY ? "\u00A0\u001B[K" : "");
}

/**
 * Shows in `teamsfx -v`.
 * @returns the version of teamsfx-cli.
 */
export function getVersion(): string {
  const pkgPath = path.resolve(__dirname, "..", "package.json");
  const pkgContent = fs.readJsonSync(pkgPath);
  return pkgContent.version;
}

// Determine whether feature flag is enabled based on environment variable setting
export function isFeatureFlagEnabled(featureFlagName: string, defaultValue = false): boolean {
  const flag = process.env[featureFlagName];
  if (flag === undefined) {
    return defaultValue; // allows consumer to set a default value when environment variable not set
  } else {
    return flag === "1" || flag.toLowerCase() === "true"; // can enable feature flag by set environment variable value to "1" or "true"
  }
}

export function isRemoteCollaborationEnabled(): boolean {
  return isFeatureFlagEnabled(FeatureFlags.InsiderPreview, true);
}

export function getAllFeatureFlags(): string[] | undefined {
  const result = Object.values(FeatureFlags)
    .filter((featureFlag) => {
      return isFeatureFlagEnabled(featureFlag);
    })
    .map((featureFlag) => {
      return featureFlag;
    });

  return result;
}

export async function isSpfxProject(
  rootFolder: string,
  core: FxCore
): Promise<Result<boolean | undefined, FxError>> {
  const inputs: Inputs = {
    platform: Platform.CLI,
    projectPath: rootFolder,
  };

  const configResult = await core.getProjectConfig(inputs);
  if (configResult.isErr()) {
    return err(configResult.error);
  }
  const config = configResult.value;
  const projectSettings = config?.settings;
  return ok(isSPFxProject(projectSettings));
}

export async function promptSPFxUpgrade(rootFolder: string) {
  const localEnvManager = new LocalEnvManager(cliLogger, CliTelemetry.getReporter());
  const projectSettings = await localEnvManager.getProjectSettings(rootFolder);
  const isSpfx = ProjectSettingsHelper.isSpfx(projectSettings);
  if (isSpfx) {
    let projectSPFxVersion = null;
    // eslint-disable-next-line @typescript-eslint/no-non-null-assertion, @typescript-eslint/no-non-null-asserted-optional-chain
    const yoInfoPath = path.join(rootFolder, "SPFx", ".yo-rc.json");
    if (await fs.pathExists(yoInfoPath)) {
      const yoInfo = await fs.readJson(yoInfoPath);
      projectSPFxVersion = yoInfo["@microsoft/generator-sharepoint"]?.version;
    }

    if (!projectSPFxVersion) {
      // eslint-disable-next-line @typescript-eslint/no-non-null-assertion, @typescript-eslint/no-non-null-asserted-optional-chain
      const packagePath = path.join(rootFolder, "SPFx", "package.json");
      if (await fs.pathExists(packagePath)) {
        const packageInfo = await fs.readJSON(packagePath);
        projectSPFxVersion = packageInfo.dependencies["@microsoft/sp-webpart-base"];
      }
    }

    if (projectSPFxVersion) {
      const cmp = compare(projectSPFxVersion, SUPPORTED_SPFX_VERSION);
      if (cmp === 1 || cmp === -1) {
        CLIUIInstance.showMessage(
          "warn",
          cmp === 1
            ? `You are using a newer version of SPFx in your project while the current version of TeamsFx CLI supports SPFx v${SUPPORTED_SPFX_VERSION}. Please upgrade to the latest version of TeamsFx CLI.`
            : `You are using a legacy version of SPFx in your project while the current version of TeamsFx CLI supports SPFx v${SUPPORTED_SPFX_VERSION}. Please upgrade your project using 'CLI for Microsoft 365'("https://pnp.github.io/cli-microsoft365/cmd/spfx/project/project-upgrade/") to continue.`,
          false
        );
      }
    }
  }
}

export function compare(v1: string | any, v2: string | any) {
  if (typeof v1 === "string") {
    v1 = fromString(v1);
  }
  if (typeof v2 === "string") {
    v2 = fromString(v2);
  }

  if (v1.major > v2.major) return 1;
  if (v1.major < v2.major) return -1;

  if (v1.minor > v2.minor) return 1;
  if (v1.minor < v2.minor) return -1;

  if (v1.patch > v2.patch) return 1;
  if (v1.patch < v2.patch) return -1;

  if (v1.pre === undefined && v2.pre !== undefined) return 1;
  if (v1.pre !== undefined && v2.pre === undefined) return -1;

  if (v1.pre !== undefined && v2.pre !== undefined) {
    return v1.pre.localeCompare(v2.pre);
  }

  return 0;
}

export function fromString(version: string): any {
  const [ver, pre] = version.split("-");
  const [major, minor, patch] = ver.split(".");
  return {
    major: typeof major === "string" ? parseInt(major, 10) : major,
    minor: typeof minor === "string" ? parseInt(minor, 10) : minor,
    patch: patch == null ? 0 : typeof patch === "string" ? parseInt(patch, 10) : patch,
    pre: pre,
  };
}
