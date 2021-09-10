// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConfigFolderName, InputConfigsFolderName } from "@microsoft/teamsfx-api";
import { deserializeDict } from "@microsoft/teamsfx-core";
import { exec } from "child_process";
import fs from "fs-extra";
import os from "os";
import path from "path";
import { promisify } from "util";
import { v4 as uuidv4 } from "uuid";
import { sleep } from "../../src/utils";

import { cfg, AadManager, ResourceGroupManager } from "../commonlib";

export const TEN_MEGA_BYTE = 1024 * 1024 * 10;
export const execAsync = promisify(exec);

export async function execAsyncWithRetry(
  command: string,
  options: {
    cwd?: string;
    env?: NodeJS.ProcessEnv;
    timeout?: number;
  },
  retries = 3,
  newCommand?: string
): Promise<{
  stdout: string;
  stderr: string;
}> {
  while (retries > 0) {
    retries--;
    try {
      const result = await execAsync(command, options);
      return result;
    } catch (e) {
      console.log(`Run \`${command}\` failed with error msg: ${JSON.stringify(e)}.`);
      if (newCommand) {
        command = newCommand;
      }
      await sleep(10000);
    }
  }
  return execAsync(command, options);
}

const testFolder = path.resolve(os.homedir(), "test-folder");

export function getTestFolder() {
  if (!fs.pathExistsSync(testFolder)) {
    fs.mkdirSync(testFolder);
  }
  return testFolder;
}

export function getAppNamePrefix() {
  return "fxE2E";
}

export function getUniqueAppName() {
  return getAppNamePrefix() + Date.now().toString() + uuidv4().slice(0, 2);
}

export function getSubscriptionId() {
  return cfg.AZURE_SUBSCRIPTION_ID || "";
}

const envFilePathSuffix = path.join(".fx", "env.default.json");
const defaultBicepParameterFileSuffix = path.join(
  `.${ConfigFolderName}`,
  InputConfigsFolderName,
  "azure.parameters.dev.json"
);

export function getConfigFileName(appName: string): string {
  return path.resolve(testFolder, appName, envFilePathSuffix);
}

const aadPluginName = "fx-resource-aad-app-for-teams";
const simpleAuthPluginName = "fx-resource-simple-auth";
const botPluginName = "fx-resource-bot";
const apimPluginName = "fx-resource-apim";

export async function setSimpleAuthSkuNameToB1(projectPath: string) {
  const envFilePath = path.resolve(projectPath, envFilePathSuffix);
  const context = await fs.readJSON(envFilePath);
  context[simpleAuthPluginName]["skuName"] = "B1";
  return fs.writeJSON(envFilePath, context, { spaces: 4 });
}

export async function setSimpleAuthSkuNameToB1Bicep(projectPath: string) {
  const parametersFilePath = path.resolve(projectPath, defaultBicepParameterFileSuffix);
  const parameters = await fs.readJSON(parametersFilePath);
  parameters["parameters"]["simpleAuth_sku"] = { value: "B1" };
  return fs.writeJSON(parametersFilePath, parameters, { spaces: 4 });
}

export async function setBotSkuNameToB1(projectPath: string) {
  const envFilePath = path.resolve(projectPath, envFilePathSuffix);
  const context = await fs.readJSON(envFilePath);
  context[botPluginName]["skuName"] = "B1";
  return fs.writeJSON(envFilePath, context, { spaces: 4 });
}

export async function cleanUpAadApp(
  projectPath: string,
  hasAadPlugin?: boolean,
  hasBotPlugin?: boolean,
  hasApimPlugin?: boolean
) {
  const envFilePath = path.resolve(projectPath, envFilePathSuffix);
  const context = await fs.readJSON(envFilePath);
  const manager = await AadManager.init();
  const promises: Promise<boolean>[] = [];

  const clean = async (objectId?: string) => {
    return new Promise<boolean>(async (resolve) => {
      if (objectId) {
        const result = await manager.deleteAadAppById(objectId);
        if (result) {
          console.log(`[Successfully] clean up the Aad app with id: ${objectId}.`);
        } else {
          console.error(`[Failed] clean up the Aad app with id: ${objectId}.`);
        }
        return resolve(result);
      }
      return resolve(false);
    });
  };

  if (hasAadPlugin) {
    const objectId = context[aadPluginName].objectId;
    promises.push(clean(objectId));
  }

  if (hasBotPlugin) {
    const objectId = context[botPluginName].objectId;
    promises.push(clean(objectId));
  }

  if (hasApimPlugin) {
    const objectId = context[apimPluginName].apimClientAADObjectId;
    promises.push(clean(objectId));
  }

  return Promise.all(promises);
}

export async function cleanUpResourceGroup(appName: string) {
  return new Promise<boolean>(async (resolve) => {
    const manager = await ResourceGroupManager.init();
    if (appName) {
      const name = `${appName}-rg`;
      if (await manager.hasResourceGroup(name)) {
        const result = await manager.deleteResourceGroup(name);
        if (result) {
          console.log(`[Successfully] clean up the Azure resource group with name: ${name}.`);
        } else {
          console.error(`[Faild] clean up the Azure resource group with name: ${name}.`);
        }
        return resolve(result);
      }
    }
    return resolve(false);
  });
}

export async function cleanUpLocalProject(projectPath: string, necessary?: Promise<any>) {
  return new Promise<boolean>(async (resolve) => {
    try {
      await necessary;
      await fs.remove(projectPath);
      console.log(`[Successfully] clean up the local folder: ${projectPath}.`);
      return resolve(true);
    } catch {
      console.log(`[Failed] clean up the local folder: ${projectPath}.`);
      return resolve(false);
    }
  });
}

export async function cleanUp(
  appName: string,
  projectPath: string,
  hasAadPlugin = true,
  hasBotPlugin = false,
  hasApimPlugin = false
) {
  const cleanUpAadAppPromise = cleanUpAadApp(
    projectPath,
    hasAadPlugin,
    hasBotPlugin,
    hasApimPlugin
  );
  return Promise.all([
    // delete aad app
    cleanUpAadAppPromise,
    // remove resouce group
    cleanUpResourceGroup(appName),
    // remove project
    cleanUpLocalProject(projectPath, cleanUpAadAppPromise),
  ]);
}

export async function cleanUpResourcesCreatedHoursAgo(
  type: "aad" | "rg",
  contains: string,
  hours?: number,
  retryTimes = 5
) {
  if (type === "aad") {
    const aadManager = await AadManager.init();
    await aadManager.deleteAadApps(contains, hours, retryTimes);
  } else {
    const rgManager = await ResourceGroupManager.init();
    const groups = await rgManager.searchResourceGroups(contains);
    const filteredGroups =
      hours && hours > 0
        ? groups.filter((group) => {
            const name = group.name!;
            const startPos = name.indexOf(contains) + contains.length;
            const createdTime = Number(name.slice(startPos, startPos + 13));
            return Date.now() - createdTime > hours * 3600 * 1000;
          })
        : groups;

    const promises = filteredGroups.map((rg) =>
      rgManager.deleteResourceGroup(rg.name!, retryTimes)
    );
    const results = await Promise.all(promises);
    results.forEach((result, index) => {
      if (result) {
        console.log(
          `[Successfully] clean up the Azure resource group with name: ${filteredGroups[index].name}.`
        );
      } else {
        console.error(
          `[Faild] clean up the Azure resource group with name: ${filteredGroups[index].name}.`
        );
      }
    });
    return results;
  }
}

// TODO: add encrypt
export async function readContext(projectPath: string): Promise<any> {
  const contextFilePath = `${projectPath}/.fx/env.default.json`;
  const userDataFilePath = `${projectPath}/.fx/default.userdata`;

  // Read Context and UserData
  const context = await fs.readJSON(`${projectPath}/.fx/env.default.json`);

  let userData: Record<string, string> = {};
  if (await fs.pathExists(userDataFilePath)) {
    const dictContent = await fs.readFile(userDataFilePath, "UTF-8");
    userData = deserializeDict(dictContent);
  }

  // Read from userdata.
  for (const plugin in context) {
    const pluginContext = context[plugin];
    for (const key in pluginContext) {
      if (typeof pluginContext[key] === "string" && isSecretPattern(pluginContext[key])) {
        const secretKey = `${plugin}.${key}`;
        pluginContext[key] = userData[secretKey] ?? undefined;
      }
    }
  }

  return context;
}

export function mockTeamsfxMultiEnvFeatureFlag() {
  const env = Object.assign(process.env, {});
  env["TEAMSFX_MULTI_ENV"] = "true";
  env["TEAMSFX_ARM_SUPPORT"] = "true";
  env["TEAMSFX_BICEP_ENV_CHECKER_ENABLE"] = "true";
  return env;
}

function isSecretPattern(value: string) {
  console.log(value);
  return value.startsWith("{{") && value.endsWith("}}");
}
