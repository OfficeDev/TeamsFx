// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Siglud <fanhu@microsoft.com>
 **/

import * as path from "path";
import { BotValidator } from "../../commonlib";
import {
  cleanUp,
  execAsync,
  execAsyncWithRetry,
  getSubscriptionId,
  getTestFolder,
  getUniqueAppName,
  readContextMultiEnv,
  readContextMultiEnvV3,
} from "../commonUtils";
import { environmentManager } from "@microsoft/teamsfx-core/build/core/environment";
import { Runtime, CliCapabilities, CliTriggerType } from "../../commonlib/constants";

export async function happyPathTest(
  runtime: Runtime,
  capabilities: CliCapabilities,
  trigger?: CliTriggerType[]
): Promise<void> {
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const subscription = getSubscriptionId();
  const projectPath = path.resolve(testFolder, appName);
  const envName = environmentManager.getDefaultEnvName();

  const env = Object.assign({}, process.env);
  if (runtime === Runtime.Dotnet) {
    env["TEAMSFX_CLI_DOTNET"] = "true";
  }

  const triggerStr = trigger === undefined ? "" : `--bot-host-type-trigger ${trigger.join(" ")} `;
  const cmdBase = `teamsfx new --interactive false --app-name ${appName} --capabilities ${capabilities} ${triggerStr}`;
  const cmd =
    runtime === Runtime.Dotnet
      ? `${cmdBase} --runtime dotnet`
      : `${cmdBase} --programming-language typescript`;
  console.log(`ready to run CMD: ${cmd}`);
  await execAsync(cmd, {
    cwd: testFolder,
    env: env,
    timeout: 0,
  });
  console.log(`[Successfully] scaffold to ${projectPath}`);

  // set subscription
  // await CliHelper.setSubscription(subscription, projectPath, env);

  console.log(`[Successfully] set subscription for ${projectPath}`);

  // provision
  await execAsyncWithRetry(`teamsfx provision`, {
    cwd: projectPath,
    env: env,
    timeout: 0,
  });

  console.log(`[Successfully] provision for ${projectPath}`);

  {
    // Validate provision
    // Get context
    const context = await readContextMultiEnvV3(projectPath, envName);

    // Validate Bot Provision
    const bot = new BotValidator(context, projectPath, envName);
    await bot.validateProvisionV3(false);
  }

  // deploy
  const cmdStr = "teamsfx deploy";
  await execAsyncWithRetry(cmdStr, {
    cwd: projectPath,
    env: env,
    timeout: 0,
  });
  console.log(`[Successfully] deploy for ${projectPath}`);

  {
    // Validate deployment

    // Get context
    const context = await readContextMultiEnvV3(projectPath, envName);

    // Validate Bot Deploy
    const bot = new BotValidator(context, projectPath, envName);
    await bot.validateDeploy();
  }

  // test (validate)
  await execAsyncWithRetry(`teamsfx validate --env ${envName}`, {
    cwd: projectPath,
    env: env,
    timeout: 0,
  });

  // package
  await execAsyncWithRetry(`teamsfx package --env ${envName}`, {
    cwd: projectPath,
    env: env,
    timeout: 0,
  });

  console.log(`[Successfully] start to clean up for ${projectPath}`);
  await cleanUp(appName, projectPath, false, true, false);
}
