// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assign, CommentArray, CommentJSONValue, CommentObject, parse } from "comment-json";
import { DebugMigrationContext } from "./debugMigrationContext";
import {
  defaultNpmInstallArg,
  FolderName,
  Prerequisite,
  TaskCommand,
  TaskDefaultValue,
  TaskLabel,
} from "../../../../common/local";
import {
  createResourcesTask,
  generateLabel,
  isCommentArray,
  isCommentObject,
  OldProjectSettingsHelper,
  setUpLocalProjectsTask,
  updateLocalEnv,
} from "./debugV3MigrationUtils";
import { InstallToolArgs } from "../../../../component/driver/prerequisite/interfaces/InstallToolArgs";
import { BuildArgs } from "../../../../component/driver/interface/buildAndDeployArgs";
import { LocalCrypto } from "../../../crypto";

export async function migrateTransparentPrerequisite(
  context: DebugMigrationContext
): Promise<void> {
  for (const task of context.tasks) {
    if (
      !isCommentObject(task) ||
      !(task["type"] === "teamsfx") ||
      !(task["command"] === TaskCommand.checkPrerequisites)
    ) {
      continue;
    }

    if (isCommentObject(task["args"]) && isCommentArray(task["args"]["prerequisites"])) {
      const newPrerequisites: string[] = [];
      const toolsArgs: InstallToolArgs = {};

      for (const prerequisite of task["args"]["prerequisites"]) {
        if (prerequisite === Prerequisite.nodejs) {
          newPrerequisites.push(`"${Prerequisite.nodejs}", // Validate if Node.js is installed.`);
        } else if (prerequisite === Prerequisite.m365Account) {
          newPrerequisites.push(
            `"${Prerequisite.m365Account}", // Sign-in prompt for Microsoft 365 account, then validate if the account enables the sideloading permission.`
          );
        } else if (prerequisite === Prerequisite.portOccupancy) {
          newPrerequisites.push(
            `"${Prerequisite.portOccupancy}", // Validate available ports to ensure those debug ones are not occupied.`
          );
        } else if (prerequisite === Prerequisite.func) {
          toolsArgs.func = true;
        } else if (prerequisite === Prerequisite.devCert) {
          toolsArgs.devCert = { trust: true };
        } else if (prerequisite === Prerequisite.dotnet) {
          toolsArgs.dotnet = true;
        }
      }

      task["args"]["prerequisites"] = parse(`[
        ${newPrerequisites.join("\n  ")}
      ]`);
      if (Object.keys(toolsArgs).length > 0) {
        if (!context.appYmlConfig.deploy) {
          context.appYmlConfig.deploy = {};
        }
        context.appYmlConfig.deploy.tools = toolsArgs;
      }
    }
  }
}

export async function migrateTransparentLocalTunnel(context: DebugMigrationContext): Promise<void> {
  for (const task of context.tasks) {
    if (
      !isCommentObject(task) ||
      !(task["type"] === "teamsfx") ||
      !(task["command"] === TaskCommand.startLocalTunnel)
    ) {
      continue;
    }

    if (isCommentObject(task["args"])) {
      const comment = `
        {
          // Keep consistency with migrated configuration.
        }
      `;
      task["args"]["env"] = "local";
      task["args"]["output"] = assign(parse(comment), {
        endpoint: context.placeholderMapping.botEndpoint,
        domain: context.placeholderMapping.botDomain,
      });
    }
  }
}

export async function migrateTransparentNpmInstall(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "teamsfx") ||
      !(task["command"] === TaskCommand.npmInstall)
    ) {
      ++index;
      continue;
    }

    if (isCommentObject(task["args"]) && isCommentArray(task["args"]["projects"])) {
      for (const npmArgs of task["args"]["projects"]) {
        if (!isCommentObject(npmArgs) || !(typeof npmArgs["cwd"] === "string")) {
          continue;
        }
        const npmInstallArg: BuildArgs = { args: "install" };
        npmInstallArg.workingDirectory = npmArgs["cwd"].replace("${workspaceFolder}", ".");

        if (typeof npmArgs["npmInstallArgs"] === "string") {
          npmInstallArg.args = `install ${npmArgs["npmInstallArgs"]}`;
        } else if (
          isCommentArray(npmArgs["npmInstallArgs"]) &&
          npmArgs["npmInstallArgs"].length > 0
        ) {
          npmInstallArg.args = `install ${npmArgs["npmInstallArgs"].join(" ")}`;
        }

        if (!context.appYmlConfig.deploy) {
          context.appYmlConfig.deploy = {};
        }
        if (!context.appYmlConfig.deploy.npmCommands) {
          context.appYmlConfig.deploy.npmCommands = [];
        }
        context.appYmlConfig.deploy.npmCommands.push(npmInstallArg);
      }
    }

    if (typeof task["label"] === "string") {
      // TODO: remove preLaunchTask in launch.json
      replaceInDependsOn(task["label"], context.tasks);
    }
    context.tasks.splice(index, 1);
  }
}

export async function migrateSetUpTab(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "teamsfx") ||
      !(task["command"] === TaskCommand.setUpTab)
    ) {
      ++index;
      continue;
    }

    if (typeof task["label"] !== "string") {
      ++index;
      continue;
    }

    let url = new URL("https://localhost:53000");
    if (isCommentObject(task["args"]) && typeof task["args"]["baseUrl"] === "string") {
      try {
        url = new URL(task["args"]["baseUrl"]);
      } catch {}
    }

    if (!context.appYmlConfig.configureApp) {
      context.appYmlConfig.configureApp = {};
    }
    if (!context.appYmlConfig.configureApp.tab) {
      context.appYmlConfig.configureApp.tab = {};
    }
    context.appYmlConfig.configureApp.tab.domain = url.host;
    context.appYmlConfig.configureApp.tab.endpoint = url.origin;

    if (!context.appYmlConfig.deploy) {
      context.appYmlConfig.deploy = {};
    }
    if (!context.appYmlConfig.deploy.tab) {
      context.appYmlConfig.deploy.tab = {};
    }
    context.appYmlConfig.deploy.tab.port = parseInt(url.port);

    const label = task["label"];
    index = handleProvisionAndDeploy(context, index, label);
  }
}

export async function migrateSetUpBot(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "teamsfx") ||
      !(task["command"] === TaskCommand.setUpBot)
    ) {
      ++index;
      continue;
    }

    if (typeof task["label"] !== "string") {
      ++index;
      continue;
    }

    if (!context.appYmlConfig.provision) {
      context.appYmlConfig.provision = {};
    }
    context.appYmlConfig.provision.bot = {
      messagingEndpoint: `$\{{${context.placeholderMapping.botEndpoint}}}/api/messages`,
    };

    if (!context.appYmlConfig.deploy) {
      context.appYmlConfig.deploy = {};
    }
    context.appYmlConfig.deploy.bot = true;

    const envs: { [key: string]: string } = {};
    if (isCommentObject(task["args"])) {
      if (task["args"]["botId"] && typeof task["args"]["botId"] === "string") {
        envs["BOT_ID"] = task["args"]["botId"];
      }
      if (task["args"]["botPassword"] && typeof task["args"]["botPassword"] === "string") {
        const envReferencePattern = /^\$\{env:(.*)\}$/;
        const matchResult = task["args"]["botPassword"].match(envReferencePattern);
        const botPassword = matchResult ? process.env[matchResult[1]] : task["args"]["botPassword"];
        if (botPassword) {
          const cryptoProvider = new LocalCrypto(context.oldProjectSettings.projectId);
          const result = cryptoProvider.encrypt(botPassword);
          if (result.isOk()) {
            envs["SECRET_BOT_PASSWORD"] = result.value;
          }
        }
      }
      if (
        task["args"]["botMessagingEndpoint"] &&
        typeof task["args"]["botMessagingEndpoint"] === "string"
      ) {
        if (task["args"]["botMessagingEndpoint"].startsWith("http")) {
          context.appYmlConfig.provision.bot.messagingEndpoint =
            task["args"]["botMessagingEndpoint"];
        } else if (task["args"]["botMessagingEndpoint"].startsWith("/")) {
          context.appYmlConfig.provision.bot.messagingEndpoint = `$\{{${context.placeholderMapping.botEndpoint}}}${task["args"]["botMessagingEndpoint"]}`;
        }
      }
    }
    await updateLocalEnv(context.migrationContext, envs);

    const label = task["label"];
    index = handleProvisionAndDeploy(context, index, label);
  }
}

export async function migrateSetUpSSO(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "teamsfx") ||
      !(task["command"] === TaskCommand.setUpSSO)
    ) {
      ++index;
      continue;
    }

    if (typeof task["label"] !== "string") {
      ++index;
      continue;
    }

    if (!context.appYmlConfig.registerApp) {
      context.appYmlConfig.registerApp = {};
    }
    context.appYmlConfig.registerApp.aad = true;

    if (!context.appYmlConfig.configureApp) {
      context.appYmlConfig.configureApp = {};
    }
    context.appYmlConfig.configureApp.aad = true;

    if (!context.appYmlConfig.deploy) {
      context.appYmlConfig.deploy = {};
    }
    context.appYmlConfig.deploy.sso = true;

    const envs: { [key: string]: string } = {};
    if (isCommentObject(task["args"])) {
      if (task["args"]["objectId"] && typeof task["args"]["objectId"] === "string") {
        envs["AAD_APP_OBJECT_ID"] = task["args"]["objectId"];
      }
      if (task["args"]["clientId"] && typeof task["args"]["clientId"] === "string") {
        envs["AAD_APP_CLIENT_ID"] = task["args"]["clientId"];
      }
      if (task["args"]["clientSecret"] && typeof task["args"]["clientSecret"] === "string") {
        const envReferencePattern = /^\$\{env:(.*)\}$/;
        const matchResult = task["args"]["clientSecret"].match(envReferencePattern);
        const clientSecret = matchResult
          ? process.env[matchResult[1]]
          : task["args"]["clientSecret"];
        if (clientSecret) {
          const cryptoProvider = new LocalCrypto(context.oldProjectSettings.projectId);
          const result = cryptoProvider.encrypt(clientSecret);
          if (result.isOk()) {
            envs["SECRET_AAD_APP_CLIENT_SECRET"] = result.value;
          }
        }
      }
      if (
        task["args"]["accessAsUserScopeId"] &&
        typeof task["args"]["accessAsUserScopeId"] === "string"
      ) {
        envs["AAD_APP_ACCESS_AS_USER_PERMISSION_ID"] = task["args"]["accessAsUserScopeId"];
      }
    }
    await updateLocalEnv(context.migrationContext, envs);

    const label = task["label"];
    index = handleProvisionAndDeploy(context, index, label);
  }
}

export async function migratePrepareManifest(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "teamsfx") ||
      !(task["command"] === TaskCommand.prepareManifest)
    ) {
      ++index;
      continue;
    }

    if (typeof task["label"] !== "string") {
      ++index;
      continue;
    }

    let appPackagePath: string | undefined = undefined;
    if (isCommentObject(task["args"]) && typeof task["args"]["appPackagePath"] === "string") {
      appPackagePath = task["args"]["appPackagePath"];
    }

    if (!appPackagePath) {
      if (!context.appYmlConfig.registerApp) {
        context.appYmlConfig.registerApp = {};
      }
      context.appYmlConfig.registerApp.teamsApp = true;
    }

    if (!context.appYmlConfig.configureApp) {
      context.appYmlConfig.configureApp = {};
    }
    if (!context.appYmlConfig.configureApp.teamsApp) {
      context.appYmlConfig.configureApp.teamsApp = {};
    }
    context.appYmlConfig.configureApp.teamsApp.appPackagePath = appPackagePath;

    const label = task["label"];
    index = handleProvisionAndDeploy(context, index, label);
  }
}

export async function migrateValidateDependencies(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "shell") ||
      !(typeof task["command"] === "string") ||
      !task["command"].includes("${command:fx-extension.validate-dependencies}")
    ) {
      ++index;
      continue;
    }

    const newTask = generatePrerequisiteTask(task, context);

    context.tasks.splice(index, 1, newTask);
    ++index;

    const toolsArgs: InstallToolArgs = {};
    if (OldProjectSettingsHelper.includeTab(context.oldProjectSettings)) {
      toolsArgs.devCert = {
        trust: true,
      };
    }
    if (OldProjectSettingsHelper.includeFunction(context.oldProjectSettings)) {
      toolsArgs.func = true;
      toolsArgs.dotnet = true;
    }
    if (Object.keys(toolsArgs).length > 0) {
      if (!context.appYmlConfig.deploy) {
        context.appYmlConfig.deploy = {};
      }
      context.appYmlConfig.deploy.tools = toolsArgs;
    }
  }
}

export async function migrateBackendExtensionsInstall(
  context: DebugMigrationContext
): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "shell") ||
      !(
        typeof task["command"] === "string" &&
        task["command"].includes("${command:fx-extension.backend-extensions-install}")
      )
    ) {
      ++index;
      continue;
    }

    if (!context.appYmlConfig.deploy) {
      context.appYmlConfig.deploy = {};
    }
    context.appYmlConfig.deploy.dotnetCommand = {
      args: "build extensions.csproj -o ./bin --ignore-failed-sources",
      workingDirectory: `./${FolderName.Function}`,
      execPath: "${{DOTNET_PATH}}",
    };

    const label = task["label"];
    if (typeof label === "string") {
      replaceInDependsOn(label, context.tasks);
    }
    context.tasks.splice(index, 1);
  }
}

export async function migrateValidateLocalPrerequisites(
  context: DebugMigrationContext
): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "shell") ||
      !(
        typeof task["command"] === "string" &&
        task["command"].includes("${command:fx-extension.validate-local-prerequisites}")
      )
    ) {
      ++index;
      continue;
    }

    const newTask = generatePrerequisiteTask(task, context);
    context.tasks.splice(index, 1, newTask);
    ++index;

    const toolsArgs: InstallToolArgs = {};
    const npmCommands: BuildArgs[] = [];
    let dotnetCommand: BuildArgs | undefined;
    if (OldProjectSettingsHelper.includeTab(context.oldProjectSettings)) {
      toolsArgs.devCert = {
        trust: true,
      };
      npmCommands.push({
        args: `install ${defaultNpmInstallArg}`,
        workingDirectory: `./${FolderName.Frontend}`,
      });
    }

    if (OldProjectSettingsHelper.includeFunction(context.oldProjectSettings)) {
      toolsArgs.func = true;
      toolsArgs.dotnet = true;
      npmCommands.push({
        args: `install ${defaultNpmInstallArg}`,
        workingDirectory: `./${FolderName.Function}`,
      });
      dotnetCommand = {
        args: "build extensions.csproj -o ./bin --ignore-failed-sources",
        workingDirectory: `./${FolderName.Function}`,
        execPath: "${{DOTNET_PATH}}",
      };
    }

    if (OldProjectSettingsHelper.includeBot(context.oldProjectSettings)) {
      if (OldProjectSettingsHelper.includeFuncHostedBot(context.oldProjectSettings)) {
        toolsArgs.func = true;
      }
      npmCommands.push({
        args: `install ${defaultNpmInstallArg}`,
        workingDirectory: `./${FolderName.Bot}`,
      });
    }

    if (Object.keys(toolsArgs).length > 0 || npmCommands.length > 0 || dotnetCommand) {
      if (!context.appYmlConfig.deploy) {
        context.appYmlConfig.deploy = {};
      }
      if (Object.keys(toolsArgs).length > 0) {
        context.appYmlConfig.deploy.tools = toolsArgs;
      }
      if (npmCommands.length > 0) {
        context.appYmlConfig.deploy.npmCommands = npmCommands;
      }
      context.appYmlConfig.deploy.dotnetCommand = dotnetCommand;
    }
  }
}

export async function migrateNgrokStartTask(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      isCommentObject(task) &&
      ((typeof task["dependsOn"] === "string" && task["dependsOn"] === "teamsfx: ngrok start") ||
        (isCommentArray(task["dependsOn"]) && task["dependsOn"].includes("teamsfx: ngrok start")))
    ) {
      const newTask = generateLocalTunnelTask(context);
      context.tasks.splice(index + 1, 0, newTask);
      break;
    } else {
      ++index;
    }
  }
  replaceInDependsOn("teamsfx: ngrok start", context.tasks, TaskLabel.StartLocalTunnel);
}

export async function migrateNgrokStartCommand(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "teamsfx") ||
      !(task["command"] === "ngrok start")
    ) {
      ++index;
      continue;
    }

    const newTask = generateLocalTunnelTask(context, task);
    context.tasks.splice(index, 1, newTask);
    ++index;
  }
}

function generatePrerequisiteTask(
  task: CommentObject,
  context: DebugMigrationContext
): CommentObject {
  const comment = `{
    // Check if all required prerequisites are installed and will install them if not.
    // See https://aka.ms/teamsfx-check-prerequisites-task to know the details and how to customize the args.
  }`;
  const newTask: CommentObject = assign(parse(comment), task) as CommentObject;

  newTask["type"] = "teamsfx";
  newTask["command"] = "debug-check-prerequisites";

  const prerequisites = [
    `"${Prerequisite.nodejs}", // Validate if Node.js is installed.`,
    `"${Prerequisite.m365Account}", // Sign-in prompt for Microsoft 365 account, then validate if the account enables the sideloading permission.`,
    `"${Prerequisite.portOccupancy}", // Validate available ports to ensure those debug ones are not occupied.`,
  ];
  const prerequisitesComment = `
  [
    ${prerequisites.join("\n  ")}
  ]`;

  const ports: string[] = [];
  if (OldProjectSettingsHelper.includeTab(context.oldProjectSettings)) {
    ports.push(`${TaskDefaultValue.checkPrerequisites.ports.tabService}, // tab service port`);
  }
  if (OldProjectSettingsHelper.includeBot(context.oldProjectSettings)) {
    ports.push(`${TaskDefaultValue.checkPrerequisites.ports.botService}, // bot service port`);
    ports.push(
      `${TaskDefaultValue.checkPrerequisites.ports.botDebug}, // bot inspector port for Node.js debugger`
    );
  }
  if (OldProjectSettingsHelper.includeFunction(context.oldProjectSettings)) {
    ports.push(
      `${TaskDefaultValue.checkPrerequisites.ports.backendService}, // backend service port`
    );
    ports.push(
      `${TaskDefaultValue.checkPrerequisites.ports.backendDebug}, // backend inspector port for Node.js debugger`
    );
  }
  const portsComment = `
  [
    ${ports.join("\n  ")}
  ]
  `;

  const args: { [key: string]: CommentJSONValue } = {
    prerequisites: parse(prerequisitesComment),
    portOccupancy: parse(portsComment),
  };

  newTask["args"] = args as CommentJSONValue;
  return newTask;
}

function generateLocalTunnelTask(context: DebugMigrationContext, task?: CommentObject) {
  const comment = `{
      // Start the local tunnel service to forward public ngrok URL to local port and inspect traffic.
      // See https://aka.ms/teamsfx-local-tunnel-task for the detailed args definitions,
      // as well as samples to:
      //   - use your own ngrok command / configuration / binary
      //   - use your own tunnel solution
      //   - provide alternatives if ngrok does not work on your dev machine
    }`;
  const placeholderComment = `
    {
      // Keep consistency with migrated configuration.
    }
  `;
  const newTask = assign(task ?? parse(`{"label": "${TaskLabel.StartLocalTunnel}"}`), {
    type: "teamsfx",
    command: TaskCommand.startLocalTunnel,
    args: {
      ngrokArgs: TaskDefaultValue.startLocalTunnel.ngrokArgs,
      env: "local",
      output: assign(parse(placeholderComment), {
        endpoint: context.placeholderMapping.botEndpoint,
        domain: context.placeholderMapping.botDomain,
      }),
    },
    isBackground: true,
    problemMatcher: "$teamsfx-local-tunnel-watch",
  });
  return assign(parse(comment), newTask);
}

function handleProvisionAndDeploy(
  context: DebugMigrationContext,
  index: number,
  label: string
): number {
  context.tasks.splice(index, 1);

  const existingLabels = getLabels(context.tasks);

  const generatedBefore = context.generatedLabels.find((value) =>
    value.startsWith("Create resources")
  );
  const createResourcesLabel = generatedBefore || generateLabel("Create resources", existingLabels);

  const setUpLocalProjectsLabel =
    context.generatedLabels.find((value) => value.startsWith("Set up local projects")) ||
    generateLabel("Set up local projects", existingLabels);

  if (!generatedBefore) {
    context.generatedLabels.push(createResourcesLabel);
    const createResources = createResourcesTask(createResourcesLabel);
    context.tasks.splice(index, 0, createResources);
    ++index;

    context.generatedLabels.push(setUpLocalProjectsLabel);
    const setUpLocalProjects = setUpLocalProjectsTask(setUpLocalProjectsLabel);
    context.tasks.splice(index, 0, setUpLocalProjects);
    ++index;
  }

  replaceInDependsOn(label, context.tasks, createResourcesLabel, setUpLocalProjectsLabel);

  return index;
}

function replaceInDependsOn(
  label: string,
  tasks: CommentArray<CommentJSONValue>,
  ...replacements: string[]
): void {
  for (const task of tasks) {
    if (isCommentObject(task) && task["dependsOn"]) {
      if (typeof task["dependsOn"] === "string") {
        if (task["dependsOn"] === label) {
          if (replacements.length > 0) {
            task["dependsOn"] = new CommentArray(...replacements);
          } else {
            delete task["dependsOn"];
          }
        }
      } else if (Array.isArray(task["dependsOn"])) {
        const index = task["dependsOn"].findIndex((value) => value === label);
        if (index !== -1) {
          if (replacements.length > 0 && !task["dependsOn"].includes(replacements[0])) {
            task["dependsOn"].splice(index, 1, ...replacements);
          } else {
            task["dependsOn"].splice(index, 1);
          }
        }
      }
    }
  }
}

function getLabels(tasks: CommentArray<CommentJSONValue>): string[] {
  const labels: string[] = [];
  for (const task of tasks) {
    if (isCommentObject(task) && typeof task["label"] === "string") {
      labels.push(task["label"]);
    }
  }

  return labels;
}
