// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { TaskOverallLabel } from "../../../common/local/constants";
import { CommentJSONValue, CommentObject, CommentArray } from "comment-json";
import * as commentJson from "comment-json";
import {
  Prerequisite,
  TaskCommand,
  TaskDefaultValue,
  TaskLabel,
} from "../../../common/local/constants";
import { ProgrammingLanguage } from "../../constants";

export function generateTasksJson(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean,
  includeFuncHostedBot: boolean,
  includeSSO: boolean,
  programmingLanguage: string
): CommentJSONValue {
  const comment = `
  // This file is automatically generated by Teams Toolkit.
  // The teamsfx tasks defined in this file require Teams Toolkit version >= 4.1.0.
  // See https://aka.ms/teamsfx-debug-tasks for details on how to customize each task and how to integrate with existing Teams Toolkit projects.
  {}
  `;
  return commentJson.assign(commentJson.parse(comment), {
    version: "2.0.0",
    tasks: generateTasks(
      includeFrontend,
      includeBackend,
      includeBot,
      includeFuncHostedBot,
      includeSSO,
      programmingLanguage
    ),
  });
}

export function generateM365TasksJson(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean,
  includeFuncHostedBot: boolean,
  includeSSO: boolean,
  programmingLanguage: string
): CommentJSONValue {
  const comment = `
  // This file is automatically generated by Teams Toolkit.
  // See https://aka.ms/teamsfx-debug-tasks to know the details and how to customize each task.
  {}
  `;
  return commentJson.assign(commentJson.parse(comment), {
    version: "2.0.0",
    tasks: generateM365Tasks(
      includeFrontend,
      includeBackend,
      includeBot,
      includeFuncHostedBot,
      includeSSO,
      programmingLanguage
    ),
  });
}

export function generateTasks(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean,
  includeFuncHostedBot: boolean,
  includeSSO: boolean,
  programmingLanguage: string
): (Record<string, unknown> | CommentJSONValue)[] {
  /**
   * Referenced by launch.json
   *   - Start Teams App Locally
   *
   * Referenced inside tasks.json
   *   - Validate & install prerequisites
   *   - Install npm packages
   *   - Start local tunnel
   *   - Set up tab
   *   - Set up bot
   *   - Set up SSO
   *   - Build & upload Teams manifest
   *   - Start services
   *   - Start frontend
   *   - Start backend
   *   - Install Azure Functions binding extensions
   *   - Watch backend
   *   - Start bot
   *   - Start Azurite emulator
   */
  const tasks: (Record<string, unknown> | CommentJSONValue)[] = [
    startTeamsAppLocally(includeFrontend, includeBackend, includeBot, includeSSO),
    validateAndInstallPrerequisites(
      includeFrontend,
      includeBackend,
      includeBot,
      includeFuncHostedBot,
      false
    ),
    installNPMpackages(includeFrontend, includeBackend, includeBot, false),
  ];

  if (includeBot) {
    tasks.push(startLocalTunnel());
  }

  if (includeFrontend) {
    tasks.push(setUpTab());
  }

  if (includeBot) {
    tasks.push(setUpBot());
  }

  if (includeSSO) {
    tasks.push(setUpSSO());
  }

  tasks.push(buildAndUploadTeamsManifest());

  tasks.push(startServices(includeFrontend, includeBackend, includeBot));

  if (includeFrontend) {
    tasks.push(startFrontend());
  }

  if (includeBackend) {
    tasks.push(startBackend(programmingLanguage));
    tasks.push(installAzureFunctionsBindingExtensions());
    if (programmingLanguage === ProgrammingLanguage.TS) {
      tasks.push(watchBackend());
    }
  }

  if (includeBot) {
    if (includeFuncHostedBot) {
      tasks.push(startFuncHostedBot(includeFrontend, programmingLanguage));
      tasks.push(startAzuriteEmulator());
      if (programmingLanguage === ProgrammingLanguage.TS) {
        tasks.push(watchFuncHostedBot());
      }
    } else {
      tasks.push(startBot(includeFrontend));
    }
  }

  return tasks;
}

export function generateM365Tasks(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean,
  includeFuncHostedBot: boolean,
  includeSSO: boolean,
  programmingLanguage: string
): (Record<string, unknown> | CommentJSONValue)[] {
  /**
   * Referenced by launch.json
   *   - Start Teams App Locally
   *   - Start Teams App Locally & Install App
   *
   * Referenced inside tasks.json
   *   - Validate & install prerequisites
   *   - Install npm packages
   *   - Start local tunnel
   *   - Set up tab
   *   - Set up bot
   *   - Set up SSO
   *   - Build & upload Teams manifest
   *   - Start services
   *   - Start frontend
   *   - Start backend
   *   - Install Azure Functions binding extensions
   *   - Watch backend
   *   - Start bot
   *   - Start Azurite emulator
   *   - Install app in Teams
   */
  const tasks = generateTasks(
    includeFrontend,
    includeBackend,
    includeBot,
    includeFuncHostedBot,
    includeSSO,
    programmingLanguage
  );
  tasks.splice(
    1,
    0,
    startTeamsAppLocallyAndInstallApp(includeFrontend, includeBackend, includeBot, includeSSO)
  );
  tasks.push(installAppInTeams());
  return tasks;
}

export function mergeTasksJson(existingData: CommentObject, newData: CommentObject): CommentObject {
  const mergedData = commentJson.assign(commentJson.parse(`{}`), existingData) as CommentObject;

  if (mergedData.version === undefined) {
    mergedData.version = newData.version;
  }

  if (mergedData.tasks === undefined) {
    mergedData.tasks = newData.tasks;
  } else {
    const existingTasks = mergedData.tasks as CommentArray<CommentObject>;
    const newTasks = newData.tasks as CommentArray<CommentObject>;
    const keptTasks = new CommentArray<CommentObject>();
    for (const existingTask of existingTasks) {
      if (
        !newTasks.some(
          (newTask) => existingTask.label === newTask.label && existingTask.type === newTask.type
        )
      ) {
        keptTasks.push(existingTask);
      }
    }
    mergedData.tasks = new CommentArray<CommentObject>(...keptTasks, ...newTasks);
  }

  if (mergedData.inputs === undefined) {
    mergedData.inputs = newData.inputs;
  } else if (newData.inputs !== undefined) {
    const existingInputs = mergedData.inputs as CommentArray<CommentObject>;
    const newInputs = newData.inputs as CommentArray<CommentObject>;
    const keptInputs = new CommentArray<CommentObject>();
    for (const existingInput of existingInputs) {
      if (
        !newInputs.some(
          (newInput) => existingInput.id === newInput.id && existingInput.type === newInput.type
        )
      ) {
        keptInputs.push(existingInput);
      }
    }
    mergedData.inputs = new CommentArray<CommentObject>(...keptInputs, ...newInputs);
  }

  return mergedData;
}

function startTeamsAppLocally(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean,
  includeSSO: boolean
): Record<string, unknown> {
  const result = {
    label: TaskOverallLabel.TransparentDefault,
    dependsOn: [TaskLabel.PrerequisiteCheck, TaskLabel.InstallNpmPackages],
    dependsOrder: "sequence",
  };
  if (includeBot) {
    result.dependsOn.push(TaskLabel.StartLocalTunnel);
  }
  if (includeFrontend) {
    result.dependsOn.push(TaskLabel.SetUpTab);
  }
  if (includeBot) {
    result.dependsOn.push(TaskLabel.SetUpBot);
  }
  if (includeSSO) {
    result.dependsOn.push(TaskLabel.SetUpSSO);
  }
  result.dependsOn.push(TaskLabel.PrepareManifest, TaskLabel.StartServices);

  return result;
}

function startTeamsAppLocallyAndInstallApp(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean,
  includeSSO: boolean
): Record<string, unknown> {
  const result = startTeamsAppLocally(includeFrontend, includeBackend, includeBot, includeSSO);
  result.label = TaskOverallLabel.TransparentM365;
  (result.dependsOn as string[]).push(TaskLabel.InstallAppInTeams);

  return result;
}

function validateAndInstallPrerequisites(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean,
  includeFuncHostedBot: boolean,
  includeSpfx: boolean
): CommentJSONValue {
  const prerequisites = [
    `"${Prerequisite.nodejs}", // Validate if Node.js is installed.`,
    `"${Prerequisite.m365Account}", // Sign-in prompt for Microsoft 365 account, then validate if the account enables the sideloading permission.`,
  ];
  const ports: string[] = [];
  if (includeFrontend) {
    prerequisites.push(
      `"${Prerequisite.devCert}", // Install localhost SSL certificate. It's used to serve the development sites over HTTPS to debug the Tab app in Teams.`
    );
    ports.push(`${TaskDefaultValue.checkPrerequisites.ports.tabService}, // tab service port`);
  }
  if (includeBackend) {
    prerequisites.push(
      `"${Prerequisite.func}", // Install Azure Functions Core Tools. It's used to serve Azure Functions hosted project locally.`,
      `"${Prerequisite.dotnet}", // Ensure .NET Core SDK is installed. TeamsFx Azure Functions project depends on extra .NET binding extensions for HTTP trigger authorization.`
    );
    ports.push(
      `${TaskDefaultValue.checkPrerequisites.ports.backendService}, // backend service port`,
      `${TaskDefaultValue.checkPrerequisites.ports.backendDebug}, // backend inspector port for Node.js debugger`
    );
  }
  if (includeFuncHostedBot && !includeBackend) {
    prerequisites.push(
      `"${Prerequisite.func}", // Install Azure Functions Core Tools. It's used to serve Azure Functions hosted project locally.`
    );
  }
  if (includeBot) {
    prerequisites.push(
      `"${Prerequisite.ngrok}", // Install Ngrok. Bot project requires a public message endpoint, and ngrok can help create public tunnel for your local service.`
    );
    ports.push(
      `${TaskDefaultValue.checkPrerequisites.ports.botService}, // bot service port`,
      `${TaskDefaultValue.checkPrerequisites.ports.botDebug}, // bot inspector port for Node.js debugger`
    );
  }
  if (includeSpfx) {
    ports.push(`${TaskDefaultValue.checkPrerequisites.ports.spfxService}, // SPFx service port`);
  }
  prerequisites.push(
    `"${Prerequisite.portOccupancy}", // Validate available ports to ensure those debug ones are not occupied.`
  );

  const prerequisitesComment = `
  [
    ${prerequisites.join("\n  ")}
  ]`;
  const portsComment = `
  [
    ${ports.join("\n  ")}
  ]
  `;

  const comment = `{
    // Check if all required prerequisites are installed and will install them if not.
    // See https://aka.ms/teamsfx-check-prerequisites-task to know the details and how to customize the args.
  }`;

  const args: { [key: string]: CommentJSONValue } = {
    prerequisites: commentJson.parse(prerequisitesComment),
  };

  args["portOccupancy"] = commentJson.parse(portsComment);

  const task = {
    label: TaskLabel.PrerequisiteCheck,
    type: "teamsfx",
    command: TaskCommand.checkPrerequisites,
    args: args,
  };

  return commentJson.assign(commentJson.parse(comment), task);
}

function installNPMpackages(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean,
  includeSpfx: boolean
): CommentJSONValue {
  const comment = `{
    // Check if all the npm packages are installed and will install them if not.
    // See https://aka.ms/teamsfx-npm-package-task to know the details and how to customize the args.
  }`;
  const result = {
    label: TaskLabel.InstallNpmPackages,
    type: "teamsfx",
    command: TaskCommand.npmInstall,
    args: {
      projects: [] as Record<string, unknown>[],
    },
  };
  if (includeFrontend) {
    result.args.projects.push({
      cwd: "${workspaceFolder}/tabs",
      npmInstallArgs: TaskDefaultValue.npmInstall.npmInstallArgs,
    });
  }
  if (includeBackend) {
    result.args.projects.push({
      cwd: "${workspaceFolder}/api",
      npmInstallArgs: TaskDefaultValue.npmInstall.npmInstallArgs,
    });
  }
  if (includeBot) {
    result.args.projects.push({
      cwd: "${workspaceFolder}/bot",
      npmInstallArgs: TaskDefaultValue.npmInstall.npmInstallArgs,
    });
  }
  if (includeSpfx) {
    result.args.projects.push({
      cwd: "${workspaceFolder}/SPFx",
      npmInstallArgs: TaskDefaultValue.npmInstall.npmInstallArgs,
    });
  }

  return commentJson.assign(commentJson.parse(comment), result);
}

function installAzureFunctionsBindingExtensions(): CommentJSONValue {
  const comment = `{
    // TeamsFx Azure Functions project depends on extra Azure Functions binding extensions for HTTP trigger authorization.
  }`;
  const task = {
    label: TaskLabel.InstallAzureFuncBindingExt,
    type: "shell",
    command: "dotnet build extensions.csproj -o ./bin --ignore-failed-sources",
    options: {
      cwd: "${workspaceFolder}/api",
      env: {
        PATH: "${command:fx-extension.get-dotnet-path}${env:PATH}",
      },
    },
    presentation: {
      reveal: "silent",
    },
  };
  return commentJson.assign(commentJson.parse(comment), task);
}

function startLocalTunnel(): CommentJSONValue {
  const comment = `{
    // Start the local tunnel service to forward public ngrok URL to local port and inspect traffic.
    // See https://aka.ms/teamsfx-local-tunnel-task for the detailed args definitions,
    // as well as samples to:
    //   - use your own ngrok command / configuration / binary
    //   - use your own tunnel solution
    //   - provide alternatives if ngrok does not work on your dev machine
  }`;
  const task = {
    label: TaskLabel.StartLocalTunnel,
    type: "teamsfx",
    command: TaskCommand.startLocalTunnel,
    args: {
      ngrokArgs: TaskDefaultValue.startLocalTunnel.ngrokArgs,
    },
    isBackground: true,
    problemMatcher: "$teamsfx-local-tunnel-watch",
  };
  return commentJson.assign(commentJson.parse(comment), task);
}

function setUpTab(): CommentJSONValue {
  const comment = `{
    // Prepare local launch information for Tab.
    // See https://aka.ms/teamsfx-debug-set-up-tab-task to know the details and how to customize the args.
  }`;
  const task = {
    label: TaskLabel.SetUpTab,
    type: "teamsfx",
    command: TaskCommand.setUpTab,
    args: {
      baseUrl: TaskDefaultValue.setUpTab.baseUrl,
    },
  };
  return commentJson.assign(commentJson.parse(comment), task);
}

function setUpBot(): CommentJSONValue {
  const comment = `{
    // Register resources and prepare local launch information for Bot.
    // See https://aka.ms/teamsfx-debug-set-up-bot-task to know the details and how to customize the args.
  }`;
  const existingBot = `
  {
    //// Enter your own bot information if using the existing bot. ////
    // "botId": "",
    // "botPassword": "", // use plain text or environment variable reference like \${env:BOT_PASSWORD}
    "botMessagingEndpoint": "/api/messages", // use your own routing "/any/path", or full URL "https://contoso.com/any/path"
  }
  `;
  const task = {
    label: TaskLabel.SetUpBot,
    type: "teamsfx",
    command: TaskCommand.setUpBot,
    args: commentJson.parse(existingBot),
  };
  return commentJson.assign(commentJson.parse(comment), task);
}

function setUpSSO(): CommentJSONValue {
  const comment = `{
    // Register resources and prepare local launch information for SSO functionality.
    // See https://aka.ms/teamsfx-debug-set-up-sso-task to know the details and how to customize the args.
  }`;
  const existingAAD = `
  {
    //// Enter your own AAD app information if using the existing AAD app. ////
    // "objectId": "",
    // "clientId": "",
    // "clientSecret": "", // use plain text or environment variable reference like \${env:CLIENT_SECRET}
    // "accessAsUserScopeId": "
  }
  `;
  const task = {
    label: TaskLabel.SetUpSSO,
    type: "teamsfx",
    command: TaskCommand.setUpSSO,
    args: commentJson.parse(existingAAD),
  };
  return commentJson.assign(commentJson.parse(comment), task);
}

function buildAndUploadTeamsManifest(): CommentJSONValue {
  const comment = `
  {
    // Build and upload Teams manifest.
    // See https://aka.ms/teamsfx-debug-prepare-manifest-task to know the details and how to customize the args.
  }`;
  const existingApp = `
  {
    //// Enter your own Teams app package path if using the existing Teams manifest. ////
    // "appPackagePath": ""
  }
  `;
  const task = {
    label: TaskLabel.PrepareManifest,
    type: "teamsfx",
    command: TaskCommand.prepareManifest,
    args: commentJson.parse(existingApp),
  };
  return commentJson.assign(commentJson.parse(comment), task);
}

function startFrontend(): Record<string, unknown> {
  return {
    label: TaskLabel.StartFrontend,
    type: "shell",
    command: "npm run dev:teamsfx",
    isBackground: true,
    options: {
      cwd: "${workspaceFolder}/tabs",
    },
    problemMatcher: {
      pattern: {
        regexp: "^.*$",
        file: 0,
        location: 1,
        message: 2,
      },
      background: {
        activeOnStart: true,
        beginsPattern: ".*",
        endsPattern: "Compiled|Failed|compiled|failed",
      },
    },
  };
}

function startBackend(programmingLanguage: string): Record<string, unknown> {
  const result = {
    label: TaskLabel.StartBackend,
    type: "shell",
    command: "npm run dev:teamsfx",
    isBackground: true,
    options: {
      cwd: "${workspaceFolder}/api",
      env: {
        PATH: "${command:fx-extension.get-func-path}${env:PATH}",
      },
    },
    problemMatcher: {
      pattern: {
        regexp: "^.*$",
        file: 0,
        location: 1,
        message: 2,
      },
      background: {
        activeOnStart: true,
        beginsPattern: "^.*(Job host stopped|signaling restart).*$",
        endsPattern:
          "^.*(Worker process started and initialized|Host lock lease acquired by instance ID).*$",
      },
    },
    presentation: {
      reveal: "silent",
    },
    dependsOn: [TaskLabel.InstallAzureFuncBindingExt],
  } as Record<string, unknown>;

  if (programmingLanguage === ProgrammingLanguage.TS) {
    (result.dependsOn as string[]).push(TaskLabel.WatchBackend);
  }

  return result;
}

function watchBackend(): Record<string, unknown> {
  return {
    label: TaskLabel.WatchBackend,
    type: "shell",
    command: "npm run watch:teamsfx",
    isBackground: true,
    options: {
      cwd: "${workspaceFolder}/api",
    },
    problemMatcher: "$tsc-watch",
    presentation: {
      reveal: "silent",
    },
  };
}

function watchFuncHostedBot(): Record<string, unknown> {
  return {
    label: TaskLabel.WatchBot,
    type: "shell",
    command: "npm run watch:teamsfx",
    isBackground: true,
    options: {
      cwd: "${workspaceFolder}/bot",
    },
    problemMatcher: "$tsc-watch",
    presentation: {
      reveal: "silent",
    },
  };
}

function startBot(includeFrontend: boolean): Record<string, unknown> {
  const result: Record<string, unknown> = {
    label: TaskLabel.StartBot,
    type: "shell",
    command: "npm run dev:teamsfx",
    isBackground: true,
    options: {
      cwd: "${workspaceFolder}/bot",
    },
    problemMatcher: {
      pattern: [
        {
          regexp: "^.*$",
          file: 0,
          location: 1,
          message: 2,
        },
      ],
      background: {
        activeOnStart: true,
        beginsPattern: "[nodemon] starting",
        endsPattern: "restify listening to|Bot/ME service listening at|[nodemon] app crashed",
      },
    },
  };

  if (includeFrontend) {
    result.presentation = { reveal: "silent" };
  }

  return result;
}

function startFuncHostedBot(
  includeFrontend: boolean,
  programmingLanguage: string
): Record<string, unknown> {
  const result: Record<string, unknown> = {
    label: TaskLabel.StartBot,
    type: "shell",
    command: "npm run dev:teamsfx",
    isBackground: true,
    options: {
      cwd: "${workspaceFolder}/bot",
      env: {
        PATH: "${command:fx-extension.get-func-path}${env:PATH}",
      },
    },
    problemMatcher: {
      pattern: {
        regexp: "^.*$",
        file: 0,
        location: 1,
        message: 2,
      },
      background: {
        activeOnStart: true,
        beginsPattern: "^.*(Job host stopped|signaling restart).*$",
        endsPattern:
          "^.*(Worker process started and initialized|Host lock lease acquired by instance ID).*$",
      },
    },
  };

  if (includeFrontend) {
    result.presentation = { reveal: "silent" };
  }

  const dependsOn: string[] = [TaskLabel.StartAzuriteEmulator];
  if (programmingLanguage === ProgrammingLanguage.TS) {
    dependsOn.push(TaskLabel.WatchBot);
  }
  result.dependsOn = dependsOn;

  return result;
}

function startServices(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean
): Record<string, unknown> {
  const dependsOn: string[] = [];
  if (includeFrontend) {
    dependsOn.push(TaskLabel.StartFrontend);
  }
  if (includeBackend) {
    dependsOn.push(TaskLabel.StartBackend);
  }
  if (includeBot) {
    dependsOn.push(TaskLabel.StartBot);
  }
  return {
    label: TaskLabel.StartServices,
    dependsOn,
  };
}

function startAzuriteEmulator(): Record<string, unknown> {
  return {
    label: TaskLabel.StartAzuriteEmulator,
    type: "shell",
    command: "npm run prepare-storage:teamsfx",
    isBackground: true,
    problemMatcher: {
      pattern: [
        {
          regexp: "^.*$",
          file: 0,
          location: 1,
          message: 2,
        },
      ],
      background: {
        activeOnStart: true,
        beginsPattern: "Azurite",
        endsPattern: "successfully listening",
      },
    },
    options: {
      cwd: "${workspaceFolder}/bot",
    },
    presentation: { reveal: "silent" },
  };
}

function installAppInTeams(): Record<string, unknown> {
  return {
    label: TaskLabel.InstallAppInTeams,
    type: "shell",
    command: "exit ${command:fx-extension.install-app-in-teams}",
    presentation: {
      reveal: "never",
    },
  };
}

export function generateSpfxTasksJson(): CommentJSONValue {
  const comment = `
  // This file is automatically generated by Teams Toolkit.
  // The teamsfx tasks defined in this file require Teams Toolkit version >= 4.1.0.
  // See https://aka.ms/teamsfx-debug-tasks for details on how to customize each task and how to integrate with existing Teams Toolkit projects.
  {}
  `;
  return commentJson.assign(commentJson.parse(comment), {
    version: "2.0.0",
    tasks: generateSpfxTasks(),
    inputs: generateInputs(),
  });
}

function generateSpfxTasks(): (Record<string, unknown> | CommentJSONValue)[] {
  /**
   * Referenced by launch.json
   *   - Start Teams App Locally
   *
   * Referenced inside tasks.json
   *   - Validate & install prerequisites
   *   - Build & upload Teams manifest
   *   - gulp serve
   */
  return [
    startSpfxTeamsAppLocally(),
    startSpfxTeamsAppLocallyAndInstallAppInTeams(),
    validateAndInstallPrerequisites(false, false, false, false, true),
    buildAndUploadTeamsManifest(),
    installNPMpackages(false, false, false, true),
    gulpTrustDevCert(),
    gulpServe(),
    installAppInTeams(),
    terminateAllTasks(),
  ];
}

function startSpfxTeamsAppLocally(): Record<string, unknown> {
  return {
    label: TaskOverallLabel.TransparentDefault,
    dependsOn: [TaskLabel.PrerequisiteCheck, TaskLabel.PrepareManifest, TaskLabel.GulpServe],
    dependsOrder: "sequence",
  };
}

function startSpfxTeamsAppLocallyAndInstallAppInTeams(): Record<string, unknown> {
  return {
    label: TaskOverallLabel.TransparentM365,
    dependsOn: [
      TaskLabel.PrerequisiteCheck,
      TaskLabel.PrepareManifest,
      TaskLabel.GulpServe,
      TaskLabel.InstallAppInTeams,
    ],
    dependsOrder: "sequence",
  };
}

function gulpTrustDevCert(): Record<string, unknown> {
  return {
    label: TaskLabel.GulpTrustDevCert,
    type: "process",
    command: "node",
    args: ["${workspaceFolder}/SPFx/node_modules/gulp/bin/gulp.js", "trust-dev-cert"],
    options: {
      cwd: "${workspaceFolder}/SPFx",
    },
    dependsOn: TaskLabel.InstallNpmPackages,
  };
}

function gulpServe(): Record<string, unknown> {
  return {
    label: "gulp serve",
    type: "process",
    command: "node",
    args: ["${workspaceFolder}/SPFx/node_modules/gulp/bin/gulp.js", "serve", "--nobrowser"],
    problemMatcher: [
      {
        pattern: [
          {
            regexp: ".",
            file: 1,
            location: 2,
            message: 3,
          },
        ],
        background: {
          activeOnStart: true,
          beginsPattern: "^.*Starting gulp.*",
          endsPattern: "^.*Finished subtask 'reload'.*",
        },
      },
    ],
    isBackground: true,
    options: {
      cwd: "${workspaceFolder}/SPFx",
    },
    dependsOn: TaskLabel.GulpTrustDevCert,
  };
}

function terminateAllTasks(): Record<string, unknown> {
  return {
    label: "Terminate All Tasks",
    command: "echo ${input:terminate}",
    type: "shell",
    problemMatcher: [],
  };
}

function generateInputs(): Record<string, unknown>[] {
  return [
    {
      id: "terminate",
      type: "command",
      command: "workbench.action.tasks.terminate",
      args: "terminateAll",
    },
  ];
}
