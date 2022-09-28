// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { ProgrammingLanguage } from "../../../../../common/local/constants";
import { CommentJSONValue, CommentObject, CommentArray } from "comment-json";
import * as commentJson from "comment-json";

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
  // See https://aka.ms/teamsfx-debug-tasks to know the details and how to customize each task.
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
   */
  const tasks: (Record<string, unknown> | CommentJSONValue)[] = [
    startTeamsAppLocally(includeFrontend, includeBackend, includeBot, includeSSO),
    validateAndInstallPrerequisites(
      includeFrontend,
      includeBackend,
      includeBot,
      includeFuncHostedBot
    ),
    installNPMpackages(includeFrontend, includeBackend, includeBot),
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
    if (programmingLanguage === ProgrammingLanguage.typescript) {
      tasks.push(watchBackend());
    }
  }

  if (includeBot) {
    if (includeFuncHostedBot) {
      tasks.push(startFuncHostedBot(includeFrontend, programmingLanguage));
      tasks.push(startAzuriteEmulator());
      if (programmingLanguage === ProgrammingLanguage.typescript) {
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
   *   - install app in Teams
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
    label: "Start Teams App Locally",
    dependsOn: ["Validate & install prerequisites", "Install npm packages"],
    dependsOrder: "sequence",
  };
  if (includeBot) {
    result.dependsOn.push("Start local tunnel");
  }
  if (includeFrontend) {
    result.dependsOn.push("Set up tab");
  }
  if (includeBot) {
    result.dependsOn.push("Set up bot");
  }
  if (includeSSO) {
    result.dependsOn.push("Set up SSO");
  }
  result.dependsOn.push("Build & upload Teams manifest", "Start services");

  return result;
}

function startTeamsAppLocallyAndInstallApp(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean,
  includeSSO: boolean
): Record<string, unknown> {
  const result = startTeamsAppLocally(includeFrontend, includeBackend, includeBot, includeSSO);
  result.label = "Start Teams App Locally & Install App";
  (result.dependsOn as string[]).push("Install app in Teams");

  return result;
}

function validateAndInstallPrerequisites(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean,
  includeFuncHostedBot: boolean
): CommentJSONValue {
  const prerequisites = [
    `"nodejs", // Validate if Node.js is installed.`,
    `"m365Account", // Sign-in prompt for Microsoft 365 account, then validate if the account enables the sideloading permission.`,
  ];
  const ports: string[] = [];
  if (includeFrontend) {
    prerequisites.push(
      `"devCert", // Install localhost SSL certificate. It's used to serve the development sites over HTTPS to debug the Tab app in Teams.`
    );
    ports.push("53000, // tab service port");
  }
  if (includeBackend) {
    prerequisites.push(
      `"func", // Install Azure Functions Core Tools. It's used to serve Azure Functions hosted project locally.`,
      `"dotnet", // Ensure .NET Core SDK is installed. TeamsFx Azure Functions project depends on extra .NET binding extensions for HTTP trigger authorization.`
    );
    ports.push("7071, // backend service port", "9229, // backend debug port");
  }
  if (includeFuncHostedBot && !includeBackend) {
    prerequisites.push(
      `"func", // Install Azure Functions Core Tools. It's used to serve Azure Functions hosted project locally.`
    );
  }
  if (includeBot) {
    prerequisites.push(
      `"ngrok", // Install Ngrok. Bot project requires a public message endpoint, and ngrok can help create public tunnel for your local service.`
    );
    ports.push("3978, // bot service port", "9239, // bot debug port");
  }
  prerequisites.push(
    `"portOccupancy", // Validate available ports to ensure those local debug ones are not occupied.`
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
    // See https://aka.ms/teamsfx-debug-tasks#debug-check-prerequisites to know the details and how to customize the args.
  }`;

  const task = {
    label: "Validate & install prerequisites",
    type: "teamsfx",
    command: "debug-check-prerequisites",
    args: {
      prerequisites: commentJson.parse(prerequisitesComment),
      portOccupancy: commentJson.parse(portsComment),
    },
  };

  return commentJson.assign(commentJson.parse(comment), task);
}

function installNPMpackages(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean
): CommentJSONValue {
  const comment = `{
    // Check if all the npm packages are installed and will install them if not.
    // See https://aka.ms/teamsfx-debug-tasks#debug-npm-install to know the details and how to customize the args.
  }`;
  const result = {
    label: "Install npm packages",
    type: "teamsfx",
    command: "debug-npm-install",
    args: {
      projects: [] as Record<string, unknown>[],
    },
  };
  if (includeFrontend) {
    result.args.projects.push({
      cwd: "${workspaceFolder}/tabs",
      npmInstallArgs: ["--no-audit"],
    });
  }
  if (includeBackend) {
    result.args.projects.push({
      cwd: "${workspaceFolder}/api",
      npmInstallArgs: ["--no-audit"],
    });
  }
  if (includeBot) {
    result.args.projects.push({
      cwd: "${workspaceFolder}/bot",
      npmInstallArgs: ["--no-audit"],
    });
  }

  return commentJson.assign(commentJson.parse(comment), result);
}

function installAzureFunctionsBindingExtensions(): CommentJSONValue {
  const comment = `{
    // TeamsFx Azure Functions project depends on extra Azure Functions binding extensions for HTTP trigger authorization.
  }`;
  const task = {
    label: "Install Azure Functions binding extensions",
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
    // See https://aka.ms/teamsfx-debug-tasks#debug-start-local-tunnel to know the details and how to customize the args.
  }`;
  const task = {
    label: "Start local tunnel",
    type: "teamsfx",
    command: "debug-start-local-tunnel",
    args: {
      ngrokArgs: "http 3978 --log=stdout --log-format=logfmt",
    },
    isBackground: true,
    problemMatcher: "$teamsfx-local-tunnel-watch",
  };
  return commentJson.assign(commentJson.parse(comment), task);
}

function setUpTab(): CommentJSONValue {
  const comment = `{
    // Prepare local launch information for Tab.
    // See https://aka.ms/teamsfx-debug-tasks#debug-set-up-tab to know the details and how to customize the args.
  }`;
  const task = {
    label: "Set up tab",
    type: "teamsfx",
    command: "debug-set-up-tab",
    args: {
      baseUrl: "https://localhost:53000",
    },
  };
  return commentJson.assign(commentJson.parse(comment), task);
}

function setUpBot(): CommentJSONValue {
  const comment = `{
    // Register resources and prepare local launch information for Bot.
    // See https://aka.ms/teamsfx-debug-tasks#debug-set-up-bot to know the details and how to customize the args.
  }`;
  const existingBot = `
  {
    //// Enter you own bot information if using the existing bot. ////
    // "botId": "",
    // "botPassword": "",
  }
  `;
  const task = {
    label: "Set up bot",
    type: "teamsfx",
    command: "debug-set-up-bot",
    args: commentJson.assign(commentJson.parse(existingBot), {
      botMessagingEndpoint: "/api/messages",
    }),
  };
  return commentJson.assign(commentJson.parse(comment), task);
}

function setUpSSO(): CommentJSONValue {
  const comment = `{
    // Register resources and prepare local launch information for SSO functionality.
    // See https://aka.ms/teamsfx-debug-tasks#debug-set-up-sso to know the details and how to customize the args.
  }`;
  const existingAAD = `
  {
    //// Enter you own AAD app information if using the existing AAD app. ////
    // "objectId": "",
    // "clientId": "",
    // "clientSecret": "",
    // "accessAsUserScopeId": "
  }
  `;
  const task = {
    label: "Set up SSO",
    type: "teamsfx",
    command: "debug-set-up-sso",
    args: commentJson.parse(existingAAD),
  };
  return commentJson.assign(commentJson.parse(comment), task);
}

function buildAndUploadTeamsManifest(): CommentJSONValue {
  const comment = `
  {
    // Build and upload Teams manifest.
    // See https://aka.ms/teamsfx-debug-tasks#debug-prepare-manifest to know the details and how to customize the args.
  }`;
  const existingApp = `
  {
    //// Enter your own Teams app package path if using the existing Teams manifest. ////
    // "appPackagePath": ""
  }
  `;
  const task = {
    label: "Build & upload Teams manifest",
    type: "teamsfx",
    command: "debug-prepare-manifest",
    args: commentJson.parse(existingApp),
  };
  return commentJson.assign(commentJson.parse(comment), task);
}

function startFrontend(): Record<string, unknown> {
  return {
    label: "Start frontend",
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
    label: "Start backend",
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
    dependsOn: ["Install Azure Functions binding extensions"],
  } as Record<string, unknown>;

  if (programmingLanguage === ProgrammingLanguage.typescript) {
    (result.dependsOn as string[]).push("Watch backend");
  }

  return result;
}

function watchBackend(): Record<string, unknown> {
  return {
    label: "Watch backend",
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
    label: "Watch bot",
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
    label: "Start bot",
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
    label: "Start bot",
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

  const dependsOn: string[] = ["Start Azurite emulator"];
  if (programmingLanguage === ProgrammingLanguage.typescript) {
    dependsOn.push("Watch bot");
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
    dependsOn.push("Start frontend");
  }
  if (includeBackend) {
    dependsOn.push("Start backend");
  }
  if (includeBot) {
    dependsOn.push("Start bot");
  }
  return {
    label: "Start services",
    dependsOn,
  };
}

function startAzuriteEmulator(): Record<string, unknown> {
  return {
    label: "Start Azurite emulator",
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
    label: "install app in Teams",
    type: "shell",
    command: "exit ${command:fx-extension.install-app-in-teams}",
    presentation: {
      reveal: "never",
    },
  };
}

export function generateSpfxTasks(): Record<string, unknown>[] {
  return [
    {
      label: "Validate & install prerequisites",
      type: "teamsfx",
      command: "debug-check-prerequisites",
      args: {
        prerequisites: ["nodejs"],
      },
    },
    {
      label: "Install npm packages",
      type: "teamsfx",
      command: "debug-npm-install",
      args: {
        projects: [
          {
            cwd: "${workspaceFolder}/SPFx",
            npmInstallArgs: ["--no-audit"],
          },
        ],
        forceUpdate: false,
      },
    },
    {
      label: "gulp trust-dev-cert",
      type: "process",
      command: "node",
      args: ["${workspaceFolder}/SPFx/node_modules/gulp/bin/gulp.js", "trust-dev-cert"],
      options: {
        cwd: "${workspaceFolder}/SPFx",
      },
      dependsOn: "Install npm packages",
    },
    {
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
      dependsOn: "gulp trust-dev-cert",
    },
    {
      label: "prepare local environment",
      type: "shell",
      command: "exit ${command:fx-extension.pre-debug-check}",
    },
    {
      label: "prepare dev env",
      dependsOn: ["Validate & install prerequisites", "prepare local environment", "gulp serve"],
      dependsOrder: "sequence",
    },
    {
      label: "Terminate All Tasks",
      command: "echo ${input:terminate}",
      type: "shell",
      problemMatcher: [],
    },
  ];
}
