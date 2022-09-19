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
): Record<string, unknown>[] {
  /**
   * Referenced by launch.json
   *   - Start Teams App Locally
   *
   * Referenced inside tasks.json
   *   - Validate & Install prerequisites
   *   - Install NPM packages
   *   - Install Azure Functions binding extensions
   *   - Start local tunnel
   *   - Set up Tab
   *   - Set up Bot
   *   - Set up SSO
   *   - Build & Upload Teams manifest
   *   - Start services
   *   - Start frontend
   *   - Start backend
   *   - Watch backend
   *   - Start bot
   */
  const tasks: Record<string, unknown>[] = [
    startTeamsAppLocally(includeFrontend, includeBackend, includeBot, includeSSO),
    validateAndInstallPrerequisites(
      includeFrontend,
      includeBackend,
      includeBot,
      includeFuncHostedBot
    ),
    installNPMpackages(includeFrontend, includeBackend, includeBot),
  ];

  if (includeBackend) {
    tasks.push(installAzureFunctionsBindingExtensions());
  }

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
): Record<string, unknown>[] {
  /**
   * Referenced by launch.json
   *   - Start Teams App Locally
   *   - Start Teams App Locally & Install App
   *
   * Referenced inside tasks.json
   *   - Validate & Install prerequisites
   *   - Install NPM packages
   *   - Install Azure Functions binding extensions
   *   - Start local tunnel
   *   - Set up Tab
   *   - Set up Bot
   *   - Set up SSO
   *   - Build & Upload Teams manifest
   *   - Start services
   *   - Start frontend
   *   - Start backend
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
  } else {
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
    dependsOn: ["Validate & Install prerequisites", "Install NPM packages"],
    dependsOrder: "sequence",
  };
  if (includeBackend) {
    result.dependsOn.push("Install Azure Functions binding extensions");
  }
  if (includeBot) {
    result.dependsOn.push("Start local tunnel");
  }
  if (includeFrontend) {
    result.dependsOn.push("Set up Tab");
  }
  if (includeBot) {
    result.dependsOn.push("Set up Bot");
  }
  if (includeSSO) {
    result.dependsOn.push("Set up SSO");
  }
  result.dependsOn.push("Build & Upload Teams manifest", "Start services");

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
): Record<string, unknown> {
  const prerequisites = ["nodejs", "m365Account"];
  const comments: string[] = [];
  if (includeFrontend) {
    prerequisites.push("devCert");
    comments.push("53000, // tab service port");
  }
  if (includeBackend) {
    prerequisites.push("func", "dotnet");
    comments.push("7071, // backend service port", "9229, // backend debug port");
  }
  if (includeFuncHostedBot && !includeBackend) {
    prerequisites.push("func");
  }
  if (includeBot) {
    prerequisites.push("ngrok");
    comments.push("3978, // bot service port", "9239, // bot debug port");
  }
  prerequisites.push("portOccupancy");
  const comment = `
  [
    ${comments.join("\n  ")}
  ]
  `;

  return {
    label: "Validate & Install prerequisites",
    type: "teamsfx",
    command: "debug-check-prerequisites",
    args: {
      prerequisites,
      portOccupancy: commentJson.parse(comment),
    },
  };
}

function installNPMpackages(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean
): Record<string, unknown> {
  const result = {
    label: "Install NPM packages",
    type: "teamsfx",
    command: "debug-npm-install",
    args: {
      projects: [] as Record<string, unknown>[],
      forceUpdate: false,
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

  return result;
}

function installAzureFunctionsBindingExtensions(): Record<string, unknown> {
  return {
    label: "Install Azure Functions binding extensions",
    type: "shell",
    command: "dotnet build extensions.csproj -o ./bin --ignore-failed-sources",
    options: {
      cwd: "${workspaceFolder}/api",
      env: {
        PATH: "${command:fx-extension.get-dotnet-path}${env:PATH}",
      },
    },
  };
}

function startLocalTunnel(): Record<string, unknown> {
  return {
    label: "Start local tunnel",
    type: "teamsfx",
    command: "debug-start-local-tunnel",
    args: {
      configFile: ".fx/configs/ngrok.yml",
      binFolder: "${teamsfx:ngrokBinFolder}",
      reuse: false,
    },
    isBackground: true,
    problemMatcher: "$teamsfx-local-tunnel-watch",
  };
}

function setUpTab(): Record<string, unknown> {
  return {
    label: "Set up Tab",
    type: "teamsfx",
    command: "debug-set-up-tab",
    args: {
      baseUrl: "https://localhost:53000",
    },
  };
}

function setUpBot(): Record<string, unknown> {
  const comment = `
  {
    //// Enter you own bot information if using the existing bot. ////
    // "botId": "",
    // "botPassword": "",
  }
  `;

  return {
    label: "Set up Bot",
    type: "teamsfx",
    command: "debug-set-up-bot",
    args: commentJson.assign(commentJson.parse(comment), {
      botMessagingEndpoint: "${teamsfx:botTunnelEndpoint}/api/messages",
    }),
  };
}

function setUpSSO(): Record<string, unknown> {
  const comment = `
  {
    //// Enter you own AAD app information if using the existing AAD app. ////
    // "objectId": "",
    // "clientId": "",
    // "clientSecret": "",
    // "accessAsUserScopeId": "
  }
  `;

  return {
    label: "Set up SSO",
    type: "teamsfx",
    command: "debug-set-up-sso",
    args: commentJson.parse(comment),
  };
}

function buildAndUploadTeamsManifest(): Record<string, unknown> {
  const comment = `
  {
    //// Enter your own Teams manifest app package path if using the existing Teams manifest app package. ////
    // "manifestPackagePath": ""
  }
  `;

  return {
    label: "Build & Upload Teams manifest",
    type: "teamsfx",
    command: "debug-prepare-manifest",
    args: commentJson.parse(comment),
  };
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
  } as Record<string, unknown>;

  if (programmingLanguage === ProgrammingLanguage.typescript) {
    result.dependsOn = "Watch backend";
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
