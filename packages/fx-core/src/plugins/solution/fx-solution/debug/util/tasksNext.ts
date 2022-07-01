// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { ProductName } from "@microsoft/teamsfx-api";
import { isOfficeAddinEnabled } from "../../../../../common";
import { ProgrammingLanguage } from "../../../../../common/local/constants";

// TODO: add spfx tasks with "validate-local-prerequisites"
export function generateTasks(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean,
  includeFuncHostedBot: boolean,
  programmingLanguage: string,
  includeOfficeAddin: boolean
): Record<string, unknown>[] {
  /**
   * Referenced by launch.json
   *   - Pre Debug Check & Start All
   *
   * Referenced inside tasks.json
   *   - validate local prerequisites
   *   - start ngrok
   *   - prepare local environment
   *   - Start All
   *   - Start Frontend
   *   - Start Backend
   *   - Watch Backend
   *   - Start Bot
   */
  const tasks: Record<string, unknown>[] = [
    preDebugCheckAndStartAll(includeBot),
    validateLocalPrerequisites(),
  ];

  if (includeBot) {
    tasks.push(startNgrok());
  }

  tasks.push(prepareLocalEnvironment());

  tasks.push(startAll(includeFrontend, includeBackend, includeBot));

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

  if (isOfficeAddinEnabled() && includeOfficeAddin) {
    // adds entries into tasks
  }

  return tasks;
}

export function generateM365Tasks(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean,
  programmingLanguage: string
): Record<string, unknown>[] {
  /**
   * Referenced by launch.json
   *   - Pre Debug Check & Start All
   *   - Pre Debug Check & Start All & Install App
   *
   * Referenced inside tasks.json
   *   - validate local prerequisites
   *   - start ngrok
   *   - prepare local environment
   *   - Start All
   *   - install app in Teams
   *   - Start Frontend
   *   - Start Backend
   *   - Watch Backend
   *   - Start Bot
   */
  const tasks: Record<string, unknown>[] = [
    preDebugCheckAndStartAll(includeBot),
    preDebugCheckAndStartAllAndInstallApp(includeBot),
    validateLocalPrerequisites(),
  ];

  if (includeBot) {
    tasks.push(startNgrok());
  }

  tasks.push(prepareLocalEnvironment());

  tasks.push(startAll(includeFrontend, includeBackend, includeBot));

  tasks.push(installAppInTeams());

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
    tasks.push(startBot(includeFrontend));
  }

  return tasks;
}

export function mergeTasks(
  existingData: Record<string, unknown>,
  newData: Record<string, unknown>
): Record<string, unknown> {
  const mergedData = {} as Record<string, unknown>;
  Object.assign(mergedData, existingData);

  if (mergedData.version === undefined) {
    mergedData.version = "2.0.0";
  }

  if (mergedData.tasks === undefined) {
    mergedData.tasks = newData.tasks;
  } else {
    const existingTasks = mergedData.tasks as Record<string, unknown>[];
    const newTasks = (newData.tasks ?? []) as Record<string, unknown>[];
    const keptTasks = [];
    for (const existingTask of existingTasks) {
      if (
        !newTasks.some(
          (newTask) => existingTask.label === newTask.label && existingTask.type === newTask.type
        )
      ) {
        keptTasks.push(existingTask);
      }
    }
    mergedData.tasks = [...keptTasks, ...newTasks];
  }

  if (mergedData.inputs === undefined) {
    mergedData.inputs = newData.inputs;
  } else {
    const existingInputs = mergedData.inputs as Record<string, unknown>[];
    const newInputs = (newData.inputs ?? []) as Record<string, unknown>[];
    const keptInputs = [];
    for (const existingInput of existingInputs) {
      if (
        !newInputs.some(
          (newInput) => existingInput.id === newInput.id && existingInput.type === newInput.type
        )
      ) {
        keptInputs.push(existingInput);
      }
    }
    mergedData.inputs = [...keptInputs, ...newInputs];
  }

  return mergedData;
}

function preDebugCheckAndStartAll(includeBot: boolean): Record<string, unknown> {
  return {
    label: "Pre Debug Check & Start All",
    dependsOn: includeBot
      ? ["validate local prerequisites", "start ngrok", "prepare local environment", "Start All"]
      : ["validate local prerequisites", "prepare local environment", "Start All"],
    dependsOrder: "sequence",
  };
}

function preDebugCheckAndStartAllAndInstallApp(includeBot: boolean): Record<string, unknown> {
  return {
    label: "Pre Debug Check & Start All & Install App",
    dependsOn: includeBot
      ? [
          "validate local prerequisites",
          "start ngrok",
          "prepare local environment",
          "Start All",
          "install app in Teams",
        ]
      : [
          "validate local prerequisites",
          "prepare local environment",
          "Start All",
          "install app in Teams",
        ],
    dependsOrder: "sequence",
  };
}

function validateLocalPrerequisites(): Record<string, unknown> {
  return {
    label: "validate local prerequisites",
    type: "shell",
    command: "exit ${command:fx-extension.validate-local-prerequisites}",
    presentation: {
      reveal: "never",
    },
  };
}

function prepareLocalEnvironment(): Record<string, unknown> {
  return {
    label: "prepare local environment",
    type: "shell",
    command: "exit ${command:fx-extension.pre-debug-check}",
    presentation: {
      reveal: "never",
    },
  };
}

function startFrontend(): Record<string, unknown> {
  return {
    label: "Start Frontend",
    type: "shell",
    command: "npm run dev:teamsfx",
    isBackground: true,
    problemMatcher: "$teamsfx-frontend-watch",
    options: {
      cwd: "${workspaceFolder}/tabs",
    },
  };
}

function startBackend(programmingLanguage: string): Record<string, unknown> {
  const result = {
    label: "Start Backend",
    type: "shell",
    command: "npm run dev:teamsfx",
    isBackground: true,
    problemMatcher: "$teamsfx-backend-watch",
    options: {
      cwd: "${workspaceFolder}/api",
      env: {
        PATH: "${command:fx-extension.get-func-path}${env:PATH}",
      },
    },
    presentation: {
      reveal: "silent",
    },
  } as Record<string, unknown>;

  if (programmingLanguage === ProgrammingLanguage.typescript) {
    result.dependsOn = "Watch Backend";
  }

  return result;
}

function watchBackend(): Record<string, unknown> {
  return {
    label: "Watch Backend",
    type: "shell",
    command: "npm run watch:teamsfx",
    isBackground: true,
    problemMatcher: "$tsc-watch",
    options: {
      cwd: "${workspaceFolder}/api",
    },
    presentation: {
      reveal: "silent",
    },
  };
}

function watchFuncHostedBot(): Record<string, unknown> {
  return {
    label: "Watch Bot",
    type: "shell",
    command: "npm run watch:teamsfx",
    isBackground: true,
    problemMatcher: "$tsc-watch",
    options: {
      cwd: "${workspaceFolder}/bot",
    },
    presentation: {
      reveal: "silent",
    },
  };
}

function startBot(includeFrontend: boolean): Record<string, unknown> {
  const result = {
    label: "Start Bot",
    type: "shell",
    command: "npm run dev:teamsfx",
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
        beginsPattern: "[nodemon] starting",
        endsPattern: "restify listening to|Bot/ME service listening at|[nodemon] app crashed",
      },
    },
    options: {
      cwd: "${workspaceFolder}/bot",
    },
  } as Record<string, unknown>;

  if (includeFrontend) {
    result.presentation = { reveal: "silent" };
  }

  return result;
}

function startFuncHostedBot(
  includeFrontend: boolean,
  programmingLanguage: string
): Record<string, unknown> {
  const result = {
    label: "Start Bot",
    type: "shell",
    command: "npm run dev:teamsfx",
    isBackground: true,
    problemMatcher: "$teamsfx-func-hosted-bot-watch",
    options: {
      cwd: "${workspaceFolder}/bot",
      env: {
        PATH: "${command:fx-extension.get-func-path}${env:PATH}",
      },
    },
  } as Record<string, unknown>;

  if (includeFrontend) {
    result.presentation = { reveal: "silent" };
  }

  const dependsOn: string[] = ["Start Azurite Emulator"];
  if (programmingLanguage === ProgrammingLanguage.typescript) {
    dependsOn.push("Watch Bot");
  }
  result.dependsOn = dependsOn;

  return result;
}

function startNgrok(): Record<string, unknown> {
  return {
    label: "start ngrok",
    dependsOn: `${ProductName}: ngrok start`,
  };
}

function startAll(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean
): Record<string, unknown> {
  const dependsOn: string[] = [];
  if (includeFrontend) {
    dependsOn.push("Start Frontend");
  }
  if (includeBackend) {
    dependsOn.push("Start Backend");
  }
  if (includeBot) {
    dependsOn.push("Start Bot");
  }
  return {
    label: "Start All",
    dependsOn,
  };
}

function startAzuriteEmulator(): Record<string, unknown> {
  return {
    label: "Start Azurite Emulator",
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
