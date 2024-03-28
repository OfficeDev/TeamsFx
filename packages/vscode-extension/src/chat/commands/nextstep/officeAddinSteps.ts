// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CommandKey } from "../../../constants";
import { CHAT_EXECUTE_COMMAND_ID } from "../../consts";
import {
  isDidNoActionAfterScaffolded,
  isFirstInstalled,
  isHaveReadMe,
  isPrequisitesCheckSucceeded,
  isProjectOpened,
  isPublishedSucceededBefore,
} from "./condition";
import { NextStep, WholeStatus } from "./types";

// TODO: align the description with PM
export const officeAddinSteps: NextStep[] = [
  {
    title: "Teams Toolkit",
    description: `Teams Toolkit makes it simple to get started with app development for Microsoft Teams using Visual Studio Code. You can start with a project template for a common custom app built for your org (LOB app) scenarios or from a sample. You can save setup time with automated app registration and configuration. You can run and debug your app in Teams directly from familiar tools. You can smart defaults for hosting in Azure using infrastructure-as-code and Bicep. You can create unique configurations like dev, test, and prod using the environment features.`,
    docLink:
      "https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/install-teams-toolkit?tabs=vscode&pivots=visual-studio-code-v5",
    commands: [
      {
        title: "Open Get-Started Page",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.OpenWelcome],
      },
      {
        title: "Open Document",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.OpenDocument],
      },
    ],
    followUps: [],
    condition: (status: WholeStatus) => isFirstInstalled(status),
    priority: 0,
  },
  {
    title: "New Project",
    description:
      "You can start with built-in Office Add-in templates or start with official Office Add-in samples in Teams Toolkit.",
    docLink:
      "https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/create-new-project?pivots=visual-studio-code-v5",
    commands: [
      {
        title: "Open Sample Gallery",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.OpenSamples],
      },
    ],
    followUps: [
      {
        label: "@officeaddin /create",
        command: "create",
        prompt: "",
      },
    ],
    condition: (status: WholeStatus) => !isProjectOpened(status),
    priority: 0,
  },
  {
    title: "Prerequisites",
    description: (status: WholeStatus) =>
      `Ensure the following requirements are met before you start building your Teams app. It seems you met the prerequisites error: ${
        status.machineStatus.resultOfPrerequistes || ""
      }. You can fix it and try again.`,
    docLink:
      "https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/tools-prerequisites",
    commands: [
      {
        title: "Check Prerequisites Again",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.ValidateGetStartedPrerequisites],
      },
    ],
    followUps: [],
    condition: (status: WholeStatus) =>
      isProjectOpened(status) && !isPrequisitesCheckSucceeded(status),
    priority: 0,
  },
  {
    title: "Summary of README",
    description: (status: WholeStatus) => {
      // readme must exist because the condition has checked it
      const readme = status.projectOpened!.readmeContent!;
      let description = "";
      let findFirstSharp = false;
      for (const line of readme.split("\n")) {
        if (line.trim().startsWith("#")) {
          findFirstSharp = true;
        }
        if (!findFirstSharp) {
          continue;
        }
        if (line.toLocaleLowerCase().includes("prerequisite")) {
          break;
        }
        description += line.trim() + " ";
      }
      return description;
    },
    commands: [
      {
        title: "Open README",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.OpenReadMe],
      },
    ],
    followUps: [],
    condition: (status: WholeStatus) =>
      isProjectOpened(status) &&
      isPrequisitesCheckSucceeded(status) &&
      isDidNoActionAfterScaffolded(status) &&
      isHaveReadMe(status),
    priority: 1,
  },
];
