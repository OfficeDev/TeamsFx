// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios from "axios";
import * as fs from "fs-extra";
import * as path from "path";
import * as tmp from "tmp";
import {
  CancellationToken,
  ChatContext,
  ChatRequest,
  ChatResponseFileTree,
  ChatResponseStream,
  ChatResult,
  LanguageModelChatUserMessage,
  Uri,
} from "vscode";

import { sampleProvider } from "@microsoft/teamsfx-core";
import {
  getSampleFileInfo,
  runWithLimitedConcurrency,
  sendRequestWithRetry,
} from "@microsoft/teamsfx-core/build/component/generator/utils";

import { TelemetryEvent } from "../../../telemetry/extTelemetryEvents";
import {
  CHAT_CREATE_SAMPLE_COMMAND_ID,
  CHAT_EXECUTE_COMMAND_ID,
  TeamsChatCommand,
} from "../../consts";
import {
  brieflyDescribeProjectSystemPrompt,
  describeProjectSystemPrompt,
  getProjectMatchSystemPrompt,
} from "../../prompts";
import {
  getCopilotResponseAsString,
  getSampleDownloadUrlInfo,
  verbatimCopilotInteraction,
} from "../../utils";
import * as teamsTemplateMetadata from "./templateMetadata.json";
import { ProjectMetadata } from "./types";
import { ChatTelemetryData } from "../../telemetry";
import { IChatTelemetryData, ICopilotChatResult } from "../../types";
import * as util from "util";
import { localize } from "../../../utils/localizeUtils";
import { ExtTelemetry } from "../../../telemetry/extTelemetry";
import { CommandKey } from "../../../constants";

export default async function createCommandHandler(
  request: ChatRequest,
  context: ChatContext,
  response: ChatResponseStream,
  token: CancellationToken
): Promise<ICopilotChatResult> {
  const chatTelemetryData = ChatTelemetryData.createByCommand(TeamsChatCommand.Create);
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.CopilotChatStart, chatTelemetryData.properties);

  const matchedResult = await matchProject(request, token, chatTelemetryData);

  if (matchedResult.length === 0) {
    response.markdown(
      "No matching templates or samples found. Try a different app description or explore other templates.\n"
    );
    chatTelemetryData.markComplete();
    ExtTelemetry.sendTelemetryEvent(
      TelemetryEvent.CopilotChat,
      chatTelemetryData.properties,
      chatTelemetryData.measurements
    );
    return {
      metadata: {
        command: TeamsChatCommand.Create,
        requestId: chatTelemetryData.requestId,
      },
    };
  }
  if (matchedResult.length === 1) {
    const firstMatch = matchedResult[0];
    const describeProjectChatMessages = [
      describeProjectSystemPrompt,
      new LanguageModelChatUserMessage(
        `The project you are looking for is '${JSON.stringify(firstMatch)}'.`
      ),
    ];
    chatTelemetryData.chatMessages.push(...describeProjectChatMessages);

    await verbatimCopilotInteraction(
      "copilot-gpt-3.5-turbo",
      describeProjectChatMessages,
      response,
      token
    );
    if (firstMatch.type === "sample") {
      const folder = await showFileTree(firstMatch, response);
      const sampleTitle = localize("teamstoolkit.chatParticipants.create.sample");
      response.button({
        command: CHAT_CREATE_SAMPLE_COMMAND_ID,
        arguments: [folder],
        title: sampleTitle,
      });
    } else if (firstMatch.type === "template") {
      const templateTitle = localize("teamstoolkit.chatParticipants.create.template");
      response.button({
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.Create, chatTelemetryData.requestId, firstMatch.data],
        title: templateTitle,
      });
    }

    chatTelemetryData.markComplete();
    ExtTelemetry.sendTelemetryEvent(
      TelemetryEvent.CopilotChat,
      chatTelemetryData.properties,
      chatTelemetryData.measurements
    );
    return {
      metadata: {
        command: TeamsChatCommand.Create,
        requestId: chatTelemetryData.requestId,
      },
    };
  } else {
    response.markdown(
      `We've found ${
        matchedResult.slice(0, 3).length
      } projects that match your description. Take a look at them below.\n`
    );
    for (const project of matchedResult.slice(0, 3)) {
      response.markdown(`- ${project.name}: `);

      const brieflyDescribeProjectChatMessages = [
        brieflyDescribeProjectSystemPrompt,
        new LanguageModelChatUserMessage(
          `The project you are looking for is '${JSON.stringify(project)}'.`
        ),
      ];
      chatTelemetryData.chatMessages.push(...brieflyDescribeProjectChatMessages);

      await verbatimCopilotInteraction(
        "copilot-gpt-3.5-turbo",
        brieflyDescribeProjectChatMessages,
        response,
        token
      );
      if (project.type === "sample") {
        const sampleTitle = localize("teamstoolkit.chatParticipants.create.sample");
        response.button({
          command: CHAT_CREATE_SAMPLE_COMMAND_ID,
          arguments: [project],
          title: sampleTitle,
        });
      } else if (project.type === "template") {
        const templateTitle = localize("teamstoolkit.chatParticipants.create.template");
        response.button({
          command: CHAT_EXECUTE_COMMAND_ID,
          arguments: [CommandKey.Create, chatTelemetryData.requestId, project.data],
          title: templateTitle,
        });
      }
    }

    chatTelemetryData.markComplete();
    ExtTelemetry.sendTelemetryEvent(
      TelemetryEvent.CopilotChat,
      chatTelemetryData.properties,
      chatTelemetryData.measurements
    );
    return {
      metadata: {
        command: TeamsChatCommand.Create,
        requestId: chatTelemetryData.requestId,
      },
    };
  }
}

async function matchProject(
  request: ChatRequest,
  token: CancellationToken,
  telemetryMetadata: IChatTelemetryData
): Promise<ProjectMetadata[]> {
  const allProjectMetadata = [...getTeamsTemplateMetadata(), ...(await getTeamsSampleMetadata())];
  const messages = [
    getProjectMatchSystemPrompt(allProjectMetadata),
    new LanguageModelChatUserMessage(request.prompt),
  ];

  telemetryMetadata.chatMessages.push(...messages);

  const response = await getCopilotResponseAsString("copilot-gpt-3.5-turbo", messages, token);
  const matchedProjectId: string[] = [];
  if (response) {
    try {
      const responseJson = JSON.parse(response);
      if (responseJson && responseJson.app) {
        matchedProjectId.push(...(responseJson.app as string[]));
      }
    } catch (e) {}
  }
  const result: ProjectMetadata[] = [];
  for (const id of matchedProjectId) {
    const matchedProject = allProjectMetadata.find((config) => config.id === id);
    if (matchedProject) {
      result.push(matchedProject);
    }
  }
  return result;
}

function getTeamsTemplateMetadata(): ProjectMetadata[] {
  return teamsTemplateMetadata.map((config) => {
    return {
      id: config.id,
      type: "template",
      platform: "Teams",
      name: config.name,
      description: config.description,
      data: {
        capabilities: config.id,
        "project-type": config["project-type"],
      },
    };
  });
}

async function getTeamsSampleMetadata(): Promise<ProjectMetadata[]> {
  const sampleCollection = await sampleProvider.SampleCollection;
  const result: ProjectMetadata[] = [];
  for (const sample of sampleCollection.samples) {
    result.push({
      id: sample.id,
      type: "sample",
      platform: "Teams",
      name: sample.title,
      description: sample.fullDescription,
    });
  }
  return result;
}

export async function showFileTree(
  projectMetadata: ProjectMetadata,
  response: ChatResponseStream
): Promise<string> {
  response.markdown(
    "\nWe've found a sample project that matches your description. Take a look at it below."
  );
  const downloadUrlInfo = await getSampleDownloadUrlInfo(projectMetadata.id);
  const { samplePaths, fileUrlPrefix } = await getSampleFileInfo(downloadUrlInfo, 2);
  const tempFolder = tmp.dirSync({ unsafeCleanup: true }).name;
  const nodes = await buildFileTree(
    fileUrlPrefix,
    samplePaths,
    tempFolder,
    downloadUrlInfo.dir,
    2,
    20
  );
  response.filetree(nodes, Uri.file(path.join(tempFolder, downloadUrlInfo.dir)));
  return path.join(tempFolder, downloadUrlInfo.dir);
}

async function buildFileTree(
  fileUrlPrefix: string,
  samplePaths: string[],
  dstPath: string,
  relativeFolderName: string,
  retryLimits: number,
  concurrencyLimits: number
): Promise<ChatResponseFileTree[]> {
  const root: ChatResponseFileTree = {
    name: relativeFolderName,
    children: [],
  };
  const downloadCallback = async (samplePath: string) => {
    const file = (await sendRequestWithRetry(async () => {
      return await axios.get(fileUrlPrefix + samplePath, {
        responseType: "arraybuffer",
      });
    }, retryLimits)) as unknown as any;
    const relativePath = path.relative(`${relativeFolderName}/`, samplePath);
    const filePath = path.join(dstPath, samplePath);
    fileTreeAdd(root, relativePath);
    await fs.ensureFile(filePath);
    await fs.writeFile(filePath, Buffer.from(file.data));
  };
  await runWithLimitedConcurrency(samplePaths, downloadCallback, concurrencyLimits);
  return root.children ?? [];
}

export function fileTreeAdd(root: ChatResponseFileTree, relativePath: string) {
  const filename = path.basename(relativePath);
  const folderName = path.dirname(relativePath);
  const segments = path.sep === "\\" ? folderName.split("\\") : folderName.split("/");
  let parent = root;
  for (let i = 0; i < segments.length; i++) {
    const segment = segments[i];
    if (segment === ".") {
      continue;
    }
    let child = parent.children?.find((child) => child.name === segment);
    if (!child) {
      child = {
        name: segment,
        children: [],
      };
      parent.children?.push(child);
    }
    parent = child;
  }
  parent.children?.push({
    name: filename,
  });
}
