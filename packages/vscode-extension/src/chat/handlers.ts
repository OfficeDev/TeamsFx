// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs-extra";
import {
  CancellationToken,
  ChatContext,
  ChatRequest,
  ChatResponseStream,
  commands,
  env,
  LanguageModelChatUserMessage,
  ProviderResult,
  Uri,
  window,
  workspace,
  ChatResultFeedback,
} from "vscode";

import { downloadDirectory } from "@microsoft/teamsfx-core/build/component/generator/utils";
import * as uuid from "uuid";

import createCommandHandler from "./commands/create/createCommandHandler";
import { ProjectMetadata } from "./commands/create/types";
import nextStepCommandHandler from "./commands/nextstep/nextstepCommandHandler";
import { TeamsChatCommand, OfficeAddinChatCommand } from "./consts";
import followupProvider from "./followupProvider";
import { defaultSystemPrompt } from "./prompts";
import { getSampleDownloadUrlInfo, verbatimCopilotInteraction } from "./utils";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetryTriggerFrom,
} from "../telemetry/extTelemetryEvents";
import { ICopilotChatResult, ITelemetryData } from "./types";
import { ChatTelemetryData } from "./telemetry";
import { localize } from "../utils/localizeUtils";
import { Correlator } from "@microsoft/teamsfx-core";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import generatecodeCommandHandler from "./commands/generatecode/generatecodeCommandHandler";
import officeAddinCreateCommandHandler from "./commands/create/officeAddinCreateCommandHandler";
import officeAddinNextStepCommandHandler from "./commands/nextstep/officeAddinNextstepCommandHandler";
import { FxError, Result } from "@microsoft/teamsfx-api";

export function chatRequestHandler(
  request: ChatRequest,
  context: ChatContext,
  response: ChatResponseStream,
  token: CancellationToken
): ProviderResult<ICopilotChatResult> {
  // Matching chat commands in the package.json
  followupProvider.clearFollowups();
  if (request.command == TeamsChatCommand.Create) {
    return createCommandHandler(request, context, response, token);
  } else if (request.command == TeamsChatCommand.NextStep) {
    return nextStepCommandHandler(request, context, response, token);
  } else {
    return defaultHandler(request, context, response, token);
  }
  return {};
}

export function officeAddinChatRequestHandler(
  request: ChatRequest,
  context: ChatContext,
  response: ChatResponseStream,
  token: CancellationToken
): ProviderResult<ICopilotChatResult> {
  followupProvider.clearFollowups();
  if (request.command == OfficeAddinChatCommand.Create) {
    return officeAddinCreateCommandHandler(request, context, response, token);
  } else if (request.command == OfficeAddinChatCommand.GenerateCode) {
    return generatecodeCommandHandler(request, context, response, token);
  } else if (request.command == OfficeAddinChatCommand.NextStep) {
    return officeAddinNextStepCommandHandler(request, context, response, token);
  } else {
    return defaultHandler(request, context, response, token);
  }
}

async function defaultHandler(
  request: ChatRequest,
  context: ChatContext,
  response: ChatResponseStream,
  token: CancellationToken
): Promise<ICopilotChatResult> {
  const chatTelemetryData = ChatTelemetryData.createByCommand("");
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.CopilotChatStart, chatTelemetryData.properties);

  const messages = [defaultSystemPrompt(), new LanguageModelChatUserMessage(request.prompt)];
  chatTelemetryData.chatMessages.push(...messages);
  await verbatimCopilotInteraction("copilot-gpt-4", messages, response, token);

  chatTelemetryData.markComplete();
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.CopilotChat,
    chatTelemetryData.properties,
    chatTelemetryData.measurements
  );
  return { metadata: { command: undefined, requestId: chatTelemetryData.requestId } };
}

export async function chatCreateCommandHandler(folderOrSample: string | ProjectMetadata) {
  // Let user choose the project folder
  let dstPath = "";
  let folderChoice: string | undefined = undefined;
  if (workspace.workspaceFolders !== undefined && workspace.workspaceFolders.length > 0) {
    folderChoice = await window.showQuickPick([
      localize("teamstoolkit.chatParticipants.create.quickPick.workspace"),
      localize("teamstoolkit.qm.browse"),
    ]);
    if (!folderChoice) {
      return;
    }
    if (folderChoice === localize("teamstoolkit.chatParticipants.create.quickPick.workspace")) {
      dstPath = workspace.workspaceFolders[0].uri.fsPath;
    }
  }
  if (dstPath === "") {
    const customFolder = await window.showOpenDialog({
      title: localize("teamstoolkit.chatParticipants.create.selectFolder.title"),
      openLabel: localize("teamstoolkit.chatParticipants.create.selectFolder.label"),
      canSelectFiles: false,
      canSelectFolders: true,
      canSelectMany: false,
    });
    if (!customFolder) {
      return;
    }
    dstPath = customFolder[0].fsPath;
  }
  try {
    if (typeof folderOrSample === "string") {
      await fs.copy(folderOrSample, dstPath);
    } else {
      const downloadUrlInfo = await getSampleDownloadUrlInfo(folderOrSample.id);
      await downloadDirectory(downloadUrlInfo, dstPath, 2, 20);
    }
    if (folderChoice !== localize("teamstoolkit.chatParticipants.create.quickPick.workspace")) {
      void commands.executeCommand("vscode.openFolder", Uri.file(dstPath));
    } else {
      void window.showInformationMessage(
        localize("teamstoolkit.chatParticipants.create.successfullyCreated")
      );
      void commands.executeCommand("workbench.view.extension.teamsfx");
    }
  } catch (error) {
    console.error("Error copying files:", error);
    void window.showErrorMessage(localize("teamstoolkit.chatParticipants.create.failToCreate"));
  }
}

export async function chatExecuteCommandHandler(
  command: string,
  requestId: string,
  ...args: unknown[]
): Promise<Result<unknown, FxError>> {
  const chatTelemetryData = ChatTelemetryData.get(requestId);
  const correlationId = uuid.v4();
  if (chatTelemetryData) {
    ExtTelemetry.sendTelemetryEvent(
      TelemetryEvent.CopilotChatClickButton,
      {
        ...chatTelemetryData.properties,
        [TelemetryProperty.CopilotChatRunCommandId]: command,
        [TelemetryProperty.CorrelationId]: correlationId,
      },
      chatTelemetryData.measurements
    );
  }
  return await commands.executeCommand<Result<unknown, FxError>>(
    command,
    correlationId,
    TelemetryTriggerFrom.CopilotChat,
    ...args
  );
}

export async function openUrlCommandHandler(url: string) {
  await env.openExternal(Uri.parse(url));
}

export function handleFeedback(e: ChatResultFeedback): void {
  const result = e.result as ICopilotChatResult;
  const telemetryData: ITelemetryData = {
    properties: {
      [TelemetryProperty.CopilotChatRequestId]: result.metadata?.requestId ?? "",
      [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.CopilotChat,
      [TelemetryProperty.CopilotChatCommand]: result.metadata?.command ?? "",
      [TelemetryProperty.CorrelationId]: Correlator.getId(),
    },
    measurements: {
      [TelemetryProperty.CopilotChatFeedbackHelpful]: e.kind,
    },
  };

  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.CopilotChatFeedback,
    telemetryData.properties,
    telemetryData.measurements
  );
}
