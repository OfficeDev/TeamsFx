/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

/* eslint-disable @typescript-eslint/no-unsafe-assignment */

import * as vscode from "vscode";

import { ext } from "../extensionVariables";
import { getCodeToCloudCommand } from "../subCommand/codeToCloudSlashCommand";
import {
  CREATE_SAMPLE_COMMAND_ID,
  createCommand,
  getCreateCommand,
} from "../subCommand/createSlashCommand";
import {
  getAgentHelpCommand,
  helpCommandName,
} from "../subCommand/helpSlashCommand";
import {
  DefaultNextStep,
  EXECUTE_COMMAND_ID,
  OPENURL_COMMAND_ID,
  executeCommand,
  getNextStepCommand,
  openUrlCommand,
} from "../subCommand/nextStep/command";
import { getTestCommand } from "../subCommand/testCommand";
import { agentDescription, agentName, maxFollowUps } from "./agentConsts";
import {
  LanguageModelID,
  verbatimCopilotInteraction,
} from "./copilotInteractions";
import { SlashCommandHandlerResult, SlashCommandsOwner } from "./slashCommands";

export interface ITeamsChatAgentResult extends vscode.ChatResult {
  metadata: {
    slashCommand?: string;
    sampleIds?: string[];
  };
}

export type CommandVariables = {
  languageModelID?: LanguageModelID;
  chatMessageHistory?: vscode.LanguageModelChatMessage[];
};

export type AgentRequest = {
  slashCommand?: string;
  userPrompt: string;
  variables: readonly vscode.ChatResolvedVariable[];

  context: vscode.ChatContext;
  response: vscode.ChatExtendedResponseStream;
  token: vscode.CancellationToken;

  commandVariables?: CommandVariables;
};

export interface IAgentRequestHandler {
  handleRequestOrPrompt(
    request: AgentRequest
  ): Promise<SlashCommandHandlerResult>;
  getFollowUpForLastHandledSlashCommand(
    result: vscode.ChatResult,
    context: vscode.ChatContext,
    token: vscode.CancellationToken
  ): vscode.ChatFollowup[] | undefined;
}

/**
 * Owns slash commands that are knowingly exposed to the user.
 */
const agentSlashCommandsOwner = new SlashCommandsOwner(
  {
    noInput: helpCommandName,
    default: defaultHandler,
  },
  { disableIntentDetection: true }
);
agentSlashCommandsOwner.addInvokeableSlashCommands(
  new Map([
    getCreateCommand(),
    getNextStepCommand(),
    getAgentHelpCommand(agentSlashCommandsOwner),
    getTestCommand(),
    getCodeToCloudCommand(),
  ])
);

export function registerChatAgent() {
  try {
    const participant = vscode.chat.createChatParticipant(agentName, handler);
    participant.description = agentDescription;
    participant.iconPath = vscode.Uri.joinPath(
      ext.context.extensionUri,
      "resources",
      "teams.png"
    );
    participant.followupProvider = { provideFollowups: followUpProvider };
    registerVSCodeCommands(participant);
  } catch (e) {
    console.log(e);
  }
}

async function handler(
  request: vscode.ChatRequest,
  context: vscode.ChatContext,
  stream: vscode.ChatResponseStream,
  token: vscode.CancellationToken
): Promise<vscode.ChatResult | undefined> {
  const agentRequest: AgentRequest = {
    slashCommand: request.command,
    userPrompt: request.prompt,
    variables: request.variables,
    context: context,
    response: stream,
    token: token,
  };
  let handleResult: SlashCommandHandlerResult | undefined;

  const handlers = [agentSlashCommandsOwner];
  for (const handler of handlers) {
    handleResult = await handler.handleRequestOrPrompt(agentRequest);
    if (handleResult !== undefined) {
      break;
    }
  }

  if (handleResult !== undefined) {
    handleResult.followUp = handleResult.followUp?.slice(0, maxFollowUps);
    return handleResult.chatAgentResult;
  } else {
    return undefined;
  }
}

function followUpProvider(
  result: ITeamsChatAgentResult,
  context: vscode.ChatContext,
  token: vscode.CancellationToken
): vscode.ProviderResult<vscode.ChatFollowup[]> {
  const providers = [agentSlashCommandsOwner];

  let followUp: vscode.ChatFollowup[] | undefined;
  for (const provider of providers) {
    followUp = provider.getFollowUpForLastHandledSlashCommand(
      result,
      context,
      token
    );
    if (followUp !== undefined) {
      break;
    }
  }
  followUp = followUp ?? [];
  if (followUp.length === 0) {
    followUp.push(DefaultNextStep);
  }
  return followUp;
}

function getCommands(
  _context: vscode.ChatContext,
  _token: vscode.CancellationToken
): vscode.ProviderResult<vscode.ChatCommand[]> {
  return agentSlashCommandsOwner.getSlashCommands().map(([name, config]) => ({
    name: name,
    description: config.shortDescription,
  }));
}

async function defaultHandler(
  request: AgentRequest
): Promise<SlashCommandHandlerResult> {
  const defaultSystemPrompt = `You are an expert in Teams Toolkit Extension for VS Code. The user wants to use Teams Toolkit Extension for VS Code. They want to use them to solve a problem or accomplish a task. Your job is to help the user learn about how they can use Teams Toolkit Extension for VS Code to solve a problem or accomplish a task. Do not suggest using any other tools other than what has been previously mentioned. Assume the the user is only interested in using Teams Toolkit Extension to develop teams app. Finally, do not overwhelm the user with too much information. Keep responses short and sweet.`;

  request.commandVariables = { languageModelID: "copilot-gpt-4" };
  const { copilotResponded } = await verbatimCopilotInteraction(
    defaultSystemPrompt,
    request
  );
  if (!copilotResponded) {
    request.response.report({
      content: vscode.l10n.t("Sorry, I can't help with that right now.\n"),
    });
    return {
      chatAgentResult: { metadata: { slashCommand: "" } },
      followUp: [],
    };
  } else {
    return {
      chatAgentResult: { metadata: { slashCommand: "" } },
      followUp: [],
    };
  }
}

function registerVSCodeCommands(participant: vscode.ChatParticipant) {
  ext.context.subscriptions.push(
    participant,
    vscode.commands.registerCommand(CREATE_SAMPLE_COMMAND_ID, createCommand),
    vscode.commands.registerCommand(EXECUTE_COMMAND_ID, executeCommand),
    vscode.commands.registerCommand(OPENURL_COMMAND_ID, openUrlCommand)
  );
}
