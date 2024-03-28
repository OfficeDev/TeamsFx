// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as tmp from "tmp";
import * as crypto from "crypto";
import * as path from "path";
import * as fs from "fs-extra";
import * as vscode from "vscode";
import {
  ChatRequest,
  ChatResponseStream,
  LanguageModelChatAssistantMessage,
  CancellationToken,
  ChatResponseFileTree,
  Uri,
} from "vscode";
import { ISkill } from "./iSkill";
import { Spec } from "./spec";
import { ExecutionResultEnum } from "./executionResultEnum";
import { fileTreeAdd } from "../../commands/create/createCommandHandler";
import { Inputs, Platform, Stage } from "@microsoft/teamsfx-api";
import { TelemetryTriggerFrom } from "../../../telemetry/extTelemetryEvents";
import { CHAT_EXECUTE_COMMAND_ID } from "../../consts";
import { CommandKey } from "../../../constants";

export class projectCreator implements ISkill {
  name: string | undefined;
  capability: string | undefined;

  constructor() {
    this.name = "Project Creator";
    this.capability = "Create a new project template";
  }

  public canInvoke(request: ChatRequest, spec: Spec): boolean {
    return (
      !!spec.userInput &&
      !!spec.appendix.codeSnippet &&
      !!spec.appendix.codeTaskBreakdown &&
      spec.appendix.codeTaskBreakdown.length > 0
    );
  }

  // eslint-disable-next-line @typescript-eslint/require-await
  public async invoke(
    languageModel: LanguageModelChatAssistantMessage,
    request: ChatRequest,
    response: ChatResponseStream,
    token: CancellationToken,
    spec: Spec
  ): Promise<{ result: ExecutionResultEnum; spec: Spec }> {
    const tempFolder = tmp.dirSync({ unsafeCleanup: true }).name;
    const tempAppName = `office-addin-${crypto.randomBytes(8).toString("hex")}`;
    const nodes = await this.buildProjectFromSpec(spec, tempFolder, tempAppName);
    response.filetree(nodes, Uri.file(path.join(tempFolder, tempAppName)));
    spec.appendix.tempAppLocation = path.join(tempFolder, tempAppName);
    return { result: ExecutionResultEnum.Success, spec: spec };
  }

  private async buildProjectFromSpec(
    spec: Spec,
    tempFolder: string,
    tempAppName: string
  ): Promise<ChatResponseFileTree[]> {
    const createInputs = {
      capabilities: spec.appendix.isCustomFunction
        ? "excel-cfshared"
        : `${spec.appendix.host}-taskpane`,
      "project-type": "office-xml-addin-type",
      "addin-host": spec.appendix.host.toLowerCase(),
      "programming-language": "javascript",
      folder: tempFolder,
      "app-name": tempAppName,
      isFromCodeGen: true,
    };
    await vscode.commands.executeCommand(
      CHAT_EXECUTE_COMMAND_ID,
      CommandKey.Create,
      TelemetryTriggerFrom.CopilotChat,
      createInputs
    );
    const rootFolder = path.join(tempFolder, tempAppName);
    const root: ChatResponseFileTree = {
      name: rootFolder,
      children: [],
    };
    // this.buildTemplateFileTree(rootFolder, root);
    this.traverseFiles(rootFolder, (fullPath) => {
      const relativePath = path.relative(rootFolder, fullPath);
      fileTreeAdd(root, relativePath);
    });
    return root.children ?? [];
  }

  private traverseFiles(dir: string, callback: (relativePath: string) => void): void {
    fs.readdirSync(dir).forEach((file) => {
      const fullPath = path.join(dir, file);
      if (fs.lstatSync(fullPath).isDirectory()) {
        this.traverseFiles(fullPath, callback);
      } else {
        callback(fullPath);
      }
    });
  }
}
