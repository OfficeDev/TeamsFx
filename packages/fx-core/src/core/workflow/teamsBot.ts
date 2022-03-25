// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, ok, Result, v2 } from "@microsoft/teamsfx-api";
import "reflect-metadata";
import { Service } from "typedi";
import { ensureSolutionSettings } from "../../plugins/solution/fx-solution/utils/solutionSettingsHelper";
import {
  Action,
  AddInstanceAction,
  ResourcePlugin,
  GroupAction,
  MaybePromise,
  CallAction,
} from "./interface";

export interface TeamsBotInputs extends v2.InputsWithProjectPath {
  language: "csharp" | "javascript" | "typescript";
  scenario: "notification" | "commandAndResponse" | "messageExtension";
  hostingResource: "azure-web-app" | "azure-function";
}

/**
 * teams bot - feature level action
 */
@Service("teams-bot")
export class TeamsBotFeature implements ResourcePlugin {
  name = "teams-bot";
  addInstance(
    context: v2.Context,
    inputs: v2.InputsWithProjectPath
  ): MaybePromise<Result<Action | undefined, FxError>> {
    const botInputs = inputs as TeamsBotInputs;
    const addInstance: AddInstanceAction = {
      name: "teams-bot.addInstance",
      type: "function",
      plan: (context: v2.Context, inputs: v2.InputsWithProjectPath) => {
        return ok(
          `ensure entry '${botInputs.hostingResource}', 'azure-bot' in projectSettings.solutionSettings.activeResourcePlugins`
        );
      },
      execute: async (
        context: v2.Context,
        inputs: v2.InputsWithProjectPath
      ): Promise<Result<undefined, FxError>> => {
        ensureSolutionSettings(context.projectSetting);
        if (
          !context.projectSetting.solutionSettings?.activeResourcePlugins.includes(
            botInputs.hostingResource
          )
        )
          context.projectSetting.solutionSettings?.activeResourcePlugins.push(
            botInputs.hostingResource
          );
        if (!context.projectSetting.solutionSettings?.activeResourcePlugins.includes("azure-bot"))
          context.projectSetting.solutionSettings?.activeResourcePlugins.push("azure-bot");
        console.log(
          `ensure entry '${botInputs.hostingResource}', 'azure-bot' in projectSettings.solutionSettings.activeResourcePlugins`
        );
        return ok(undefined);
      },
    };
    const group: GroupAction = {
      type: "group",
      actions: [
        addInstance,
        {
          type: "call",
          required: true,
          targetAction: "teams-manifest.addCapability",
          inputs: {
            capabilities: ["Bot"],
          },
        },
      ],
    };
    return ok(group);
  }
  generateCode(
    context: v2.Context,
    inputs: v2.InputsWithProjectPath
  ): MaybePromise<Result<Action | undefined, FxError>> {
    const action: CallAction = {
      name: "nodejs-bot.generateCode",
      type: "call",
      required: true,
      targetAction: "bot-scaffold.generateCode",
    };
    return ok(action);
  }
  generateBicep(
    context: v2.Context,
    inputs: v2.InputsWithProjectPath
  ): MaybePromise<Result<Action | undefined, FxError>> {
    return ok({
      type: "call",
      required: true,
      targetAction: `${inputs.hostingResource}.generateBicep`,
    });
  }
}
