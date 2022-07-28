// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ContextV3,
  InputsWithProjectPath,
  ok,
  FunctionAction,
  ResourceContextV3,
  Platform,
  QTreeNode,
} from "@microsoft/teamsfx-api";
import { ComponentNames, ActionTypeFunction, ActionNames } from "../../../constants";
import { AadAppForTeamsImpl } from "../../../../plugins/resource/aad/plugin";
import { convertContext } from "./utils";
import { getLocalizedString } from "../../../../common/localizeUtils";
import { Constants } from "../../../../plugins/resource/aad/constants";
import { AzureSolutionQuestionNames } from "../../../../plugins/solution/fx-solution/question";

export function GetActionDeploy(): FunctionAction {
  return {
    name: `${ComponentNames.AadApp}.${ActionNames.configure}`,
    type: ActionTypeFunction,
    question: (context: ContextV3, inputs: InputsWithProjectPath) => {
      const aadQuestions = new QTreeNode({
        type: "group",
      });
      if (inputs.platform === Platform.CLI_HELP || inputs.platform === Platform.CLI) {
        const node = new QTreeNode({
          name: Constants.INCLUDE_AAD_MANIFEST,
          type: "singleSelect",
          staticOptions: ["yes", "no"],
          title: getLocalizedString("core.aad.includeAadQuestionTitle"),
          default: "no",
        });
        aadQuestions.addChild(node);
      }
      return ok(aadQuestions);
    },
    plan: async (context: ContextV3, inputs: InputsWithProjectPath) => {
      return ok([
        {
          type: "service",
          name: "teams.microsoft.com",
          remarks: "deploy aad app",
        },
      ]);
    },
    execute: async (context: ContextV3, inputs: InputsWithProjectPath) => {
      const ctx = context as ResourceContextV3;
      const aadAppImplement = new AadAppForTeamsImpl();
      const convertCtx = convertContext(ctx, inputs);
      await aadAppImplement.deploy(convertCtx);
      const convertState = convertCtx.envInfo.state.get("fx-resource-aad-app-for-teams");
      convertState.forEach((v: any, k: string) => {
        ctx.envInfo!.state[ComponentNames.AadApp][k] = v;
      });
      return ok([
        {
          type: "service",
          name: "teams.microsoft.com",
          remarks: "deploy aad app",
        },
      ]);
    },
    condition: (context, inputs) => {
      if (
        inputs.platform === Platform.CLI_HELP ||
        (inputs.platform === Platform.CLI && inputs[Constants.INCLUDE_AAD_MANIFEST] === "yes") ||
        inputs[AzureSolutionQuestionNames.Features] !== "TabNonSsoItem.id"
      ) {
        return ok(true);
      }
      return ok(false);
    },
  };
}
