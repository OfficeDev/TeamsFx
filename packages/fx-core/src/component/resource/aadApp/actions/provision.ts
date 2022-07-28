// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ContextV3,
  InputsWithProjectPath,
  ok,
  FunctionAction,
  ResourceContextV3,
} from "@microsoft/teamsfx-api";
import { ComponentNames, ActionTypeFunction, ActionNames } from "../../../constants";
import { AadAppForTeamsImpl } from "../../../../plugins/resource/aad/plugin";
import { convertContext } from "./utils";

export function GetActionProvision(): FunctionAction {
  return {
    name: `${ComponentNames.AadApp}.${ActionNames.provision}`,
    type: ActionTypeFunction,
    plan: async (context: ContextV3, inputs: InputsWithProjectPath) => {
      return ok([
        {
          type: "service",
          name: "teams.microsoft.com",
          remarks: "create or update aad app",
        },
      ]);
    },
    execute: async (context: ContextV3, inputs: InputsWithProjectPath) => {
      const ctx = context as ResourceContextV3;
      ctx.envInfo!.state[ComponentNames.AadApp] ??= {};
      const aadAppImplement = new AadAppForTeamsImpl();
      const convertCtx = convertContext(ctx, inputs);
      await aadAppImplement.provisionUsingManifest(convertCtx);

      const convertState = convertCtx.envInfo.state.get("fx-resource-aad-app-for-teams");
      convertState.forEach((v: any, k: string) => {
        ctx.envInfo!.state[ComponentNames.AadApp][k] = v;
      });

      return ok([
        {
          type: "service",
          name: "teams.microsoft.com",
          remarks: "create or update aad app",
        },
      ]);
    },
  };
}
