// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import { ContextV3, FxError, InputsWithProjectPath, ok, Result } from "@microsoft/teamsfx-api";
import "reflect-metadata";
import { Service } from "typedi";
import { ApiConnectorImpl } from "../../plugins/resource/apiconnector/plugin";
import "../connection/azureWebAppConfig";
import { ComponentNames } from "../constants";
import { ActionExecutionMW } from "../middleware/actionExecutionMW";
import "../resource/azureSql";

const apiConnectorImpl: ApiConnectorImpl = new ApiConnectorImpl();
@Service(ComponentNames.ApiConnector)
export class ApiConnector {
  name = ComponentNames.ApiConnector;
  @hooks([
    ActionExecutionMW({
      question: async (context: ContextV3, inputs: InputsWithProjectPath) => {
        return await apiConnectorImpl.generateQuestion(context, inputs);
      },
    }),
  ])
  async add(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<undefined, FxError>> {
    await apiConnectorImpl.scaffold(context, inputs);
    return ok(undefined);
  }
}
