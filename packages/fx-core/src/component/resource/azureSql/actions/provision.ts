// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  ContextV3,
  InputsWithProjectPath,
  ok,
  FunctionAction,
  ResourceContextV3,
  Effect,
  FxError,
  Result,
  traverse,
} from "@microsoft/teamsfx-api";
import { ComponentNames, ActionNames, ActionTypeFunction } from "../../../constants";
import { LoggerMW, ActionLogger } from "../../../middleware/logger";
import { RunWithCatchErrorMW, ActionErrorHandler } from "../../../middleware/runWithCatchError";
import { TelemetryMW, ActionTelemetryImplement } from "../../../middleware/telemetry";
import { ManagementClient } from "../clients/management";
import { LoadManagementConfig, removeDatabases } from "../config";
import { Constants } from "../constants";
import { ErrorMessage } from "../errors";
import { buildQuestionNode } from "../questions";
import { SqlResultFactory } from "../results";

export class ProvisionActionImplement {
  static readonly source = "SQL";
  static readonly stage = "pre-provision";
  static readonly telemetryComponentName = "fx-resource-azure-sql";
  static readonly loggerPrefix = "[SQL Component]";
  static readonly logFormatter = (message: string) =>
    `${ProvisionActionImplement.loggerPrefix} ${message}`;

  @hooks([
    TelemetryMW(
      ActionTelemetryImplement.bind(
        null,
        ProvisionActionImplement.stage,
        ProvisionActionImplement.telemetryComponentName
      )
    ),
    RunWithCatchErrorMW(ProvisionActionImplement.source, ActionErrorHandler),
    LoggerMW(ActionLogger.bind(null, ProvisionActionImplement.logFormatter)),
  ])
  static async execute(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<Effect[], FxError>> {
    const ctx = context as ResourceContextV3;
    const state = (ctx.envInfo.state[ComponentNames.AzureSQL] ??= {});
    removeDatabases(state);
    let shouldAsk;
    if (state.sqlResourceId) {
      const sqlMgrConfig = LoadManagementConfig(state);
      const sqlMgrClient = await ManagementClient.create(
        ctx.tokenProvider.azureAccountProvider,
        sqlMgrConfig
      );
      shouldAsk = !(await sqlMgrClient.existAzureSQL());
    } else {
      shouldAsk = true;
    }

    if (shouldAsk) {
      const node = buildQuestionNode();
      const res = await traverse(node, inputs, ctx.userInteraction);
      if (res.isErr()) {
        throw SqlResultFactory.UserError(
          ErrorMessage.SqlAskInputError.name,
          ErrorMessage.SqlAskInputError.message(),
          res.error
        );
      }
      state.admin = inputs[Constants.questionKey.adminName];
      state.adminPassword = inputs[Constants.questionKey.adminPassword];
    }
    return ok([{ type: "service", name: "azure", remarks: "configure azure-sql" }]);
  }

  static get(): FunctionAction {
    return {
      name: `${ComponentNames.AzureSQL}.${ActionNames.provision}`,
      type: ActionTypeFunction,
      question: (context: ContextV3, inputs: InputsWithProjectPath) => {
        return ok(undefined);
      },
      plan: (context: ContextV3, inputs: InputsWithProjectPath) => {
        return ok(["collect user inputs for sql account"]);
      },
      execute: ProvisionActionImplement.execute,
    };
  }
}
