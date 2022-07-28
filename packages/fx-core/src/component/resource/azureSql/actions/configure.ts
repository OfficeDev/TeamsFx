// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  ContextV3,
  InputsWithProjectPath,
  ok,
  FunctionAction,
  ResourceContextV3,
  v3,
  AzureAccountProvider,
  LogProvider,
  Result,
  FxError,
  Effect,
} from "@microsoft/teamsfx-api";
import Container from "typedi";
import {
  ComponentNames,
  ActionNames,
  ActionTypeFunction,
  TelemetryConstants,
} from "../../../constants";
import { ActionLogger, LoggerMW } from "../../../middleware/logger";
import { ProgressBarMW } from "../../../middleware/progressbar";
import { ActionErrorHandler, RunWithCatchErrorMW } from "../../../middleware/runWithCatchError";
import { ActionTelemetryImplement, TelemetryMW } from "../../../middleware/telemetry";
import { ActionContext } from "../../../middleware/types";
import { ManagementClient } from "../clients/management";
import { SqlClient } from "../clients/sql";
import { loadDatabases, LoadManagementConfig, LoadSqlConfig } from "../config";
import { Constants, HelpLinks, Telemetry, Message } from "../constants";
import { ErrorMessage } from "../errors";
import { SqlResultFactory } from "../results";
import { parseToken, TokenInfo, UserType } from "../utils/common";
export class ConfigureActionImplement {
  static readonly source = "SQL";
  static readonly stage = "post-provision";
  static readonly telemetryComponentName = "fx-resource-azure-sql";
  static readonly progressTitle = "Configuring SQL";
  static readonly progressMessage = {
    addAadmin: "Configure aad admin for SQL",
    addUser: "Configure database user",
  };
  static readonly loggerPrefix = "[SQL Component]";
  static readonly logFormatter = (message: string) =>
    `${ConfigureActionImplement.loggerPrefix} ${message}`;

  @hooks([
    TelemetryMW(
      ActionTelemetryImplement.bind(
        null,
        ConfigureActionImplement.stage,
        ConfigureActionImplement.telemetryComponentName
      )
    ),
    RunWithCatchErrorMW(ConfigureActionImplement.source, ActionErrorHandler),
    ProgressBarMW(
      ConfigureActionImplement.progressTitle,
      Object.keys(ConfigureActionImplement.progressMessage).length
    ),
    LoggerMW(ActionLogger.bind(null, ConfigureActionImplement.logFormatter)),
  ])
  static async execute(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<Effect[], FxError>> {
    const ctx = context as ResourceContextV3;
    const actionContext = context as ActionContext;
    const solutionConfig = ctx.envInfo.state.solution as v3.AzureSolutionConfig;
    const state = ctx.envInfo.state[ComponentNames.AzureSQL];
    const sqlMgrConfig = LoadManagementConfig(state);
    const sqlMgrClient = await ManagementClient.create(
      ctx.tokenProvider.azureAccountProvider,
      sqlMgrConfig
    );

    actionContext.logger?.info(Message.addFirewall);
    await sqlMgrClient.addLocalFirewallRule();

    const adminInfo = await UtilFunctions.parseLoginToken(ctx.tokenProvider.azureAccountProvider);
    actionContext.progressBar?.next(ConfigureActionImplement.progressMessage.addAadmin);
    const existAdmin = await UtilFunctions.CheckAndSetAadAdmin(
      sqlMgrClient,
      adminInfo.name,
      adminInfo.objectId,
      solutionConfig.tenantId
    );
    if (existAdmin) {
      actionContext.logger?.info(Message.skipAddAadAdmin);
    } else {
      actionContext.logger?.info(Message.addSqlAadAdmin);
    }

    // update outputKeys
    const databases = loadDatabases(state);
    const resource = Container.get(ComponentNames.AzureSQL) as any;
    resource.finalOutputKeys.push(...Object.keys(databases));

    const identity = UtilFunctions.getIdentity(ctx);
    const sqlConfig = LoadSqlConfig(state, identity);
    const skipAddingUser = await UtilFunctions.getSkipAddingUser(
      solutionConfig,
      ctx.tokenProvider.azureAccountProvider
    );
    actionContext.telemetry?.addProperty(
      Telemetry.properties.skipAddingUser,
      skipAddingUser ? TelemetryConstants.values.yes : TelemetryConstants.values.no
    );
    actionContext.telemetry?.addProperty(
      Telemetry.properties.dbCount,
      sqlConfig.databases.length.toString()
    );

    if (!skipAddingUser) {
      if (adminInfo.userType === UserType.User) {
        actionContext.progressBar?.next(ConfigureActionImplement.progressMessage.addUser);
        const sqlClient = await SqlClient.create(ctx.tokenProvider.azureAccountProvider, sqlConfig);
        actionContext.logger?.info(Message.addDatabaseUser(identity));
        await UtilFunctions.addDatabaseUser(ctx.logProvider, sqlClient, sqlMgrClient);
      } else {
        const message = ErrorMessage.ServicePrincipalWarning(
          identity,
          sqlConfig.databases.join(",")
        );
        actionContext.logger?.warning(
          `${message}. You can follow ${HelpLinks.default} to add database user ${identity}`
        );
      }
    } else {
      actionContext.logger?.warning(
        `Skip adding database user. You can follow ${HelpLinks.default} to add database user ${identity}`
      );
    }
    await sqlMgrClient.deleteLocalFirewallRule();
    return ok([{ type: "service", name: "azure", remarks: "configure azure-sql" }]);
  }

  static get(): FunctionAction {
    return {
      name: `${ComponentNames.AzureSQL}.${ActionNames.configure}`,
      type: ActionTypeFunction,
      plan: (context: ContextV3, inputs: InputsWithProjectPath) => {
        return ok([{ type: "service", name: "azure", remarks: "configure azure-sql" }]);
      },
      execute: ConfigureActionImplement.execute,
    };
  }
}

export class UtilFunctions {
  static async CheckAndSetAadAdmin(
    client: ManagementClient,
    aadAdmin: string,
    objectId: string,
    tenantId: string
  ): Promise<boolean> {
    const existAdmin = await client.existAadAdmin(aadAdmin);
    if (!existAdmin) {
      await client.addAADadmin(tenantId, objectId, aadAdmin);
    }
    return existAdmin;
  }

  static async parseLoginToken(azureAccountProvider: AzureAccountProvider): Promise<TokenInfo> {
    // get login user info to set aad admin in sql
    try {
      const credential = await azureAccountProvider.getAccountCredentialAsync();
      const token = await credential!.getToken();
      const accessToken = token.accessToken;
      return parseToken(accessToken);
    } catch (error: any) {
      throw SqlResultFactory.SystemError(
        ErrorMessage.SqlUserInfoError.name,
        ErrorMessage.SqlUserInfoError.message(),
        error
      );
    }
  }

  static async getSkipAddingUser(
    config: v3.AzureSolutionConfig,
    azureAccountProvider: AzureAccountProvider
  ): Promise<boolean> {
    const skipAddingUser = config[Constants.skipAddingSqlUser];
    if (skipAddingUser === undefined) {
      return (await azureAccountProvider?.getIdentityCredentialAsync()) ? false : true;
    } else {
      return skipAddingUser as boolean;
    }
  }

  static getIdentity(ctx: ResourceContextV3): string {
    const config = ctx.envInfo.state[ComponentNames.Identity];
    const identity = config[Constants.identityName] as string;
    if (!identity) {
      const error = SqlResultFactory.SystemError(
        ErrorMessage.SqlGetConfigError.name,
        ErrorMessage.SqlGetConfigError.message(Constants.identityPlugin, Constants.identityName)
      );
      throw error;
    }
    return identity;
  }

  static async addDatabaseUser(
    logProvider: LogProvider,
    sqlClient: SqlClient,
    managementClient: ManagementClient
  ): Promise<void> {
    let retryCount = 0;
    const databaseWithUser: { [key: string]: boolean } = {};
    sqlClient.config.databases.forEach((element) => {
      databaseWithUser[element] = false;
    });
    while (true) {
      try {
        for (const database in databaseWithUser) {
          if (!databaseWithUser[database]) {
            await sqlClient.addDatabaseUser(database);
            databaseWithUser[database] = true;
          }
        }
        return;
      } catch (error) {
        if (
          !SqlClient.isFireWallError(error?.innerError) ||
          retryCount >= Constants.maxRetryTimes
        ) {
          throw error;
        } else {
          retryCount++;
          logProvider.warning(
            `[${Constants.pluginName}] Retry adding new firewall rule to access azure sql, because the local IP address has changed after added firewall rule for it. [Retry time: ${retryCount}]`
          );
          await managementClient.addLocalFirewallRule();
        }
      }
    }
  }
}
