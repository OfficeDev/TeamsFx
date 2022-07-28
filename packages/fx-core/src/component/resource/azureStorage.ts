// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ActionContext,
  AzureAccountProvider,
  err,
  FxError,
  InputsWithProjectPath,
  ok,
  ResourceContextV3,
  Result,
  UserError,
  v3,
} from "@microsoft/teamsfx-api";
import "reflect-metadata";
import { Service } from "typedi";
import * as path from "path";
import { FrontendConfig } from "../../plugins/resource/frontend/configs";
import {
  getResourceGroupNameFromResourceId,
  getStorageAccountNameFromResourceId,
  getSubscriptionIdFromResourceId,
} from "../../common/tools";
import { UnauthenticatedError } from "../../plugins/resource/frontend/v3/error";
import { AzureStorageClient } from "../../plugins/resource/frontend/clients";
import { FrontendDeployment } from "../../plugins/resource/frontend/ops/deploy";
import { AzureResource } from "./azureResource";
import { FrontendPluginInfo } from "../../plugins/resource/frontend/constants";
import { ComponentNames, StorageOutputs } from "../constants";
import { ProgressMessages, ProgressTitles } from "../messages";
import { hooks } from "@feathersjs/hooks/lib";
import { ActionExecutionMW } from "../middleware/actionExecutionMW";
@Service("azure-storage")
export class AzureStorageResource extends AzureResource {
  readonly name = "azure-storage";
  readonly bicepModuleName = "azureStorage";
  readonly outputs = StorageOutputs;
  readonly finalOutputKeys = ["domain", "endpoint", "storageResourceId", "indexPath"];
  @hooks([
    ActionExecutionMW({
      enableTelemetry: true,
      telemetryComponentName: FrontendPluginInfo.PluginName,
      telemetryEventName: "deploy",
      errorSource: FrontendPluginInfo.ShortName,
      errorIssueLink: FrontendPluginInfo.IssueLink,
      errorHelpLink: FrontendPluginInfo.HelpLink,
      enableProgressBar: true,
      progressTitle: ProgressTitles.configureStorage,
      progressSteps: 1,
    }),
  ])
  async configure(
    context: ResourceContextV3,
    inputs: InputsWithProjectPath,
    actionContext?: ActionContext
  ): Promise<Result<undefined, FxError>> {
    const ctx = context as ResourceContextV3;
    if (context.envInfo.envName !== "local") {
      const frontendConfigRes = await buildFrontendConfig(
        ctx.envInfo,
        ComponentNames.TeamsTab,
        ctx.tokenProvider.azureAccountProvider
      );
      if (frontendConfigRes.isErr()) {
        return err(frontendConfigRes.error);
      }
      actionContext?.progressBar?.next(ProgressMessages.enableStaticWebsite);
      const client = new AzureStorageClient(frontendConfigRes.value);
      await client.enableStaticWebsite();
    } else {
      actionContext?.progressBar?.next(ProgressMessages.enableStaticWebsite);
    }
    return ok(undefined);
  }
  @hooks([
    ActionExecutionMW({
      enableTelemetry: true,
      telemetryComponentName: FrontendPluginInfo.PluginName,
      telemetryEventName: "deploy",
      errorSource: FrontendPluginInfo.ShortName,
      errorIssueLink: FrontendPluginInfo.IssueLink,
      errorHelpLink: FrontendPluginInfo.HelpLink,
      enableProgressBar: true,
      progressTitle: ProgressTitles.deployingStorage,
      progressSteps: 3,
    }),
  ])
  async deploy(
    context: ResourceContextV3,
    inputs: InputsWithProjectPath,
    actionContext?: ActionContext
  ): Promise<Result<undefined, FxError>> {
    const ctx = context as ResourceContextV3;
    const deployDir = path.resolve(inputs.projectPath, inputs.folder);
    const frontendConfigRes = await buildFrontendConfig(
      ctx.envInfo,
      inputs.componentId,
      ctx.tokenProvider.azureAccountProvider
    );
    if (frontendConfigRes.isErr()) {
      return err(frontendConfigRes.error);
    }
    const client = new AzureStorageClient(frontendConfigRes.value);
    const envName = ctx.envInfo.envName;
    await FrontendDeployment.doFrontendDeploymentV3(
      client,
      deployDir,
      envName,
      actionContext?.progressBar
    );
    return ok(undefined);
  }
}

async function buildFrontendConfig(
  envInfo: v3.EnvInfoV3,
  scenario: string,
  tokenProvider: AzureAccountProvider
): Promise<Result<FrontendConfig, FxError>> {
  const credentials = await tokenProvider.getAccountCredentialAsync();
  if (!credentials) {
    return err(new UnauthenticatedError());
  }
  const storage = envInfo.state[scenario];
  const resourceId = storage?.storageResourceId;
  if (!resourceId) {
    return err(
      new UserError({
        source: "azure-storage",
        name: "StateValueMissingError",
        message: "Missing resourceId for storage",
      })
    );
  }
  const frontendConfig = new FrontendConfig(
    getSubscriptionIdFromResourceId(resourceId),
    getResourceGroupNameFromResourceId(resourceId),
    (envInfo.state.solution as v3.AzureSolutionConfig).location,
    getStorageAccountNameFromResourceId(resourceId),
    credentials
  );
  return ok(frontendConfig);
}
