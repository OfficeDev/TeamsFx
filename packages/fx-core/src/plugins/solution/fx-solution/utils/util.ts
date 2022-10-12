// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  PluginConfig,
  SolutionContext,
  PluginContext,
  Context,
  ConfigMap,
  FxError,
  TelemetryReporter,
  UserError,
  v3,
  Result,
  err,
  ok,
  ContextV3,
  ResourceContextV3,
  InputsWithProjectPath,
  Platform,
} from "@microsoft/teamsfx-api";
import { SubscriptionClient } from "@azure/arm-subscriptions";
import { SolutionTelemetryComponentName, SolutionTelemetryProperty } from "../constants";
import { TokenCredential } from "@azure/core-auth";
import { BuiltInFeaturePluginNames } from "../v3/constants";
import { ComponentNames, PathConstants } from "../../../../component/constants";
import { updateAzureParameters } from "../arm";
import { backupFiles } from "./backupFiles";
import fs from "fs-extra";
import path from "path";
import { DeployConfigsConstants } from "../../../../common/azure-hosting/hostingConstant";

export function sendErrorTelemetryThenReturnError(
  eventName: string,
  error: FxError,
  reporter?: TelemetryReporter,
  properties?: { [p: string]: string },
  measurements?: { [p: string]: number },
  errorProps?: string[]
): FxError {
  if (!properties) {
    properties = {};
  }

  if (SolutionTelemetryProperty.Component in properties === false) {
    properties[SolutionTelemetryProperty.Component] = SolutionTelemetryComponentName;
  }

  properties[SolutionTelemetryProperty.Success] = "no";
  if (error instanceof UserError) {
    properties["error-type"] = "user";
  } else {
    properties["error-type"] = "system";
  }

  properties["error-code"] = `${error.source}.${error.name}`;
  properties["error-message"] = error.message;

  reporter?.sendTelemetryErrorEvent(eventName, properties, measurements, errorProps);
  return error;
}

export function hasBotServiceCreated(envInfo: v3.EnvInfoV3): boolean {
  if (!envInfo || !envInfo.state) {
    return false;
  }

  return (
    (!!envInfo.state[BuiltInFeaturePluginNames.bot] &&
      !!envInfo.state[BuiltInFeaturePluginNames.bot]["resourceId"]) ||
    (!!envInfo.state[ComponentNames.TeamsBot] &&
      !!envInfo.state[ComponentNames.TeamsBot]["resourceId"])
  );
}

export async function handleConfigFilesWhenSwitchAccount(
  envInfo: v3.EnvInfoV3,
  context: ResourceContextV3,
  inputs: InputsWithProjectPath,
  hasSwitchedM365Tenant: boolean,
  hasSwitchedSubscription: boolean,
  hasBotServiceCreatedBefore: boolean,
  isCSharpProject: boolean
): Promise<Result<undefined, FxError>> {
  if (!hasSwitchedM365Tenant && !hasSwitchedSubscription) {
    return ok(undefined);
  }

  const backupFilesRes = await backupFiles(
    envInfo.envName,
    inputs.projectPath,
    isCSharpProject,
    inputs.platform === Platform.VS,
    context
  );
  if (backupFilesRes.isErr()) {
    return err(backupFilesRes.error);
  }

  const updateAzureParametersRes = await updateAzureParameters(
    inputs.projectPath,
    context.projectSetting.appName,
    envInfo.envName,
    hasSwitchedM365Tenant,
    hasSwitchedSubscription,
    hasBotServiceCreatedBefore
  );
  if (updateAzureParametersRes.isErr()) {
    return err(updateAzureParametersRes.error);
  }

  if (hasSwitchedSubscription) {
    const envName = envInfo.envName;
    const maybeBotFolder = path.join(inputs.projectPath, PathConstants.botWorkingDir);
    const maybeBotDeploymentFile = path.join(
      maybeBotFolder,
      path.join(
        DeployConfigsConstants.DEPLOYMENT_FOLDER,
        DeployConfigsConstants.DEPLOYMENT_INFO_FILE
      )
    );
    if (await fs.pathExists(maybeBotDeploymentFile)) {
      try {
        const botDeployJson = await fs.readJSON(maybeBotDeploymentFile);
        const lastTime = Math.max(botDeployJson[envInfo.envName]?.time ?? 0, 0);
        if (lastTime !== 0) {
          botDeployJson[envName] = {
            time: 0,
          };

          await fs.writeJSON(maybeBotDeploymentFile, botDeployJson);
        }
      } catch (exception) {
        // do nothing
      }
    }

    const maybeTabFolder = path.join(inputs.projectPath, PathConstants.tabWorkingDir);
    const maybeTabDeploymentFile = path.join(
      maybeTabFolder,
      path.join(
        DeployConfigsConstants.DEPLOYMENT_FOLDER,
        DeployConfigsConstants.DEPLOYMENT_INFO_FILE
      )
    );
    if (await fs.pathExists(maybeTabDeploymentFile)) {
      try {
        const deploymentInfoJson = await fs.readJSON(maybeTabDeploymentFile);
        if (!!deploymentInfoJson[envName] && !!deploymentInfoJson[envName].lastDeployTime) {
          delete deploymentInfoJson[envName].lastDeployTime;
          await fs.writeJSON(maybeTabDeploymentFile, deploymentInfoJson);
        }
      } catch (exception) {
        // do nothing
      }
    }
  }

  return ok(undefined);
}
