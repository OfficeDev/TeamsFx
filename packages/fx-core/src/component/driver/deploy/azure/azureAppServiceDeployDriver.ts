// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { DeployStepArgs } from "../../interface/buildAndDeployArgs";
import { AzureDeployDriver } from "./azureDeployDriver";
import { StepDriver } from "../../interface/stepDriver";
import { Service } from "typedi";
import { DriverContext, AzureResourceInfo } from "../../interface/commonArgs";
import { TokenCredential } from "@azure/identity";
import { FxError, Result } from "@microsoft/teamsfx-api";
import { wrapRun } from "../../../utils/common";
import { hooks } from "@feathersjs/hooks/lib";
import { addStartAndEndTelemetry } from "../../middleware/addStartAndEndTelemetry";
import { TelemetryConstant } from "../../../constant/commonConstant";
import { DeployConstant } from "../../../constant/deployConstant";

const ACTION_NAME = "azureAppService/deploy";

@Service(ACTION_NAME)
export class AzureAppServiceDeployDriver implements StepDriver {
  @hooks([addStartAndEndTelemetry(ACTION_NAME, TelemetryConstant.DEPLOY_COMPONENT_NAME)])
  async run(args: unknown, context: DriverContext): Promise<Result<Map<string, string>, FxError>> {
    const impl = new AzureAppServiceDeployDriverImpl(args, context);
    return wrapRun(
      () => impl.run(),
      () => impl.cleanup()
    );
  }
}

export class AzureAppServiceDeployDriverImpl extends AzureDeployDriver {
  progressBarName = `Deploying ${this.workingDirectory ?? ""} to Azure App Service`;
  progressBarSteps = 5;
  pattern =
    /\/subscriptions\/([^\/]*)\/resourceGroups\/([^\/]*)\/providers\/Microsoft.Web\/sites\/([^\/]*)/i;

  async azureDeploy(
    args: DeployStepArgs,
    azureResource: AzureResourceInfo,
    azureCredential: TokenCredential
  ): Promise<void> {
    const startTime = Date.now();
    await this.progressBar?.start();
    await this.zipDeploy(args, azureResource, azureCredential);
    await this.progressBar?.end(true);
    if (startTime + DeployConstant.DEPLOY_OVER_TIME < Date.now()) {
      await this.context.logProvider?.info(
        `Deploying to Azure App Service takes a long time. Consider referring to this document to optimize your deployment: 
        https://learn.microsoft.com/en-us/azure/app-service/deploy-run-package`
      );
    }
  }
}
