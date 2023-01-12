// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ExecutionResult, StepDriver } from "../../interface/stepDriver";
import { Service } from "typedi";
import { DriverContext } from "../../interface/commonArgs";
import { FxError, Result } from "@microsoft/teamsfx-api";
import { hooks } from "@feathersjs/hooks/lib";
import { addStartAndEndTelemetry } from "../../middleware/addStartAndEndTelemetry";
import { TelemetryConstant } from "../../../constant/commonConstant";
import { getLocalizedString } from "../../../../common/localizeUtils";
import { AzureZipDeployDriver } from "./AzureZipDeployDriver";

const ACTION_NAME = "azureAppService/deploy";

@Service(ACTION_NAME)
export class AzureAppServiceDeployDriver implements StepDriver {
  readonly description: string = getLocalizedString(
    "driver.deploy.deployToAzureAppServiceDescription"
  );
  private static readonly SERVICE_NAME = "Azure App Service";
  private static readonly SUMMARY = [
    getLocalizedString("driver.deploy.azureAppServiceDeploySummary"),
  ];
  private static readonly SUMMARY_PREPARE = [
    getLocalizedString("driver.deploy.notice.deployDryRunComplete"),
  ];

  @hooks([addStartAndEndTelemetry(ACTION_NAME, TelemetryConstant.DEPLOY_COMPONENT_NAME)])
  async run(args: unknown, context: DriverContext): Promise<Result<Map<string, string>, FxError>> {
    const impl = new AzureZipDeployDriver(
      args,
      context,
      AzureAppServiceDeployDriver.SERVICE_NAME,
      AzureAppServiceDeployDriver.SUMMARY,
      AzureAppServiceDeployDriver.SUMMARY_PREPARE
    );
    return (await impl.run()).result;
  }

  @hooks([addStartAndEndTelemetry(ACTION_NAME, TelemetryConstant.DEPLOY_COMPONENT_NAME)])
  async execute(args: unknown, ctx: DriverContext): Promise<ExecutionResult> {
    const impl = new AzureZipDeployDriver(
      args,
      ctx,
      AzureAppServiceDeployDriver.SERVICE_NAME,
      AzureAppServiceDeployDriver.SUMMARY,
      AzureAppServiceDeployDriver.SUMMARY_PREPARE
    );
    return await impl.run();
  }
}
