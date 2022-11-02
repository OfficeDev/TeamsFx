// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Constants, TelemetryProperties, TemplateType } from "./constant";
import { deployArgs, deploymentOutput, templateArgs } from "./interface";
import { validateArgs } from "./validator";
import { hasBicepTemplate, convertOutputs, getFileExtension } from "./util/util";
import { err, FxError, ok, Result, SystemError, UserError } from "@microsoft/teamsfx-api";
import { ConstantString, PluginDisplayName } from "../../../common/constants";
import * as fs from "fs-extra";
import { expandEnvironmentVariable, getAbsolutePath } from "../../utils/common";
import { executeCommand } from "../../../common/cpUtils";
import { getDefaultString, getLocalizedString } from "../../../common/localizeUtils";
import { Deployment, DeploymentMode, ResourceManagementClient } from "@azure/arm-resources";
import { SolutionError } from "../../constants";
import { InvalidParameterUserError } from "../aad/error/invalidParameterUserError";
import { ensureBicepForDriver } from "../../utils/depsChecker/bicepChecker";
import { WrapDriverContext } from "../util/wrapUtil";

const helpLink = "https://aka.ms/teamsfx-actions/arm-deploy";

export class ArmDeployImpl {
  args: deployArgs;
  context: WrapDriverContext;
  bicepCommand?: string;
  client?: ResourceManagementClient;

  constructor(args: deployArgs, context: WrapDriverContext) {
    this.args = args;
    this.context = context;
  }

  public async run(): Promise<Map<string, string>> {
    await this.validateArgs();
    await this.createClient();
    const needBicepCli = hasBicepTemplate(this.args.templates);
    if (needBicepCli) {
      this.bicepCommand = await this.ensureBicepCli();
    }
    const deployRes = await this.deployTemplates();
    if (deployRes.isOk()) {
      const outputs = deployRes.value;
      return convertOutputs(outputs);
    } else {
      throw deployRes.error;
    }
  }

  private async validateArgs(): Promise<void> {
    const invalidParameters = await validateArgs(this.args);

    if (invalidParameters.length > 0) {
      throw new InvalidParameterUserError(Constants.actionName, invalidParameters, helpLink);
    }
  }

  public async ensureBicepCli(): Promise<string> {
    return await ensureBicepForDriver(this.context, this.args.bicepCliVersion!);
  }

  private async createClient(): Promise<void> {
    const azureToken = await this.context.azureAccountProvider.getIdentityCredentialAsync();
    if (!azureToken) {
      throw new SystemError(
        PluginDisplayName.Solution,
        SolutionError.FailedToGetAzureCredential,
        getDefaultString("core.deployArmTemplates.InvalidAzureCredential"),
        getLocalizedString("core.deployArmTemplates.InvalidAzureCredential")
      );
    }
    this.client = new ResourceManagementClient(azureToken, this.args.subscriptionId);
  }

  async deployTemplates(): Promise<Result<deploymentOutput[], FxError>> {
    const outputs: deploymentOutput[] = [];
    this.setTelemetries();
    await Promise.all(
      this.args.templates.map(async (template) => {
        const res = await this.deployTemplate(template);
        if (res.isOk() && res.value) {
          outputs.push(res.value);
        } else if (res.isErr()) {
          throw res.error;
        }
      })
    );
    return ok(outputs);
  }

  async deployTemplate(
    templateArg: templateArgs
  ): Promise<Result<deploymentOutput | undefined, FxError>> {
    try {
      const progressBar = this.context.createProgressBar(
        `Deploy arm: ${templateArg.deploymentName}`,
        1
      );
      const parameters = await this.getDeployParameters(templateArg.parameters);
      const template = await this.getDeployTemplate(templateArg.path);
      const deploymentParameters: Deployment = {
        properties: {
          parameters: parameters.parameters,
          template: template as any,
          mode: "Incremental" as DeploymentMode,
        },
      };
      const res = await this.executeDeployment(templateArg, deploymentParameters);
      progressBar?.end(true);
      return res;
    } catch (error) {
      const fxError = new UserError(
        Constants.actionName,
        "FailedToDeployArmTemplate",
        getDefaultString("driver.arm.error.deploy", templateArg.deploymentName, error.message),
        getLocalizedString("driver.arm.error.deploy", templateArg.deploymentName, error.message)
      );
      return err(fxError);
    }
  }

  async executeDeployment(
    templateArg: templateArgs,
    deploymentParameters: Deployment
  ): Promise<Result<deploymentOutput | undefined, FxError>> {
    try {
      const result = await this.client?.deployments.beginCreateOrUpdateAndWait(
        this.args.resourceGroupName,
        templateArg.deploymentName,
        deploymentParameters
      );
      return ok(result?.properties?.outputs);
    } catch (error) {
      // TODO: fetch error reason and throw
      throw error;
    }
  }

  private async getDeployParameters(parameters: string): Promise<any> {
    const filePath = getAbsolutePath(parameters, this.context.projectPath);
    const template = await fs.readFile(filePath, ConstantString.UTF8Encoding);
    const parameterJsonString = expandEnvironmentVariable(template);
    return JSON.parse(parameterJsonString);
  }

  private async getDeployTemplate(templatePath: string): Promise<string> {
    const templateType = getFileExtension(templatePath);
    const filePath = getAbsolutePath(templatePath, this.context.projectPath);
    let templateJsonString;
    if (templateType === TemplateType.Bicep) {
      templateJsonString = await this.compileBicepToJson(filePath);
    } else {
      const template = await fs.readFile(filePath, ConstantString.UTF8Encoding);
      templateJsonString = JSON.parse(template);
    }
    return templateJsonString;
  }

  async compileBicepToJson(filePath: string): Promise<JSON> {
    try {
      const result = await executeCommand(
        this.bicepCommand!,
        ["build", filePath, "--stdout"],
        this.context.logProvider,
        { shell: false }
      );
      return JSON.parse(result);
    } catch (err) {
      throw new Error(getDefaultString("driver.arm.error.CompileBicepFailed", err.message));
    }
  }

  private setTelemetries(): void {
    let bicepCount = 0;
    let jsonCount = 0;
    for (const template of this.args.templates) {
      const templateType = getFileExtension(template.path);
      if (templateType === TemplateType.Bicep) {
        bicepCount++;
      } else {
        jsonCount++;
      }
    }
    this.context.addTelemetryProperties({
      [TelemetryProperties.bicepTemplateCount]: bicepCount.toString(),
      [TelemetryProperties.jsonTemplateCount]: jsonCount.toString(),
    });
  }
}
