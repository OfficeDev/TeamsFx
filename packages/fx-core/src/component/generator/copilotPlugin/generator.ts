// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yuqzho@microsoft.com
 */

import {
  ProjectType,
  SpecParser,
  SpecParserError,
  ValidationStatus,
  WarningType,
} from "@microsoft/m365-spec-parser";
import {
  AppPackageFolderName,
  Context,
  FxError,
  GeneratorResult,
  Inputs,
  ManifestTemplateFileName,
  Platform,
  Result,
  UserError,
  Warning,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import { merge } from "lodash";
import path from "path";
import { getLocalizedString } from "../../../common/localizeUtils";
import { isValidHttpUrl } from "../../../common/stringUtils";
import { assembleError } from "../../../error";
import { CapabilityOptions, ProgrammingLanguage, QuestionNames } from "../../../question/constants";
import { manifestUtils } from "../../driver/teamsApp/utils/ManifestUtils";
import { ActionContext } from "../../middleware/actionExecutionMW";
import { Generator } from "../generator";
import { DefaultTemplateGenerator } from "../templates/templateGenerator";
import { TemplateInfo } from "../templates/templateInfo";
import {
  convertSpecParserErrorToFxError,
  copilotPluginParserOptions,
  defaultApiSpecFolderName,
  defaultApiSpecJsonFileName,
  defaultApiSpecYamlFileName,
  defaultPluginManifestFileName,
  generateScaffoldingSummary,
  getEnvName,
  invalidApiSpecErrorName,
  isYamlSpecFile,
  logValidationResults,
  specParserGenerateResultAllSuccessTelemetryProperty,
  specParserGenerateResultTelemetryEvent,
  specParserGenerateResultWarningsTelemetryProperty,
  validateSpec,
} from "../../driver/teamsApp/utils/SpecUtils";

const copilotPluginExistingApiSpecUrlTelemetryEvent = "copilot-plugin-existing-api-spec-url";
const templateName = "api-plugin-existing-api";
const templateType = ProjectType.Copilot;

const enum telemetryProperties {
  templateName = "template-name",
  generateType = "generate-type",
  isRemoteUrlTelemetryProperty = "remote-url",
  authType = "auth-type",
}

function normalizePath(path: string): string {
  return "./" + path.replace(/\\/g, "/");
}

export class CopilotPluginGenerator extends DefaultTemplateGenerator {
  // activation condition
  public activate(context: Context, inputs: Inputs): boolean {
    const capability = inputs.capabilities as string;
    return capability === CapabilityOptions.copilotPluginApiSpec().id;
  }
  public async getTemplateInfos(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<TemplateInfo[], FxError>> {
    const getTemplateInfosState: any = {};
    const authData = inputs.apiAuthData;
    merge(actionContext?.telemetryProps, {
      [telemetryProperties.templateName]: templateName,
    });
    const appName = inputs[QuestionNames.AppName];
    let language = inputs[QuestionNames.ProgrammingLanguage] as ProgrammingLanguage;
    language =
      language === ProgrammingLanguage.CSharp
        ? ProgrammingLanguage.CSharp
        : ProgrammingLanguage.None;
    const safeProjectNameFromVS =
      language === "csharp" ? inputs[QuestionNames.SafeProjectName] : undefined;
    const url = inputs[QuestionNames.ApiSpecLocation];
    getTemplateInfosState.url = url.trim();

    getTemplateInfosState.isYaml = false;
    try {
      getTemplateInfosState.isYaml = await isYamlSpecFile(url);
    } catch (e) {}

    const openapiSpecFileName = getTemplateInfosState.isYaml
      ? defaultApiSpecYamlFileName
      : defaultApiSpecJsonFileName;
    const llmService: string | undefined = inputs[QuestionNames.LLMService];
    const openAIKey: string | undefined = inputs[QuestionNames.OpenAIKey];
    const azureOpenAIKey: string | undefined = inputs[QuestionNames.AzureOpenAIKey];
    const azureOpenAIEndpoint: string | undefined = inputs[QuestionNames.AzureOpenAIEndpoint];
    const azureOpenAIDeploymentName: string | undefined =
      inputs[QuestionNames.AzureOpenAIDeploymentName];
    const llmServiceData = {
      llmService,
      openAIKey,
      azureOpenAIKey,
      azureOpenAIEndpoint,
      azureOpenAIDeploymentName,
    };
    if (authData?.authName) {
      const envName = getEnvName(authData.authName, authData.authType);
      context.templateVariables = Generator.getDefaultVariables(
        appName,
        safeProjectNameFromVS,
        inputs.targetFramework,
        inputs.placeProjectFileInSolutionDir === "true",
        {
          authName: authData.authName,
          openapiSpecPath: normalizePath(
            path.join(AppPackageFolderName, defaultApiSpecFolderName, openapiSpecFileName)
          ),
          registrationIdEnvName: envName,
          authType: authData.authType,
        },
        llmServiceData
      );
    } else {
      context.templateVariables = Generator.getDefaultVariables(
        appName,
        safeProjectNameFromVS,
        inputs.targetFramework,
        inputs.placeProjectFileInSolutionDir === "true",
        undefined,
        llmServiceData
      );
    }
    context.telemetryReporter.sendTelemetryEvent(copilotPluginExistingApiSpecUrlTelemetryEvent, {
      [telemetryProperties.isRemoteUrlTelemetryProperty]: isValidHttpUrl(url).toString(),
      [telemetryProperties.generateType]: templateType.toString(),
      [telemetryProperties.authType]: authData?.authName ?? "None",
    });
    inputs.getTemplateInfosState = getTemplateInfosState;
    return ok([
      {
        templateName: templateName,
        language: language,
        replaceMap: context.templateVariables,
      },
    ]);
  }

  public async post(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<GeneratorResult, FxError>> {
    try {
      const componentName = "copilot-generator";
      const getTemplateInfosState = inputs.getTemplateInfosState;
      // validate API spec
      const specParser = new SpecParser(getTemplateInfosState.url, copilotPluginParserOptions);
      const filters = inputs[QuestionNames.ApiOperation] as string[];
      const validationRes = await validateSpec(specParser, filters);
      const warnings = validationRes.warnings;
      if (validationRes.status === ValidationStatus.Error) {
        logValidationResults(validationRes.errors, warnings, context, false, true);
        const errorMessage =
          inputs.platform === Platform.VSCode
            ? getLocalizedString(
                "core.createProjectQuestion.apiSpec.multipleValidationErrors.vscode.message"
              )
            : getLocalizedString(
                "core.createProjectQuestion.apiSpec.multipleValidationErrors.message"
              );
        return err(
          new UserError(componentName, invalidApiSpecErrorName, errorMessage, errorMessage)
        );
      }
      const manifestPath = path.join(
        destinationPath,
        AppPackageFolderName,
        ManifestTemplateFileName
      );
      const apiSpecFolderPath = path.join(
        destinationPath,
        AppPackageFolderName,
        defaultApiSpecFolderName
      );
      const openapiSpecFileName = getTemplateInfosState.isYaml
        ? defaultApiSpecYamlFileName
        : defaultApiSpecJsonFileName;
      const openapiSpecPath = path.join(apiSpecFolderPath, openapiSpecFileName);
      // generate files
      await fs.ensureDir(apiSpecFolderPath);

      const pluginManifestPath = path.join(
        destinationPath,
        AppPackageFolderName,
        defaultPluginManifestFileName
      );
      const generateResult = await specParser.generateForCopilot(
        manifestPath,
        filters,
        openapiSpecPath,
        pluginManifestPath
      );

      context.telemetryReporter.sendTelemetryEvent(specParserGenerateResultTelemetryEvent, {
        [telemetryProperties.generateType]: templateType.toString(),
        [specParserGenerateResultAllSuccessTelemetryProperty]: generateResult.allSuccess.toString(),
        [specParserGenerateResultWarningsTelemetryProperty]: generateResult.warnings
          .map((w) => w.type.toString() + ": " + w.content)
          .join(";"),
      });

      if (generateResult.warnings.length > 0) {
        generateResult.warnings.find((o) => {
          if (o.type === WarningType.OperationOnlyContainsOptionalParam) {
            o.content = ""; // We don't care content of this warning
          }
        });
        warnings.push(...generateResult.warnings);
      }

      // update manifest based on openAI plugin manifest
      const manifestRes = await manifestUtils._readAppManifest(manifestPath);

      if (manifestRes.isErr()) {
        return err(manifestRes.error);
      }

      const teamsManifest = manifestRes.value;

      // log warnings
      if (inputs.platform === Platform.CLI || inputs.platform === Platform.VS) {
        const warnSummary = generateScaffoldingSummary(
          warnings,
          teamsManifest,
          path.relative(destinationPath, openapiSpecPath)
        );

        if (warnSummary) {
          void context.logProvider.info(warnSummary);
        }
      }

      if (inputs.platform === Platform.VSCode) {
        return ok({
          warnings: warnings.map((warning) => {
            return {
              type: warning.type,
              content: warning.content,
              data: warning.data,
            };
          }),
        });
      } else {
        return ok({ warnings: undefined });
      }
    } catch (e) {
      let error: FxError;
      if (e instanceof SpecParserError) {
        error = convertSpecParserErrorToFxError(e);
      } else {
        error = assembleError(e);
      }
      return err(error);
    }
  }
}
