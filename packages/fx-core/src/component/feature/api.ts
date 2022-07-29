// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  Bicep,
  ContextV3,
  Effect,
  err,
  FxError,
  Inputs,
  InputsWithProjectPath,
  ok,
  ProvisionContextV3,
  QTreeNode,
  Result,
  Stage,
} from "@microsoft/teamsfx-api";
import { assign, cloneDeep } from "lodash";
import * as path from "path";
import "reflect-metadata";
import Container, { Service } from "typedi";
import { isVSProject } from "../../common/projectSettingsHelper";
import { convertToAlphanumericOnly } from "../../common/utils";
import { globalVars } from "../../core/globalVars";
import { CoreQuestionNames } from "../../core/question";
import {
  DefaultValues,
  FunctionPluginPathInfo,
  RegularExpr,
} from "../../plugins/resource/function/constants";
import { FunctionLanguage, QuestionKey } from "../../plugins/resource/function/enums";
import { FunctionScaffold } from "../../plugins/resource/function/ops/scaffold";
import { functionNameQuestion } from "../../plugins/resource/function/question";
import { ErrorMessages } from "../../plugins/resource/function/resources/message";
import { BicepComponent } from "../bicep";
import { ApiCodeProvider } from "../code/apiCode";
import { ComponentNames, Scenarios } from "../constants";
import { generateLocalDebugSettings } from "../debug";
import { ActionExecutionMW } from "../middleware/actionExecutionMW";
import { Plans } from "../messages";
import { AzureFunctionResource } from "../resource/azureAppService/azureFunction";
import { generateConfigBiceps, bicepUtils } from "../utils";
import { getComponent } from "../workflow";

@Service(ComponentNames.TeamsApi)
export class TeamsApi {
  name = ComponentNames.TeamsApi;
  @hooks([
    ActionExecutionMW({
      errorSource: "bot",
      errorHandler: (error) => {
        if (error && !error?.name) {
          error.name = "addBotError";
        }
        return error as FxError;
      },
      question: (context: ContextV3, inputs: InputsWithProjectPath) => {
        functionNameQuestion.validation = getFunctionNameQuestionValidation(context, inputs);
        return ok(new QTreeNode(functionNameQuestion));
      },
    }),
  ])
  async add(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<undefined, FxError>> {
    const projectSettings = context.projectSetting;
    const effects: Effect[] = [];
    inputs[CoreQuestionNames.ProgrammingLanguage] =
      context.projectSetting.programmingLanguage ||
      inputs[CoreQuestionNames.ProgrammingLanguage] ||
      "javascript";

    // 1. scaffold function
    {
      inputs[QuestionKey.functionName] =
        inputs[QuestionKey.functionName] || DefaultValues.functionName;
      const clonedInputs = cloneDeep(inputs);
      assign(clonedInputs, {
        folder: inputs.folder || FunctionPluginPathInfo.solutionFolderName,
      });
      const apiCodeComponent = Container.get<ApiCodeProvider>(ComponentNames.ApiCode);
      const res = await apiCodeComponent.generate(context, clonedInputs);
      if (res.isErr()) return err(res.error);
      effects.push("generate api code");
    }

    const apiConfig = getComponent(projectSettings, ComponentNames.TeamsApi);
    if (apiConfig) {
      apiConfig.functionNames = apiConfig.functionNames || [];
      apiConfig.functionNames.push(inputs[QuestionKey.functionName]);
      return ok(undefined);
    }

    // 2. config teams-api
    projectSettings.components.push({
      name: ComponentNames.TeamsApi,
      hosting: ComponentNames.Function,
      functionNames: [inputs[QuestionKey.functionName]],
      deploy: true,
      build: true,
      folder: inputs.folder || FunctionPluginPathInfo.solutionFolderName,
      artifactFolder: inputs.folder || FunctionPluginPathInfo.solutionFolderName,
    });
    effects.push("config teams-api");

    const biceps: Bicep[] = [];
    // 3.1 bicep.init
    {
      const bicepComponent = Container.get<BicepComponent>("bicep");
      const res = await bicepComponent.init(inputs.projectPath);
      if (res.isErr()) return err(res.error);
    }

    // 3.2 azure-function.generateBicep
    {
      const clonedInputs = cloneDeep(inputs);
      assign(clonedInputs, {
        componentId: ComponentNames.TeamsApi,
        hosting: inputs.hosting,
        scenario: Scenarios.Api,
      });
      const functionComponent = Container.get<AzureFunctionResource>(ComponentNames.Function);
      const res = await functionComponent.generateBicep(context, clonedInputs);
      if (res.isErr()) return err(res.error);
      res.value.forEach((b) => biceps.push(b));
      context.projectSetting.components.push({
        name: ComponentNames.Function,
        scenario: Scenarios.Api,
      });
    }

    const bicepRes = await bicepUtils.persistBiceps(
      inputs.projectPath,
      convertToAlphanumericOnly(context.projectSetting.appName),
      biceps
    );
    if (bicepRes.isErr()) return bicepRes;

    // 4. generate config bicep
    {
      const res = await generateConfigBiceps(context, inputs);
      if (res.isErr()) return err(res.error);
      effects.push("generate config biceps");
    }

    // 5. local debug settings
    {
      const res = await generateLocalDebugSettings(context, inputs);
      if (res.isErr()) return err(res.error);
      effects.push("generate local debug configs");
    }

    globalVars.isVS = isVSProject(projectSettings);
    projectSettings.programmingLanguage ||= inputs[CoreQuestionNames.ProgrammingLanguage];
    return ok(undefined);
  }
  async build(
    context: ProvisionContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<undefined, FxError>> {
    const apiCode = Container.get<ApiCodeProvider>(ComponentNames.ApiCode);
    const res = await apiCode.build(context, inputs);
    if (res.isErr()) return err(res.error);
    return ok(undefined);
  }
}

const getFunctionNameQuestionValidation = (context: ContextV3, inputs: InputsWithProjectPath) => ({
  validFunc: async (input: string, previousInputs?: Inputs): Promise<string | undefined> => {
    const workingPath: string = path.join(
      inputs.projectPath,
      FunctionPluginPathInfo.solutionFolderName
    );
    const name = input as string;
    if (!name || !RegularExpr.validFunctionNamePattern.test(name)) {
      return ErrorMessages.invalidFunctionName;
    }
    if (inputs.stage === Stage.create) {
      return undefined;
    }
    const language: FunctionLanguage =
      (inputs[QuestionKey.programmingLanguage] as FunctionLanguage) ??
      (context.projectSetting.programmingLanguage as FunctionLanguage);
    // If language is unknown, skip checking and let scaffold handle the error.
    if (language && (await FunctionScaffold.doesFunctionPathExist(workingPath, language, name))) {
      return ErrorMessages.functionAlreadyExists;
    }
  },
});
