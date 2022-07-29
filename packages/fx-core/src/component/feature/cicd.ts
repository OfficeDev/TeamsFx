// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ContextV3,
  err,
  FxError,
  Inputs,
  InputsWithProjectPath,
  MultiSelectQuestion,
  ok,
  OptionItem,
  Platform,
  QTreeNode,
  Result,
  SingleSelectQuestion,
  Stage,
} from "@microsoft/teamsfx-api";
import "reflect-metadata";
import { Service } from "typedi";
import { getDefaultString, getLocalizedString } from "../../common/localizeUtils";
import { NoCapabilityFoundError } from "../../core/error";
import { environmentManager } from "../../core/environment";
import { InternalError, NoProjectOpenedError } from "../../plugins/resource/cicd/errors";
import {
  azdoOption,
  cdOption,
  ciOption,
  githubOption,
  jenkinsOption,
  provisionOption,
  publishOption,
  questionNames,
} from "../../plugins/resource/cicd/questions";
import { ExistingTemplatesStat } from "../../plugins/resource/cicd/utils/existingTemplatesStat";
import { isMiniApp } from "../../common/projectSettingsHelperV3";
import { CICDImpl } from "../../plugins/resource/cicd/plugin";
import { isV3 } from "../../core/globalVars";
import { isExistingTabApp } from "../../common/projectSettingsHelper";
import { ComponentNames } from "../constants";
import { hooks } from "@feathersjs/hooks/lib";
import { ActionExecutionMW } from "../middleware/actionExecutionMW";
@Service(ComponentNames.CICD)
export class CICD {
  name = ComponentNames.CICD;
  @hooks([
    ActionExecutionMW({
      question: async (context: ContextV3, inputs: InputsWithProjectPath) => {
        return await addCicdQuestion(context, inputs);
      },
    }),
  ])
  async add(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<undefined, FxError>> {
    const cicdImpl: CICDImpl = new CICDImpl();
    const envName = inputs.env || inputs[questionNames.Environment];
    const res = await cicdImpl.addCICDWorkflows(context, inputs, envName);
    if (res.isErr()) return err(res.error);
    return ok(undefined);
  }
}

export async function addCicdQuestion(
  ctx: ContextV3,
  inputs: InputsWithProjectPath
): Promise<Result<QTreeNode | undefined, FxError>> {
  // add CI CD workflows for minimal app is not supported.
  const isExistingApp = !isV3()
    ? isExistingTabApp(ctx.projectSetting)
    : isMiniApp(ctx.projectSetting);
  if (inputs.platform !== Platform.CLI_HELP && isExistingApp) {
    throw new NoCapabilityFoundError(Stage.addCiCdFlow);
  }

  const cicdWorkflowQuestions = new QTreeNode({
    type: "group",
  });

  const whichProvider: SingleSelectQuestion = {
    name: questionNames.Provider,
    type: "singleSelect",
    staticOptions: [githubOption, azdoOption, jenkinsOption],
    title: getLocalizedString("plugins.cicd.whichProvider.title"),
    default: githubOption.id,
  };

  const whichTemplate: MultiSelectQuestion = {
    name: questionNames.Template,
    type: "multiSelect",
    staticOptions: [ciOption, cdOption, provisionOption, publishOption],
    title: getLocalizedString("plugins.cicd.whichTemplate.title"),
    default: [ciOption.id],
  };

  // TODO: add support for VS/.Net Projects.
  if (inputs.platform === Platform.VSCode) {
    if (!inputs.projectPath) {
      throw new NoProjectOpenedError();
    }

    const envProfilesResult = await environmentManager.listRemoteEnvConfigs(inputs.projectPath);
    if (envProfilesResult.isErr()) {
      throw new InternalError(
        [
          getDefaultString("error.cicd.FailedToListMultiEnv", envProfilesResult.error.message),
          getLocalizedString("error.cicd.FailedToListMultiEnv", envProfilesResult.error.message),
        ],
        envProfilesResult.error
      );
    }

    const existingInstance = ExistingTemplatesStat.getInstance(
      inputs.projectPath!,
      envProfilesResult.value
    );
    // Mute this scan before there's initial scan on upper layers.
    // await existingInstance.scan();

    const whichEnvironment: SingleSelectQuestion = {
      type: "singleSelect",
      name: questionNames.Environment,
      title: getLocalizedString("plugins.cicd.whichEnvironment.title"),
      staticOptions: [],
      dynamicOptions: async (inputs: Inputs): Promise<OptionItem[]> => {
        // Remove the env items in which all combinations of templates are scaffolded/existing.
        return existingInstance.availableEnvOptions();
      },
      skipSingleOption: true,
    };

    whichProvider.dynamicOptions = async (inputs: Inputs): Promise<OptionItem[]> => {
      const envName = inputs[questionNames.Environment];
      return existingInstance.availableProviderOptions(envName);
    };

    whichTemplate.dynamicOptions = async (inputs: Inputs): Promise<OptionItem[]> => {
      const envName = inputs[questionNames.Environment];
      const provider = inputs[questionNames.Provider];
      return existingInstance.availableTemplateOptions(envName, provider);
    };

    cicdWorkflowQuestions.addChild(new QTreeNode(whichEnvironment));
  }

  cicdWorkflowQuestions.addChild(new QTreeNode(whichProvider));
  cicdWorkflowQuestions.addChild(new QTreeNode(whichTemplate));

  return ok(cicdWorkflowQuestions);
}
