// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  ContextV3,
  err,
  FxError,
  InputsWithProjectPath,
  ok,
  ProjectSettingsV3,
  QTreeNode,
  Result,
} from "@microsoft/teamsfx-api";
import "reflect-metadata";
import Container, { Service } from "typedi";
import { isVSProject } from "../../common/projectSettingsHelper";
import { globalVars } from "../../core/globalVars";
import { CoreQuestionNames } from "../../core/question";
import {
  frameworkQuestion,
  versionCheckQuestion,
  webpartNameQuestion,
} from "../../plugins/resource/spfx/utils/questions";
import { SPFxTabCodeProvider } from "../code/spfxTabCode";
import { ComponentNames } from "../constants";
import { generateLocalDebugSettings } from "../debug";
import { ActionExecutionMW } from "../middleware/actionExecutionMW";
@Service(ComponentNames.SPFxTab)
export class SPFxTab {
  name = ComponentNames.SPFxTab;
  @hooks([
    ActionExecutionMW({
      question: (context: ContextV3, inputs: InputsWithProjectPath) => {
        const spfx_frontend_host = new QTreeNode({
          type: "group",
        });
        const spfx_version_check = new QTreeNode(versionCheckQuestion);
        spfx_frontend_host.addChild(spfx_version_check);
        const spfx_framework_type = new QTreeNode(frameworkQuestion);
        spfx_version_check.addChild(spfx_framework_type);
        const spfx_webpart_name = new QTreeNode(webpartNameQuestion);
        spfx_version_check.addChild(spfx_webpart_name);
        return ok(spfx_frontend_host);
      },
    }),
  ])
  async add(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<undefined, FxError>> {
    const projectSettings = context.projectSetting as ProjectSettingsV3;
    // add teams-tab
    projectSettings.components.push({
      name: "teams-tab",
      hosting: ComponentNames.SPFx,
      deploy: true,
      folder: inputs.folder || "SPFx",
      build: true,
    });
    // add hosting component
    projectSettings.components.push({
      name: ComponentNames.SPFx,
      provision: true,
    });
    projectSettings.programmingLanguage =
      projectSettings.programmingLanguage || inputs[CoreQuestionNames.ProgrammingLanguage];
    globalVars.isVS = isVSProject(projectSettings);
    const effects = ["config 'teams-tab' in projectSettings"];
    {
      const spfxCode = Container.get<SPFxTabCodeProvider>(ComponentNames.SPFxTabCode);
      const res = await spfxCode.generate(context, inputs);
      if (res.isErr()) return err(res.error);
      effects.push("scaffold spfx code");
    }
    {
      const res = await generateLocalDebugSettings(context, inputs);
      if (res.isErr()) return err(res.error);
      effects.push("generate local debug settings");
    }
    return ok(undefined);
  }
}
