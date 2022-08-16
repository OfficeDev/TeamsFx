// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ContextV3,
  err,
  FxError,
  InputsWithProjectPath,
  ok,
  Platform,
  ProjectSettingsV3,
  QTreeNode,
  Result,
} from "@microsoft/teamsfx-api";
import "reflect-metadata";
import { Container, Service } from "typedi";
import { format } from "util";
import { getLocalizedString } from "../../common/localizeUtils";
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
@Service(ComponentNames.SPFxTab)
export class SPFxTab {
  name = ComponentNames.SPFxTab;

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
    globalVars.isVS = projectSettings.programmingLanguage === "csharp";
    {
      const spfxCode = Container.get<SPFxTabCodeProvider>(ComponentNames.SPFxTabCode);
      const res = await spfxCode.generate(context, inputs);
      if (res.isErr()) return err(res.error);
    }
    {
      const res = await generateLocalDebugSettings(context, inputs);
      if (res.isErr()) return err(res.error);
    }
    // notification
    const msg =
      inputs.platform === Platform.CLI
        ? getLocalizedString("core.addCapability.addCapabilityNoticeForCli")
        : getLocalizedString("core.addCapability.addCapabilitiesNotice");
    context.userInteraction.showMessage(
      "info",
      format(msg, inputs[CoreQuestionNames.Features]),
      false
    );
    return ok(undefined);
  }
}

export function getSPFxScaffoldQuestion(): QTreeNode {
  const spfx_frontend_host = new QTreeNode({
    type: "group",
  });
  const spfx_version_check = new QTreeNode(versionCheckQuestion);
  spfx_frontend_host.addChild(spfx_version_check);
  const spfx_framework_type = new QTreeNode(frameworkQuestion);
  spfx_version_check.addChild(spfx_framework_type);
  const spfx_webpart_name = new QTreeNode(webpartNameQuestion);
  spfx_version_check.addChild(spfx_webpart_name);
  return spfx_frontend_host;
}
