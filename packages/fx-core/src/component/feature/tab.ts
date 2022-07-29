// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Bicep,
  CloudResource,
  ContextV3,
  err,
  FxError,
  InputsWithProjectPath,
  ok,
  Platform,
  ProvisionContextV3,
  Result,
  Stage,
  v3,
} from "@microsoft/teamsfx-api";
import "reflect-metadata";
import { Container, Service } from "typedi";
import { format } from "util";
import { getLocalizedString } from "../../common/localizeUtils";
import { isVSProject } from "../../common/projectSettingsHelper";
import { globalVars } from "../../core/globalVars";
import { CoreQuestionNames } from "../../core/question";
import { Constants, FrontendPathInfo } from "../../plugins/resource/frontend/constants";
import {
  AzureSolutionQuestionNames,
  TabNonSsoItem,
} from "../../plugins/solution/fx-solution/question";
import { ComponentNames, Scenarios } from "../constants";
import { Plans } from "../messages";
import { getComponent, getComponentByScenario } from "../workflow";
import { assign, cloneDeep } from "lodash";
import { hasTab } from "../../common/projectSettingsHelperV3";
import { generateConfigBiceps, bicepUtils } from "../utils";
import { TabCodeProvider } from "../code/tabCode";
import { BicepComponent } from "../bicep";
import { convertToAlphanumericOnly } from "../../common/utils";
import { IdentityResource } from "../resource/identity";
import { generateLocalDebugSettings } from "../debug";
import { AppManifest } from "../resource/appManifest/appManifest";
import { FRONTEND_INDEX_PATH } from "../../plugins/resource/appstudio/constants";
import { ActionExecutionMW } from "../middleware/actionExecutionMW";
import { hooks } from "@feathersjs/hooks/lib";

@Service("teams-tab")
export class TeamsTab {
  name = "teams-tab";
  async add(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<undefined, FxError>> {
    const projectSettings = context.projectSetting;
    const effects = [];
    inputs[CoreQuestionNames.ProgrammingLanguage] =
      context.projectSetting.programmingLanguage ||
      inputs[CoreQuestionNames.ProgrammingLanguage] ||
      "javascript";
    inputs.hosting ||=
      inputs[CoreQuestionNames.ProgrammingLanguage] === "csharp"
        ? ComponentNames.AzureWebApp
        : ComponentNames.AzureStorage;
    // scaffold and config tab
    let tabConfig = getComponent(projectSettings, ComponentNames.TeamsTab);
    if (!tabConfig) {
      const clonedInputs = cloneDeep(inputs);
      clonedInputs.folder ||=
        inputs[CoreQuestionNames.ProgrammingLanguage] === "csharp"
          ? ""
          : FrontendPathInfo.WorkingDir;
      clonedInputs.language = inputs[CoreQuestionNames.ProgrammingLanguage];
      const tabCode = Container.get(ComponentNames.TabCode) as TabCodeProvider;
      const res = await tabCode.generate(context, clonedInputs);
      if (res.isErr()) return err(res.error);
      effects.push("generate tab code");
      tabConfig = {
        name: ComponentNames.TeamsTab,
        hosting: inputs.hosting,
        deploy: true,
        provision: inputs[CoreQuestionNames.ProgrammingLanguage] != "csharp",
        build: true,
        folder: clonedInputs.folder,
      };
      projectSettings.components.push(tabConfig);
      effects.push(Plans.generateSourceCodeAndConfig(ComponentNames.TeamsTab));

      // 2. generate provision bicep
      // 2.0 bicep.init
      {
        const bicepComponent = Container.get<BicepComponent>("bicep");
        const res = await bicepComponent.init(inputs.projectPath);
        if (res.isErr()) return err(res.error);
      }
      const biceps: Bicep[] = [];
      // 2.1 hosting bicep
      const hostingConfig = getComponentByScenario(projectSettings, inputs.hosting, Scenarios.Tab);
      if (!hostingConfig) {
        const clonedInputs = cloneDeep(inputs);
        assign(clonedInputs, {
          componentId: ComponentNames.TeamsTab,
          scenario: Scenarios.Tab,
        });
        const hostingComponent = Container.get<CloudResource>(inputs.hosting);
        const res = await hostingComponent.generateBicep!(context, clonedInputs);
        if (res.isErr()) return err(res.error);
        res.value.forEach((b) => biceps.push(b));
        projectSettings.components.push({
          name: inputs.hosting,
          scenario: Scenarios.Tab,
          provision: true,
        });
        effects.push(Plans.generateBicepAndConfig(inputs.hosting));
      }

      // 2.2 identity bicep
      if (!getComponent(projectSettings, ComponentNames.Identity)) {
        const clonedInputs = cloneDeep(inputs);
        assign(clonedInputs, {
          componentId: "",
          scenario: "",
        });
        const identityComponent = Container.get<IdentityResource>(ComponentNames.Identity);
        const res = await identityComponent.generateBicep(context, clonedInputs);
        if (res.isErr()) return err(res.error);
        res.value.forEach((b) => biceps.push(b));
        projectSettings.components.push({
          name: ComponentNames.Identity,
          provision: true,
        });
        effects.push(Plans.generateBicepAndConfig(ComponentNames.Identity));
      }

      //persist bicep
      const bicepRes = await bicepUtils.persistBiceps(
        inputs.projectPath,
        convertToAlphanumericOnly(context.projectSetting.appName),
        biceps
      );
      if (bicepRes.isErr()) return bicepRes;
      // 2.3 add sso
      if (
        inputs.stage === Stage.create &&
        inputs[AzureSolutionQuestionNames.Features] !== TabNonSsoItem.id
      ) {
        const ssoComponent = Container.get("sso") as any;
        const res = await ssoComponent.add(context, inputs);
        if (res.isErr()) return err(res.error);
      }

      // 3. generate config bicep
      {
        const res = await generateConfigBiceps(context, inputs);
        if (res.isErr()) return err(res.error);
        effects.push("generate config biceps");
      }

      // 4. local debug settings
      {
        const res = await generateLocalDebugSettings(context, inputs);
        if (res.isErr()) return err(res.error);
        effects.push("generate local debug configs");
      }
    }

    // 5. app-manifest.addCapability
    {
      const capabilities: v3.ManifestCapability[] = [{ name: "staticTab" }];
      if (!hasTab(projectSettings)) {
        capabilities.push({ name: "configurableTab" });
      }
      const clonedInputs = cloneDeep(inputs);
      const manifestComponent = Container.get<AppManifest>(ComponentNames.AppManifest);
      const res = await manifestComponent.addCapability(clonedInputs, capabilities);
      if (res.isErr()) return err(res.error);
      effects.push("add tab capability in app manifest");
    }
    globalVars.isVS = isVSProject(projectSettings);
    projectSettings.programmingLanguage ||= inputs[CoreQuestionNames.ProgrammingLanguage];
    const msg =
      inputs.platform === Platform.CLI
        ? getLocalizedString("core.addCapability.addCapabilityNoticeForCli")
        : getLocalizedString("core.addCapability.addCapabilitiesNoticeForCli");
    context.userInteraction.showMessage("info", format(msg, "Tab"), false);
    return ok(undefined);
  }

  @hooks([
    ActionExecutionMW({
      errorSource: "FE",
    }),
  ])
  async provision(
    context: ProvisionContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<undefined, FxError>> {
    if (context.envInfo.envName === "local") {
      context.envInfo.state[ComponentNames.TeamsTab] =
        context.envInfo.state[ComponentNames.TeamsTab] || {};
      context.envInfo.state[ComponentNames.TeamsTab][FRONTEND_INDEX_PATH] =
        Constants.FrontendIndexPath;
    }
    return ok(undefined);
  }
  @hooks([
    ActionExecutionMW({
      errorSource: "FE",
    }),
  ])
  async configure(
    context: ProvisionContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<undefined, FxError>> {
    const tabCode = new TabCodeProvider();
    const res = await tabCode.configure(context as ProvisionContextV3, inputs);
    if (res.isErr()) return err(res.error);
    return ok(undefined);
  }
  @hooks([
    ActionExecutionMW({
      errorSource: "FE",
    }),
  ])
  async build(
    context: ProvisionContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<undefined, FxError>> {
    const tabCode = new TabCodeProvider();
    const res = await tabCode.build(context, inputs);
    if (res.isErr()) return err(res.error);
    return ok(undefined);
  }
}
