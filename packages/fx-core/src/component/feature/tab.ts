// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Action,
  CallAction,
  ContextV3,
  err,
  FunctionAction,
  FxError,
  InputsWithProjectPath,
  MaybePromise,
  ok,
  Platform,
  Result,
  Stage,
} from "@microsoft/teamsfx-api";
import "reflect-metadata";
import Container, { Service } from "typedi";
import { format } from "util";
import { getLocalizedString } from "../../common/localizeUtils";
import { isVSProject } from "../../common/projectSettingsHelper";
import { globalVars } from "../../core/globalVars";
import { CoreQuestionNames } from "../../core/question";
import { FrontendPathInfo } from "../../plugins/resource/frontend/constants";
import {
  AzureSolutionQuestionNames,
  TabNonSsoItem,
} from "../../plugins/solution/fx-solution/question";
import { ComponentNames, Scenarios } from "../constants";
import { Plans } from "../messages";
import { getComponent, getComponentByScenario, runActionByName } from "../workflow";
import { assign, cloneDeep } from "lodash";
import { hasTab } from "../../common/projectSettingsHelperV3";
import { generateConfigBiceps } from "../utils";
import { TabCodeProvider } from "../code/tabCode";

@Service("teams-tab")
export class TeamsTab {
  name = "teams-tab";
  add(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): MaybePromise<Result<Action | undefined, FxError>> {
    const action: FunctionAction = {
      name: "teams-tab.add",
      type: "function",
      execute: async (context, inputs) => {
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
        // 1. scaffold tab and add tab config
        let tabConfig = getComponent(projectSettings, ComponentNames.TeamsTab);
        if (!tabConfig) {
          const clonedInputs = cloneDeep(inputs);
          clonedInputs.folder ||=
            inputs[CoreQuestionNames.ProgrammingLanguage] === "csharp"
              ? ""
              : FrontendPathInfo.WorkingDir;
          clonedInputs.language = inputs[CoreQuestionNames.ProgrammingLanguage];
          const tabCode = Container.get(ComponentNames.TabCode) as TabCodeProvider;
          // const res = await runActionByName("tab-code.generate", context, clonedInputs);
          const res = await tabCode.generateNew(
            context,
            inputs.projectPath,
            clonedInputs.folder,
            clonedInputs.language,
            inputs[CoreQuestionNames.SafeProjectName]
          );
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
            const res = await runActionByName("bicep.init", context, inputs);
            if (res.isErr()) return err(res.error);
          }
          // 2.1 hosting bicep
          const hostingConfig = getComponentByScenario(
            projectSettings,
            inputs.hosting,
            Scenarios.Tab
          );
          if (!hostingConfig) {
            const clonedInputs = cloneDeep(inputs);
            assign(clonedInputs, {
              componentId: ComponentNames.TeamsTab,
              scenario: Scenarios.Tab,
            });
            const res = await runActionByName(
              inputs.hosting + ".generateBicep",
              context,
              clonedInputs
            );
            if (res.isErr()) return err(res.error);
            projectSettings.components.push({
              name: inputs.hosting,
              scenario: Scenarios.Tab,
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
            const res = await runActionByName("identity.generateBicep", context, clonedInputs);
            if (res.isErr()) return err(res.error);
            projectSettings.components.push({
              name: ComponentNames.Identity,
              provision: true,
            });
            effects.push(Plans.generateBicepAndConfig(ComponentNames.Identity));
          }

          // 2.3 add sso
          if (
            inputs.stage === Stage.create &&
            inputs[AzureSolutionQuestionNames.Features] !== TabNonSsoItem.id
          ) {
            const res = await runActionByName("sso.add", context, inputs);
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
            const res = await runActionByName("debug.generateLocalDebugSettings", context, inputs);
            if (res.isErr()) return err(res.error);
            effects.push("generate local debug configs");
          }
        }

        // 5. app-manifest.addCapability
        {
          const capabilities = [{ name: "staticTab" }];
          if (!hasTab(projectSettings)) {
            capabilities.push({ name: "configurableTab" });
          }
          const clonedInputs = cloneDeep(inputs);
          assign(clonedInputs, {
            capabilities: capabilities,
          });
          const res = await runActionByName("app-manifest.addCapability", context, clonedInputs);
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
        return ok(effects);
      },
    };
    return ok(action);
  }
  configure(): MaybePromise<Result<Action | undefined, FxError>> {
    return ok(configureTab);
  }
  build(): MaybePromise<Result<Action | undefined, FxError>> {
    return ok(buildTab);
  }
}

//   private addTabAction(context: ContextV3, inputs: InputsWithProjectPath): Action {
//     inputs.hosting = resolveHosting(context, inputs);
//     const actions: Action[] = [];
//     this.setupCode(actions, context);
//     this.setupBicep(actions, context, inputs);
//     this.setupCapabilities(actions);
//     this.setupConfiguration(actions, context);
//     actions.push(showTabAlreadyAddMessage);
//     return {
//       type: "group",
//       name: "teams-tab.add",
//       mode: "sequential",
//       actions: actions,
//     };
//   }

//   private setupConfiguration(actions: Action[], context: ContextV3): Action[] {
//     if (hasTab(context)) {
//       return actions;
//     }
//     actions.push(addSSO);
//     actions.push(configTab);
//     return actions;
//   }

//   private setupCode(actions: Action[], context: ContextV3): Action[] {
//     if (hasTab(context)) {
//       return actions;
//     }
//     actions.push(generateCode);
//     actions.push(initLocalDebug);
//     return actions;
//   }

//   private setupBicep(
//     actions: Action[],
//     context: ContextV3,
//     inputs: InputsWithProjectPath
//   ): Action[] {
//     if (hasTab(context)) {
//       return actions;
//     }

//     actions.push(initBicep);
//     const hosting = resolveHosting(context, inputs);
//     if (hosting) {
//       actions.push(generateBicep(hosting, this.name));
//     }
//     // TODO: connect AAD for blazor web app
//     actions.push(identityAction);
//     actions.push(configureApim);
//     return actions;
//   }

//   private setupCapabilities(actions: Action[]): Action[] {
//     actions.push(addTabCapability);
//     return actions;
//   }
// }

// function hasTab(context: ContextV3): boolean {
//   const tab = getComponent(context.projectSetting, ComponentNames.TeamsTab);
//   return tab != undefined; // using != to match both undefined and null
// }

// const addTabCapability: Action = {
//   name: "call:app-manifest.addCapability",
//   type: "call",
//   required: true,
//   targetAction: "app-manifest.addCapability",
//   pre: (context: ContextV3, inputs: InputsWithProjectPath) => {
//     const capabilities = [{ name: "staticTab" }];
//     if (!hasTab(context)) {
//       capabilities.push({ name: "configurableTab" });
//     }
//     inputs.capabilities = capabilities;
//     return ok(undefined);
//   },
// };

// const configTab: Action = {
//   name: "fx.configTab",
//   type: "function",
//   plan: (context: ContextV3, inputs: InputsWithProjectPath) => {
//     const tabConfig = getComponent(context.projectSetting, ComponentNames.TeamsTab);
//     if (tabConfig) {
//       return ok([]);
//     }
//     return ok([Plans.addFeature("Tab")]);
//   },
//   execute: async (context: ContextV3, inputs: InputsWithProjectPath) => {
//     const hosting = resolveHosting(context, inputs);
//     const language =
//       inputs?.[CoreQuestionNames.ProgrammingLanguage] ||
//       context.projectSetting.programmingLanguage ||
//       "javascript";
//     const projectSettings = context.projectSetting as ProjectSettingsV3;
//     const tabConfig = getComponent(projectSettings, ComponentNames.TeamsTab);
//     if (tabConfig) {
//       return ok([]);
//     }
//     // add teams-tab
//     projectSettings.components.push({
//       name: ComponentNames.TeamsTab,
//       hosting: hosting,
//       deploy: true,
//       provision: language != "csharp",
//       build: true,
//       folder: inputs.folder || language === "csharp" ? "" : FrontendPathInfo.WorkingDir,
//     });
//     ensureComponentConnections(projectSettings);
//     globalVars.isVS = isVSProject(projectSettings);
//     projectSettings.programmingLanguage ||= inputs[CoreQuestionNames.ProgrammingLanguage];
//     return ok([Plans.addFeature("Tab")]);
//   },
// };

// const addSSO: CallAction = {
//   type: "call",
//   name: "call:sso.add",
//   targetAction: "sso.add",
//   required: true,
//   condition: (context: ContextV3, inputs: InputsWithProjectPath) => {
//     return ok(inputs[CoreQuestionNames.Features] !== TabNonSsoItem.id);
//   },
// };

// const showTabAlreadyAddMessage: Action = {
//   name: "teams-tab.showTabAlreadyAddMessage",
//   type: "function",
//   plan: (context: ContextV3, inputs: InputsWithProjectPath) => {
//     return ok([]);
//   },
//   execute: async (context: ContextV3, inputs: InputsWithProjectPath) => {
//     const msg =
//       inputs.platform === Platform.CLI
//         ? getLocalizedString("core.addCapability.addCapabilityNoticeForCli")
//         : getLocalizedString("core.addCapability.addCapabilitiesNoticeForCli");
//     context.userInteraction.showMessage("info", format(msg, "Tab"), false);
//     return ok([]);
//   },
// };

// const generateCode: Action = {
//   name: "call:tab-code.generate",
//   type: "call",
//   required: true,
//   targetAction: "tab-code.generate",
//   pre: (context: ContextV3, inputs: InputsWithProjectPath) => {
//     const language =
//       inputs?.[CoreQuestionNames.ProgrammingLanguage] ||
//       context.projectSetting.programmingLanguage ||
//       "javascript";
//     inputs.folder ||= language === "csharp" ? "" : FrontendPathInfo.WorkingDir;
//     inputs.language ||= language;
//     return ok(undefined);
//   },
// };
// const initLocalDebug: Action = {
//   name: "call:debug.generateLocalDebugSettings",
//   type: "call",
//   required: true,
//   targetAction: "debug.generateLocalDebugSettings",
// };

// const initBicep: Action = {
//   type: "call",
//   targetAction: "bicep.init",
//   required: true,
// };

// const generateBicep: (hosting: string, componentId: string) => Action = (hosting, componentId) => ({
//   name: `call:${hosting}.generateBicep`,
//   type: "call",
//   required: true,
//   targetAction: `${hosting}.generateBicep`,
//   inputs: {
//     scenario: "Tab",
//     componentId: componentId,
//   },
//   post: (context) => {
//     // add hosting component
//     context.projectSetting?.components?.push({
//       name: hosting,
//       connections: [ComponentNames.TeamsTab],
//       provision: true,
//       scenario: Scenarios.Tab,
//     });
//     ensureComponentConnections(context.projectSetting);
//     return ok(undefined);
//   },
// });

const configureApim: CallAction = {
  name: "call:apim-config.generateBicep",
  type: "call",
  required: true,
  targetAction: "apim-config.generateBicep",
  condition: (context) => {
    return ok(getComponent(context.projectSetting, ComponentNames.APIM) !== undefined);
  },
};

const configureTab: CallAction = {
  name: "teams-tab.configure",
  type: "call",
  targetAction: "tab-code.configure",
  required: true,
};
const buildTab: CallAction = {
  name: "teams-tab.build",
  type: "call",
  targetAction: "tab-code.build",
  required: true,
};
