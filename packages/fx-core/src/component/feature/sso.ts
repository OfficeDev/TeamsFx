// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Action,
  ContextV3,
  Effect,
  err,
  FxError,
  InputsWithProjectPath,
  MaybePromise,
  ok,
  Platform,
  Result,
  Stage,
  v3,
} from "@microsoft/teamsfx-api";
import "reflect-metadata";
import { Container, Service } from "typedi";
import { convertToAlphanumericOnly } from "../../common/utils";
import "../connection/azureWebAppConfig";
import { ComponentNames } from "../constants";
import { generateLocalDebugSettings } from "../debug";
import { AadApp } from "../resource/aadApp/aadApp";
import { AppManifest } from "../resource/appManifest/appManifest";
import "../resource/azureSql";
import "../resource/identity";
import { generateConfigBiceps, persistBiceps } from "../utils";
import { getComponent } from "../workflow";

@Service("sso")
export class SSO {
  name = "sso";

  add(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): MaybePromise<Result<Action | undefined, FxError>> {
    if (inputs.platform == Platform.CLI_HELP) {
      return ok(undefined);
    }
    const action: Action = {
      type: "function",
      name: "sso.add",
      execute: async (context, inputs) => {
        const updates = getUpdateComponents(context, inputs);
        const effects: Effect[] = [];

        // generate manifest
        const aadApp = Container.get<AadApp>(ComponentNames.AadApp);
        {
          const res = await aadApp.generateManifest(context, inputs);
          if (res.isErr()) return err(res.error);
          effects.push("generate aad manifest");
        }

        // generate bicep
        {
          const res = await aadApp.generateBicep(context, inputs);
          if (res.isErr()) return err(res.error);
          const bicepRes = await persistBiceps(
            inputs.projectPath,
            convertToAlphanumericOnly(context.projectSetting.appName),
            res.value
          );
          if (bicepRes.isErr()) return bicepRes;
          effects.push("generate aad bicep");
        }

        // generate config bicep
        {
          const res = await generateConfigBiceps(context, inputs);
          if (res.isErr()) return err(res.error);
          effects.push("generate config biceps");
        }

        // generate auth files
        {
          const res = await aadApp.generateAuthFiles(context, inputs, updates.tab!, updates.bot!);
          if (res.isErr()) return err(res.error);
          effects.push("generate auth files");
        }
        // update app manifest
        {
          const capabilities: v3.ManifestCapability[] = [{ name: "WebApplicationInfo" }];
          const appManifest = Container.get<AppManifest>(ComponentNames.AppManifest);
          const res = await appManifest.addCapability(inputs, capabilities);
          if (res.isErr()) return err(res.error);
          effects.push("update app manifest");
        }

        // local debug settings
        {
          const res = await generateLocalDebugSettings(context, inputs);
          if (res.isErr()) return err(res.error);
          effects.push("generate local debug configs");
        }

        // config sso
        if (updates.aad) {
          context.projectSetting.components.push({
            name: ComponentNames.AadApp,
            provision: true,
            deploy: true,
          });
        }
        if (updates.tab) {
          const teamsTabComponent = getComponent(context.projectSetting, ComponentNames.TeamsTab);
          teamsTabComponent!.sso = true;
        }
        if (updates.bot) {
          const teamsBotComponent = getComponent(context.projectSetting, ComponentNames.TeamsBot);
          teamsBotComponent!.sso = true;
        }
        effects.push("config sso");
        return ok(effects);
      },
    };
    return ok(action);
  }
}

export interface updateComponents {
  bot?: boolean;
  tab?: boolean;
  aad?: boolean;
}
function getUpdateComponents(context: ContextV3, inputs: InputsWithProjectPath): updateComponents {
  if (inputs.stage === Stage.create) {
    return {
      tab: true,
      aad: true,
    };
  }
  let needsBot = false;
  let needsTab = false;
  const aadComponent = getComponent(context.projectSetting, ComponentNames.AadApp);
  const teamsBotComponent = getComponent(context.projectSetting, ComponentNames.TeamsBot);
  if (teamsBotComponent && !teamsBotComponent.sso) {
    needsBot = teamsBotComponent.hosting !== ComponentNames.Function;
  }
  const teamsTabComponent = getComponent(context.projectSetting, ComponentNames.TeamsTab);
  if (teamsTabComponent && !teamsTabComponent.sso) {
    needsTab = true;
  }
  return {
    bot: needsBot,
    tab: needsTab,
    aad: !aadComponent,
  };
}
