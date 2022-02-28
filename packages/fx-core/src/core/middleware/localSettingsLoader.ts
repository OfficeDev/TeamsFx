// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Middleware, NextFunction } from "@feathersjs/hooks/lib";
import { AzureSolutionSettings, err, Inputs, Plugin } from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import { CoreHookContext, NoProjectOpenedError, PathNotExistError } from "..";
import { LocalSettingsProvider } from "../../common/localSettingsProvider";
import { PluginNames } from "../../plugins/solution/fx-solution/constants";
import { getActivatedResourcePlugins } from "../../plugins/solution/fx-solution/ResourcePluginContainer";
import { ObjectIsUndefinedError } from "../error";
import { shouldIgnored } from "./projectSettingsLoader";
import { isConfigUnifyEnabled, IsSimpleAuthEnabled } from "../../common/tools";

export const LocalSettingsLoaderMW: Middleware = async (
  ctx: CoreHookContext,
  next: NextFunction
) => {
  if (!shouldIgnored(ctx) && !isConfigUnifyEnabled()) {
    const inputs = ctx.arguments[ctx.arguments.length - 1] as Inputs;
    if (!inputs.projectPath) {
      ctx.result = err(NoProjectOpenedError());
      return;
    }

    const projectPathExist = await fs.pathExists(inputs.projectPath);
    if (!projectPathExist) {
      ctx.result = err(PathNotExistError(inputs.projectPath));
      return;
    }

    if (!ctx.projectSettings) {
      ctx.result = err(new ObjectIsUndefinedError("projectSettings"));
      return;
    }

    const solutionSettings = ctx.projectSettings.solutionSettings as
      | AzureSolutionSettings
      | undefined;
    if (solutionSettings) {
      const selectedPlugins: Plugin[] = getActivatedResourcePlugins(solutionSettings);
      const hasFrontend = selectedPlugins?.some((plugin) => plugin.name === PluginNames.FE);
      const hasBackend = selectedPlugins?.some((plugin) => plugin.name === PluginNames.FUNC);
      const hasBot = selectedPlugins?.some((plugin) => plugin.name === PluginNames.BOT);
      const hasSimpleAuth = IsSimpleAuthEnabled(ctx.projectSettings);
      const hasAAD = selectedPlugins?.some((plugin) => plugin.name === PluginNames.AAD);

      const localSettingsProvider = new LocalSettingsProvider(inputs.projectPath);
      let exists = await fs.pathExists(localSettingsProvider.localSettingsFilePath);
      if (exists) {
        const localSettings = await fs.readJson(localSettingsProvider.localSettingsFilePath);
        if (!localSettings || Object.keys(localSettings).length === 0) {
          // for empty localSettings.json file, we still need to re-init it!
          exists = false;
        }
      }
      //load two versions to make sure compatible
      if (exists) {
        ctx.localSettings = await localSettingsProvider.loadV2(
          ctx.contextV2?.cryptoProvider,
          hasAAD
        );
      } else {
        ctx.localSettings = localSettingsProvider.initV2(
          hasFrontend,
          hasBackend,
          hasBot,
          hasSimpleAuth,
          hasAAD
        );
      }
      if (ctx.solutionContext) {
        if (exists) {
          ctx.solutionContext.localSettings = await localSettingsProvider.load(
            ctx.solutionContext.cryptoProvider
          );
        } else {
          ctx.solutionContext.localSettings = localSettingsProvider.init(
            hasFrontend,
            hasBackend,
            hasBot,
            hasSimpleAuth,
            hasAAD
          );
        }
      }
    }
  }

  await next();
};
