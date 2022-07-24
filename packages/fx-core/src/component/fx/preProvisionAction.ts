// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ContextV3,
  Effect,
  err,
  FunctionAction,
  FxError,
  InputsWithProjectPath,
  ok,
  ProvisionContextV3,
  Result,
  UserError,
  v2,
  v3,
  Void,
} from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../common/localizeUtils";
import { hasAzureResourceV3 } from "../../common/projectSettingsHelperV3";
import { globalVars } from "../../core";
import { SolutionError, SolutionSource } from "../../plugins";
import { updateResourceBaseName } from "../../plugins/solution/fx-solution/arm";
import { resourceGroupHelper } from "../../plugins/solution/fx-solution/utils/ResourceGroupHelper";
import { BuiltInFeaturePluginNames } from "../../plugins/solution/fx-solution/v3/constants";
import {
  askForProvisionConsent,
  fillInAzureConfigs,
  getM365TenantId,
} from "../../plugins/solution/fx-solution/v3/provision";
import { ComponentNames } from "../constants";
import fs from "fs-extra";

export class FxPreProvisionAction implements FunctionAction {
  name = "fx.preProvision";
  type: "function" = "function";
  async execute(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<Effect[], FxError>> {
    const ctx = context as ProvisionContextV3;
    const envInfo = ctx.envInfo;
    // 1. check M365 tenant
    envInfo.state[ComponentNames.AppManifest] = envInfo.state[ComponentNames.AppManifest] || {};
    envInfo.state.solution = envInfo.state.solution || {};
    const appManifest = envInfo.state[ComponentNames.AppManifest];
    const solutionConfig = envInfo.state.solution;
    solutionConfig.provisionSucceeded = false;
    const tenantIdInConfig = appManifest.tenantId;
    const tenantIdInTokenRes = await getM365TenantId(ctx.tokenProvider.m365TokenProvider);
    if (tenantIdInTokenRes.isErr()) {
      return err(tenantIdInTokenRes.error);
    }
    const tenantIdInToken = tenantIdInTokenRes.value;
    if (tenantIdInConfig && tenantIdInToken && tenantIdInToken !== tenantIdInConfig) {
      const checkM365TenantRes = await checkProvisionM365Tenant(
        ctx,
        envInfo,
        tenantIdInToken,
        tenantIdInConfig,
        inputs.projectPath
      );
      if (checkM365TenantRes.isErr()) {
        return err(checkM365TenantRes.error);
      }
    }

    appManifest.tenantId = tenantIdInToken;
    solutionConfig.teamsAppTenantId = tenantIdInToken;
    globalVars.m365TenantId = tenantIdInToken;

    // 3. check Azure configs
    if (hasAzureResourceV3(ctx.projectSetting) && envInfo.envName !== "local") {
      // ask common question and fill in solution config
      const solutionConfigRes = await fillInAzureConfigs(ctx, inputs, envInfo, ctx.tokenProvider);
      if (solutionConfigRes.isErr()) {
        return err(solutionConfigRes.error);
      }

      if (!solutionConfigRes.value.hasSwitchedSubscription) {
        // ask for provision consent
        const consentResult = await askForProvisionConsent(
          ctx,
          ctx.tokenProvider.azureAccountProvider,
          envInfo
        );
        if (consentResult.isErr()) {
          return err(consentResult.error);
        }
      }

      // create resource group if needed
      if (solutionConfig.needCreateResourceGroup) {
        const createRgRes = await resourceGroupHelper.createNewResourceGroup(
          solutionConfig.resourceGroupName,
          ctx.tokenProvider.azureAccountProvider,
          solutionConfig.subscriptionId,
          solutionConfig.location
        );
        if (createRgRes.isErr()) {
          return err(createRgRes.error);
        }
      }

      if (solutionConfigRes.value.hasSwitchedSubscription) {
        updateResourceBaseName(inputs.projectPath, ctx.projectSetting.appName, envInfo.envName);
      }
    }
    return ok([]);
  }
}

export async function checkProvisionM365Tenant(
  ctx: v2.Context,
  envInfo: v3.EnvInfoV3,
  tenantIdInToken: string,
  tenantIdInConfig: string,
  projectPath: string | undefined
): Promise<Result<checkM365TenantResult, FxError>> {
  const hasSwitchedM365Tenat =
    !!tenantIdInConfig && !!tenantIdInToken && tenantIdInToken !== tenantIdInConfig;
  if (hasSwitchedM365Tenat) {
    if (envInfo.envName === "local") {
      return handleSwitchLocalDebugM365Tenant(
        ctx,
        envInfo,
        tenantIdInToken,
        tenantIdInConfig,
        projectPath
      );
    }
    const confirmResult = await askForM365TenantConfirm(
      ctx,
      envInfo,
      tenantIdInConfig,
      tenantIdInToken
    );
    if (confirmResult.isErr()) {
      return err(confirmResult.error);
    } else {
      const keysToClear = [
        BuiltInFeaturePluginNames.bot,
        BuiltInFeaturePluginNames.aad,
        ComponentNames.TeamsBot,
        ComponentNames.AadApp,
      ];
      const keys = Object.keys(envInfo.state);
      for (let index = 0; index < keys.length; index++) {
        if (keysToClear.includes(keys[index])) {
          envInfo.state[keys[index]] = {};
        }
      }

      // todo: update bot resource base name. see: https://github.com/OfficeDev/TeamsFx/compare/dev...yuqzho/switch-m365?expand=1
    }

    return ok({ hasSwitchedM365Tenant: true });
  }

  return ok({ hasSwitchedM365Tenant: false });
}

async function askForM365TenantConfirm(
  ctx: v2.Context,
  envInfo: v3.EnvInfoV3,
  localDebugTenantId: string,
  maybeM365TenantId: string
): Promise<Result<Void, FxError>> {
  const msg = getLocalizedString(
    // ask for unmatched tenant
    "core.localDebug.tenantConfirmNotice",
    localDebugTenantId,
    maybeM365TenantId,
    ""
  ); // TODO: ui
  const confirmRes = await ctx.userInteraction.showMessage("warn", msg, true, "Continue");
  const confirm = confirmRes?.isOk() ? confirmRes.value : undefined;

  if (confirm !== "Continue") {
    return err(new UserError(SolutionSource, "CancelLocalDebug", "CancelLocalDebug"));
  }

  return ok([]);
}

export async function handleSwitchLocalDebugM365Tenant(
  ctx: v2.Context,
  envInfo: v3.EnvInfoV3,
  tenantIdInToken: string,
  tenantIdInConfig: string,
  projectPath: string | undefined
): Promise<Result<checkM365TenantResult, FxError>> {
  const hasSwitchedM365Tenat =
    !!tenantIdInConfig && !!tenantIdInToken && tenantIdInToken !== tenantIdInConfig;
  if (hasSwitchedM365Tenat) {
    if (
      projectPath !== undefined &&
      (await fs.pathExists(`${projectPath}/bot/.notification.localstore.json`))
    ) {
      const errorMessage = getLocalizedString(
        "core.localDebug.tenantConfirmNotice",
        tenantIdInConfig,
        tenantIdInToken,
        "bot/.notification.localstore.json"
      );
      return err(
        new UserError("Solution", SolutionError.CannotLocalDebugInDifferentTenant, errorMessage)
      );
    }
    const confirmResult = await askForM365TenantConfirm(
      ctx,
      envInfo,
      tenantIdInConfig,
      tenantIdInToken
    );
    if (confirmResult.isErr()) {
      return err(confirmResult.error);
    } else {
      const keys = Object.keys(envInfo.state);
      for (let index = 0; index < keys.length; index++) {
        envInfo.state[keys[index]] = {};
      }
      // todo: delete local files.
    }

    return ok({ hasSwitchedM365Tenant: true });
  }

  return ok({ hasSwitchedM365Tenant: false });
}

export interface checkM365TenantResult {
  hasSwitchedM365Tenant: boolean;
}
