import { FxError, Inputs, Json, SystemError, TokenProvider, v2 } from "@microsoft/teamsfx-api";
import { isUndefined } from "lodash";
import { Container } from "typedi";
import { isExistingTabApp } from "../../../../common/projectSettingsHelper";
import { AppStudioScopes } from "../../../../common/tools";
import { environmentManager } from "../../../../core/environment";
import { PermissionRequestFileProvider } from "../../../../core/permissionRequest";
import { PluginNames, SolutionError } from "../constants";
import {
  configLocalDebugSettings,
  configLocalEnvironment,
  setupLocalDebugSettings,
  setupLocalEnvironment,
} from "../debug/provisionLocal";
import { ResourcePluginsV2 } from "../ResourcePluginContainer";
import { executeConcurrently } from "./executor";
import {
  checkLocalDebugM365Tenant,
  ensurePermissionRequest,
  getAzureSolutionSettings,
  getSelectedPlugins,
  isAzureProject,
  loadTeamsAppTenantIdForLocal,
} from "./utils";

export async function provisionLocalResource(
  ctx: v2.Context,
  inputs: Inputs,
  localSettings: Json,
  tokenProvider: TokenProvider,
  envInfo?: v2.EnvInfoV2
): Promise<v2.FxResult<Json, FxError>> {
  if (inputs.projectPath === undefined) {
    return new v2.FxFailure(
      new SystemError("Solution", SolutionError.InternelError, "projectPath is undefined")
    );
  }
  const azureSolutionSettings = getAzureSolutionSettings(ctx);
  if (isAzureProject(azureSolutionSettings)) {
    if (ctx.permissionRequestProvider === undefined) {
      ctx.permissionRequestProvider = new PermissionRequestFileProvider(inputs.projectPath);
    }
    const result = await ensurePermissionRequest(
      azureSolutionSettings!,
      ctx.permissionRequestProvider
    );
    if (result.isErr()) {
      return new v2.FxFailure(result.error);
    }
  }

  // Just to trigger M365 login before the concurrent execution of localDebug.
  // Because concurrent execution of localDebug may getAccessToken() concurrently, which
  // causes 2 M365 logins before the token caching in common lib takes effect.
  const appStudioTokenRes = await tokenProvider.m365TokenProvider.getAccessToken({
    scopes: AppStudioScopes,
  });
  if (appStudioTokenRes.isErr()) {
    return new v2.FxFailure(appStudioTokenRes.error);
  }
  // Pop-up window to confirm if local debug in another tenant
  let localDebugTenantId = "";
  localDebugTenantId = envInfo?.state.solution.teamsAppTenantId;

  const checklocalDebugM365TenantRes = await checkLocalDebugM365Tenant(
    ctx,
    envInfo,
    localDebugTenantId,
    tokenProvider.m365TokenProvider,
    inputs.projectPath
  );
  if (checklocalDebugM365TenantRes.isErr()) {
    //return new v2.FxFailure(m365TenantMatches.error);
    return new v2.FxFailure(checklocalDebugM365TenantRes.error);
  }

  const plugins = getSelectedPlugins(ctx.projectSetting);
  if (isExistingTabApp(ctx.projectSetting)) {
    // for existing tab app, enable app studio plugin when solution settings is empty.
    const appStudioPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.AppStudioPlugin);
    if (!plugins.find((p) => p.name === appStudioPlugin.name)) {
      plugins.push(appStudioPlugin);
    }
  }

  const provisionLocalResourceThunks = plugins
    .filter((plugin) => !isUndefined(plugin.provisionLocalResource))
    .map((plugin) => {
      return {
        pluginName: `${plugin.name}`,
        taskName: "provisionLocalResource",
        // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
        thunk: () =>
          plugin.provisionLocalResource!(ctx, inputs, localSettings, tokenProvider, envInfo),
      };
    });

  const provisionResult = await executeConcurrently(provisionLocalResourceThunks, ctx.logProvider);
  if (provisionResult.kind !== "success") {
    return provisionResult;
  }

  const localEnvSetupResult = await setupLocalEnvironment(ctx, inputs, envInfo!);

  if (localEnvSetupResult.isErr()) {
    return new v2.FxPartialSuccess(envInfo!, localEnvSetupResult.error);
  }

  const aadPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.AadPlugin);
  if (isAzureProject(azureSolutionSettings)) {
    if (plugins.some((plugin) => plugin.name === aadPlugin.name) && aadPlugin.executeUserTask) {
      const result = await aadPlugin.executeUserTask(
        ctx,
        inputs,
        {
          namespace: `${PluginNames.SOLUTION}/${PluginNames.AAD}`,
          method: "setApplicationInContext",
          params: { isLocal: false },
        },
        localSettings,
        envInfo!,
        tokenProvider
      );
      if (result.isErr()) {
        return new v2.FxPartialSuccess(localSettings, result.error);
      }
    }
  }

  const appStudioTokenJsonRes = await tokenProvider.m365TokenProvider.getJsonObject({
    scopes: AppStudioScopes,
  });
  const appStudioTokenJson = appStudioTokenJsonRes.isOk() ? appStudioTokenJsonRes.value : undefined;

  const parseTenantIdresult = loadTeamsAppTenantIdForLocal(
    localSettings as v2.LocalSettings,
    appStudioTokenJson,
    envInfo
  );
  if (parseTenantIdresult.isErr()) {
    return new v2.FxFailure(parseTenantIdresult.error);
  }

  const configureLocalResourceThunks = plugins
    .filter((plugin) => !isUndefined(plugin.configureLocalResource))
    .map((plugin) => {
      return {
        pluginName: `${plugin.name}`,
        taskName: "configureLocalResource",
        // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
        thunk: () =>
          plugin.configureLocalResource!(ctx, inputs, localSettings, tokenProvider, envInfo),
      };
    });

  const configureResourceResult = await executeConcurrently(
    configureLocalResourceThunks,
    ctx.logProvider
  );

  if (configureResourceResult.kind !== "success") {
    if (configureResourceResult.kind === "partialSuccess") {
      return new v2.FxPartialSuccess(localSettings, configureResourceResult.error);
    }
    return new v2.FxFailure(configureResourceResult.error);
  }

  const localConfigResult = await configLocalEnvironment(ctx, inputs, envInfo!);

  if (localConfigResult.isErr()) {
    return new v2.FxPartialSuccess(envInfo!, localConfigResult.error);
  }

  return new v2.FxSuccess(localSettings);
}
