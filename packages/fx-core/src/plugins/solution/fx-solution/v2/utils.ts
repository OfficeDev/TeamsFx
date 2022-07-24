import {
  v2,
  Inputs,
  FxError,
  Result,
  ok,
  err,
  AzureSolutionSettings,
  Void,
  PermissionRequestProvider,
  Json,
  SolutionContext,
  Plugin,
  ProjectSettings,
  UserError,
  SystemError,
  M365TokenProvider,
  v3,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import { LocalSettingsTeamsAppKeys } from "../../../../common/localSettingsConstants";
import { AppStudioScopes, isAadManifestEnabled } from "../../../../common/tools";
import {
  GLOBAL_CONFIG,
  SolutionError,
  SOLUTION_PROVISION_SUCCEEDED,
  SolutionSource,
  PluginNames,
} from "../constants";
import {
  AzureResourceApim,
  AzureResourceFunction,
  AzureResourceSQL,
  AzureSolutionQuestionNames,
  BotOptionItem,
  BotScenario,
  CommandAndResponseOptionItem,
  HostTypeOptionAzure,
  HostTypeOptionSPFx,
  M365SearchAppOptionItem,
  M365SsoLaunchPageOptionItem,
  MessageExtensionItem,
  NotificationOptionItem,
  TabNonSsoItem,
  TabOptionItem,
  TabSPFxItem,
  TabSsoItem,
} from "../question";
import { getActivatedV2ResourcePlugins, getAllV2ResourcePlugins } from "../ResourcePluginContainer";
import { getPluginContext } from "../utils/util";
import { PluginsWithContext } from "../types";
import { getDefaultString, getLocalizedString } from "../../../../common/localizeUtils";
import { handleSwitchLocalDebugM365Tenant } from "../../../../component/fx/preProvisionAction";

export function getSelectedPlugins(projectSettings: ProjectSettings): v2.ResourcePlugin[] {
  return getActivatedV2ResourcePlugins(projectSettings);
}

export function getAzureSolutionSettings(ctx: v2.Context): AzureSolutionSettings | undefined {
  return ctx.projectSetting.solutionSettings as AzureSolutionSettings | undefined;
}

export function isAzureProject(azureSettings: AzureSolutionSettings | undefined): boolean {
  return azureSettings !== undefined && HostTypeOptionAzure.id === azureSettings.hostType;
}

export function isBotProject(azureSettings: AzureSolutionSettings | undefined): boolean {
  return (
    azureSettings !== undefined &&
    (azureSettings.capabilities?.includes(BotOptionItem.id) ||
      azureSettings.capabilities?.includes(MessageExtensionItem.id))
  );
}

export interface BotTroubleShootMessage {
  troubleShootLink: string;
  textForLogging: string;
  textForMsgBox: string;
  textForActionButton: string;
}

export function getBotTroubleShootMessage(isBot: boolean): BotTroubleShootMessage {
  const botTroubleShootLink =
    "https://aka.ms/teamsfx-bot-help#how-can-i-troubleshoot-issues-when-teams-bot-isnt-responding-on-azure";
  const botTroubleShootDesc = getLocalizedString("core.deploy.botTroubleShoot");
  const botTroubleShootLearnMore = getLocalizedString("core.deploy.botTroubleShoot.learnMore");
  const botTroubleShootMsg = `${botTroubleShootDesc} ${botTroubleShootLearnMore}: ${botTroubleShootLink}.`;

  return {
    troubleShootLink: botTroubleShootLink,
    textForLogging: isBot ? botTroubleShootMsg : "",
    textForMsgBox: botTroubleShootDesc,
    textForActionButton: botTroubleShootLearnMore,
  } as BotTroubleShootMessage;
}

export function combineRecords<T>(records: { name: string; result: T }[]): Record<string, T> {
  const ret: Record<v2.PluginName, T> = {};
  for (const record of records) {
    ret[record.name] = record.result;
  }

  return ret;
}

export function extractSolutionInputs(record: Json): v2.SolutionInputs {
  return {
    resourceNameSuffix: record["resourceNameSuffix"],
    resourceGroupName: record["resourceGroupName"],
    location: record["location"],
    teamsAppTenantId: record["teamsAppTenantId"],
    remoteTeamsAppId: undefined,
    subscriptionId: record["subscriptionId"],
    provisionSucceeded: record[SOLUTION_PROVISION_SUCCEEDED],
    tenantId: record["tenantId"],
  };
}

export function setActivatedResourcePluginsV2(projectSettings: ProjectSettings): void {
  const activatedPluginNames = getAllV2ResourcePlugins()
    .filter((p) => p.activate && p.activate(projectSettings) === true)
    .map((p) => p.name);
  projectSettings.solutionSettings!.activeResourcePlugins = activatedPluginNames;
}

export async function ensurePermissionRequest(
  solutionSettings: AzureSolutionSettings,
  permissionRequestProvider: PermissionRequestProvider
): Promise<Result<Void, FxError>> {
  if (!isAzureProject(solutionSettings)) {
    return err(
      new UserError(
        SolutionSource,
        SolutionError.CannotUpdatePermissionForSPFx,
        "Cannot update permission for SPFx project"
      )
    );
  }

  if (!isAadManifestEnabled()) {
    const result = await permissionRequestProvider.checkPermissionRequest();
    if (result && result.isErr()) {
      return result.map(err);
    }
  }

  return ok(Void);
}

export function parseTeamsAppTenantId(
  appStudioToken?: Record<string, unknown>
): Result<string, FxError> {
  if (appStudioToken === undefined) {
    return err(
      new SystemError(
        SolutionSource,
        SolutionError.NoAppStudioToken,
        "Graph token json is undefined"
      )
    );
  }

  const teamsAppTenantId = appStudioToken["tid"];
  if (
    teamsAppTenantId === undefined ||
    !(typeof teamsAppTenantId === "string") ||
    teamsAppTenantId.length === 0
  ) {
    return err(
      new SystemError(
        SolutionSource,
        SolutionError.NoTeamsAppTenantId,
        getDefaultString("error.NoTeamsAppTenantId"),
        getLocalizedString("error.NoTeamsAppTenantId")
      )
    );
  }
  return ok(teamsAppTenantId);
}

export function parseUserName(appStudioToken?: Record<string, unknown>): Result<string, FxError> {
  if (appStudioToken === undefined) {
    return err(
      new SystemError("Solution", SolutionError.NoAppStudioToken, "Graph token json is undefined")
    );
  }

  const userName = appStudioToken["upn"];
  if (userName === undefined || !(typeof userName === "string") || userName.length === 0) {
    return err(
      new SystemError(
        "Solution",
        SolutionError.NoUserName,
        "Cannot find user name from App Studio token."
      )
    );
  }
  return ok(userName);
}

export async function checkWhetherLocalDebugM365TenantMatches(
  localDebugTenantId?: string,
  m365TokenProvider?: M365TokenProvider,
  projectPath?: string
): Promise<Result<Void, FxError>> {
  if (localDebugTenantId) {
    const appStudioTokenJsonRes = await m365TokenProvider?.getJsonObject({
      scopes: AppStudioScopes,
    });
    const appStudioTokenJson = appStudioTokenJsonRes?.isOk()
      ? appStudioTokenJsonRes.value
      : undefined;
    const maybeM365TenantId = parseTeamsAppTenantId(appStudioTokenJson);
    if (maybeM365TenantId.isErr()) {
      return maybeM365TenantId;
    }
    const maybeM365UserAccount = parseUserName(appStudioTokenJson);
    if (maybeM365UserAccount.isErr()) {
      return maybeM365UserAccount;
    }
    if (maybeM365TenantId.value !== localDebugTenantId) {
      const localFiles = [".fx/states/state.local.json"];
      // add notification local file if exist
      if (
        projectPath !== undefined &&
        (await fs.pathExists(`${projectPath}/bot/.notification.localstore.json`))
      ) {
        localFiles.push("bot/.notification.localstore.json");
      }

      const errorMessage = getLocalizedString(
        "core.localDebug.tenantConfirmNotice",
        localDebugTenantId,
        maybeM365UserAccount.value,
        localFiles.join(", ")
      );
      return err(
        new UserError("Solution", SolutionError.CannotLocalDebugInDifferentTenant, errorMessage)
      );
    }
  }

  return ok(Void);
}

export async function checkLocalDebugM365Tenant(
  ctx: v2.Context,
  envInfo: v2.EnvInfoV2 | undefined,
  localDebugTenantId?: string,
  m365TokenProvider?: M365TokenProvider,
  projectPath?: string
): Promise<Result<Void, FxError>> {
  if (localDebugTenantId) {
    const appStudioTokenJsonRes = await m365TokenProvider?.getJsonObject({
      scopes: AppStudioScopes,
    });
    const appStudioTokenJson = appStudioTokenJsonRes?.isOk()
      ? appStudioTokenJsonRes.value
      : undefined;
    const maybeM365TenantId = parseTeamsAppTenantId(appStudioTokenJson);
    if (maybeM365TenantId.isErr()) {
      return maybeM365TenantId;
    }

    const maybeM365UserAccount = parseUserName(appStudioTokenJson);
    if (maybeM365UserAccount.isErr()) {
      return maybeM365UserAccount;
    }

    if (!!envInfo && maybeM365TenantId.value !== localDebugTenantId) {
      const res = await handleSwitchLocalDebugM365Tenant(
        ctx,
        envInfo as v3.EnvInfoV3,
        maybeM365TenantId.value,
        localDebugTenantId,
        projectPath
      );
      return res;
    }
  }

  return ok(Void);
}

// Loads teams app tenant id into local settings.
export function loadTeamsAppTenantIdForLocal(
  localSettings: v2.LocalSettings,
  appStudioToken?: Record<string, unknown>,
  envInfo?: v2.EnvInfoV2
): Result<Void, FxError> {
  return parseTeamsAppTenantId(appStudioToken as Record<string, unknown> | undefined).andThen(
    (teamsAppTenantId) => {
      envInfo!.state.solution.teamsAppTenantId = teamsAppTenantId;
      return ok(Void);
    }
  );
}

export function fillInSolutionSettings(
  projectSettings: ProjectSettings,
  answers: Inputs
): Result<Void, FxError> {
  const solutionSettings = (projectSettings.solutionSettings as AzureSolutionSettings) || {};
  let capabilities = (answers[AzureSolutionQuestionNames.Capabilities] as string[]) || [];
  if (isAadManifestEnabled()) {
    if (capabilities.includes(TabOptionItem.id)) {
      capabilities.push(TabSsoItem.id);
    } else if (capabilities.includes(TabNonSsoItem.id)) {
      const index = capabilities.indexOf(TabNonSsoItem.id);
      capabilities.splice(index, 1);
      capabilities.push(TabOptionItem.id);
    }
  }
  if (!capabilities || capabilities.length === 0) {
    return err(
      new SystemError(SolutionSource, SolutionError.InternelError, "capabilities is empty")
    );
  }
  let hostType = answers[AzureSolutionQuestionNames.HostType] as string;
  if (
    capabilities.includes(NotificationOptionItem.id) ||
    capabilities.includes(CommandAndResponseOptionItem.id)
  ) {
    // find and replace "NotificationOptionItem" and "CommandAndResponseOptionItem" to "BotOptionItem", so it does not impact capabilities in projectSettings.json
    const scenarios: BotScenario[] = [];
    const notificationIndex = capabilities.indexOf(NotificationOptionItem.id);
    if (notificationIndex !== -1) {
      capabilities[notificationIndex] = BotOptionItem.id;
      scenarios.push(BotScenario.NotificationBot);
    }
    const commandAndResponseIndex = capabilities.indexOf(CommandAndResponseOptionItem.id);
    if (commandAndResponseIndex !== -1) {
      capabilities[commandAndResponseIndex] = BotOptionItem.id;
      scenarios.push(BotScenario.CommandAndResponseBot);
    }
    answers[AzureSolutionQuestionNames.Scenarios] = scenarios;
    // dedup
    capabilities = [...new Set(capabilities)];

    hostType = HostTypeOptionAzure.id;
  } else if (
    capabilities.includes(BotOptionItem.id) ||
    capabilities.includes(MessageExtensionItem.id) ||
    capabilities.includes(TabOptionItem.id)
  ) {
    hostType = HostTypeOptionAzure.id;
  } else if (capabilities.includes(TabSPFxItem.id)) {
    // set capabilities to TabOptionItem in case of TabSPFx item, so donot impact capabilities.includes() check overall
    capabilities = [TabOptionItem.id];
    hostType = HostTypeOptionSPFx.id;
  } else if (capabilities.includes(M365SsoLaunchPageOptionItem.id)) {
    capabilities = [TabOptionItem.id];
    if (isAadManifestEnabled()) {
      capabilities.push(TabSsoItem.id);
    }
    hostType = HostTypeOptionAzure.id;
  } else if (capabilities.includes(M365SearchAppOptionItem.id)) {
    capabilities = [MessageExtensionItem.id];
    const scenarios = [M365SearchAppOptionItem.id];
    answers[AzureSolutionQuestionNames.Scenarios] = scenarios;
    hostType = HostTypeOptionAzure.id;
  }
  if (!hostType) {
    return err(
      new SystemError(SolutionSource, SolutionError.InternelError, "hostType is undefined")
    );
  }
  solutionSettings.hostType = hostType;
  let azureResources: string[] | undefined;
  if (hostType === HostTypeOptionAzure.id && capabilities.includes(TabOptionItem.id)) {
    azureResources = answers[AzureSolutionQuestionNames.AzureResources] as string[];
    if (azureResources) {
      if (
        (azureResources.includes(AzureResourceSQL.id) ||
          azureResources.includes(AzureResourceApim.id)) &&
        !azureResources.includes(AzureResourceFunction.id)
      ) {
        azureResources.push(AzureResourceFunction.id);
      }
    } else azureResources = [];
  }
  solutionSettings.azureResources = azureResources || [];
  solutionSettings.capabilities = capabilities || [];

  // fill in activeResourcePlugins
  setActivatedResourcePluginsV2(projectSettings);
  return ok(Void);
}

export function checkWetherProvisionSucceeded(config: Json): boolean {
  return config[GLOBAL_CONFIG] && config[GLOBAL_CONFIG][SOLUTION_PROVISION_SUCCEEDED];
}

export function getPluginAndContextArray(
  ctx: SolutionContext,
  selectedPlugins: Plugin[]
): PluginsWithContext[] {
  return selectedPlugins.map((plugin) => [plugin, getPluginContext(ctx, plugin.name)]);
}
