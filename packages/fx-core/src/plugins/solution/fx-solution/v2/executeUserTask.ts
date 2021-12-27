import {
  v2,
  Inputs,
  FxError,
  Result,
  ok,
  err,
  returnUserError,
  Func,
  returnSystemError,
  TelemetryReporter,
  AzureSolutionSettings,
  Void,
  Platform,
  UserInteraction,
  SolutionSettings,
  TokenProvider,
  combine,
  Json,
  UserError,
  IStaticTab,
  IConfigurableTab,
  IBot,
  IComposeExtension,
} from "@microsoft/teamsfx-api";
import { getStrings, isArmSupportEnabled } from "../../../../common/tools";
import { getAzureSolutionSettings, reloadV2Plugins } from "./utils";
import {
  SolutionError,
  SolutionTelemetryComponentName,
  SolutionTelemetryEvent,
  SolutionTelemetryProperty,
  SolutionTelemetrySuccess,
  SolutionSource,
} from "../constants";
import * as util from "util";
import {
  AzureResourceApim,
  AzureResourceFunction,
  AzureResourceKeyVault,
  AzureResourceSQL,
  AzureSolutionQuestionNames,
  BotOptionItem,
  HostTypeOptionAzure,
  MessageExtensionItem,
  TabOptionItem,
} from "../question";
import { cloneDeep } from "lodash";
import { sendErrorTelemetryThenReturnError } from "../utils/util";
import { getAllV2ResourcePluginMap, ResourcePluginsV2 } from "../ResourcePluginContainer";
import { Container } from "typedi";
import { scaffoldByPlugins } from "./scaffolding";
import { generateResourceTemplateForPlugins } from "./generateResourceTemplate";
import { scaffoldLocalDebugSettings } from "../debug/scaffolding";
import { AppStudioPluginV3 } from "../../../resource/appstudio/v3";
import { BuiltInResourcePluginNames } from "../v3/constants";
export async function executeUserTask(
  ctx: v2.Context,
  inputs: Inputs,
  func: Func,
  localSettings: Json,
  envInfo: v2.EnvInfoV2,
  tokenProvider: TokenProvider
): Promise<Result<unknown, FxError>> {
  const namespace = func.namespace;
  const method = func.method;
  const array = namespace.split("/");
  if (method === "addCapability") {
    return addCapability(ctx, inputs, localSettings);
  }
  if (method === "addResource") {
    return addResource(ctx, inputs, localSettings, func, envInfo, tokenProvider);
  }
  if (namespace.includes("solution")) {
    if (method === "registerTeamsAppAndAad") {
      // not implemented for now
      return err(
        returnSystemError(
          new Error("Not implemented"),
          SolutionSource,
          SolutionError.FeatureNotSupported
        )
      );
    } else if (method === "VSpublish") {
      // VSpublish means VS calling cli to do publish. It is different than normal cli work flow
      // It's teamsfx init followed by teamsfx  publish without running provision.
      // Using executeUserTask here could bypass the fx project check.
      if (inputs.platform !== "vs") {
        return err(
          returnSystemError(
            new Error(`VS publish is not supposed to run on platform ${inputs.platform}`),
            SolutionSource,
            SolutionError.UnsupportedPlatform
          )
        );
      }
      const appStudioPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.AppStudioPlugin);
      if (appStudioPlugin.publishApplication) {
        return appStudioPlugin.publishApplication(
          ctx,
          inputs,
          envInfo,
          tokenProvider.appStudioToken
        );
      }
    } else if (method === "validateManifest") {
      const appStudioPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.AppStudioPlugin);
      if (appStudioPlugin.executeUserTask) {
        return await appStudioPlugin.executeUserTask(
          ctx,
          inputs,
          func,
          localSettings,
          envInfo,
          tokenProvider
        );
      }
    } else if (method === "buildPackage") {
      const appStudioPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.AppStudioPlugin);
      if (appStudioPlugin.executeUserTask) {
        return await appStudioPlugin.executeUserTask(
          ctx,
          inputs,
          func,
          localSettings,
          envInfo,
          tokenProvider
        );
      }
    } else if (method === "validateManifest") {
      const appStudioPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.AppStudioPlugin);
      if (appStudioPlugin.executeUserTask) {
        return appStudioPlugin.executeUserTask(
          ctx,
          inputs,
          func,
          localSettings,
          envInfo,
          tokenProvider
        );
      }
    } else if (array.length == 2) {
      const pluginName = array[1];
      const pluginMap = getAllV2ResourcePluginMap();
      const plugin = pluginMap.get(pluginName);
      if (plugin && plugin.executeUserTask) {
        return plugin.executeUserTask(ctx, inputs, func, localSettings, envInfo, tokenProvider);
      }
    }
  }

  return err(
    returnUserError(
      new Error(`executeUserTaskRouteFailed:${JSON.stringify(func)}`),
      SolutionSource,
      `executeUserTaskRouteFailed`
    )
  );
}

export function canAddCapability(
  settings: AzureSolutionSettings,
  telemetryReporter: TelemetryReporter
): Result<Void, FxError> {
  if (!(settings.hostType === HostTypeOptionAzure.id)) {
    const e = new UserError(
      SolutionError.AddCapabilityNotSupport,
      getStrings().solution.addCapability.OnlySupportAzure,
      SolutionSource
    );
    return err(
      sendErrorTelemetryThenReturnError(SolutionTelemetryEvent.AddCapability, e, telemetryReporter)
    );
  }
  return ok(Void);
}

export function canAddResource(
  settings: AzureSolutionSettings,
  telemetryReporter: TelemetryReporter
): Result<Void, FxError> {
  if (!(settings.hostType === HostTypeOptionAzure.id)) {
    const e = new UserError(
      SolutionError.AddResourceNotSupport,
      getStrings().solution.addResource.OnlySupportAzure,
      SolutionSource
    );
    return err(
      sendErrorTelemetryThenReturnError(SolutionTelemetryEvent.AddResource, e, telemetryReporter)
    );
  }
  return ok(Void);
}

export async function addCapability(
  ctx: v2.Context,
  inputs: Inputs,
  localSettings: Json
): Promise<
  Result<{ solutionSettings?: SolutionSettings; solutionConfig?: Record<string, unknown> }, FxError>
> {
  ctx.telemetryReporter.sendTelemetryEvent(SolutionTelemetryEvent.AddCapabilityStart, {
    [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
  });

  // 1. checking addable
  const solutionSettings: AzureSolutionSettings = getAzureSolutionSettings(ctx);
  const originalSettings = cloneDeep(solutionSettings);
  const canProceed = canAddCapability(solutionSettings, ctx.telemetryReporter);
  if (canProceed.isErr()) {
    return err(canProceed.error);
  }

  // 2. check answer
  const capabilitiesAnswer = inputs[AzureSolutionQuestionNames.Capabilities] as string[];
  if (!capabilitiesAnswer || capabilitiesAnswer.length === 0) {
    ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddCapability, {
      [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
      [SolutionTelemetryProperty.Success]: SolutionTelemetrySuccess.Yes,
      [SolutionTelemetryProperty.Capabilities]: [].join(";"),
    });
    return ok({});
  }

  // 3. check capability limit
  solutionSettings.capabilities = solutionSettings.capabilities || [];
  const appStudioPlugin = Container.get<AppStudioPluginV3>(BuiltInResourcePluginNames.appStudio);
  const inputsWithProjectPath = inputs as v2.InputsWithProjectPath;
  const isTabAddable = !(await appStudioPlugin.capabilityExceedLimit(
    ctx,
    inputsWithProjectPath,
    "staticTab"
  ));
  const isBotAddable = !(await appStudioPlugin.capabilityExceedLimit(
    ctx,
    inputsWithProjectPath,
    "Bot"
  ));
  const isMEAddable = !(await appStudioPlugin.capabilityExceedLimit(
    ctx,
    inputsWithProjectPath,
    "MessageExtension"
  ));
  if (
    (capabilitiesAnswer.includes(TabOptionItem.id) && !isTabAddable) ||
    (capabilitiesAnswer.includes(BotOptionItem.id) && !isBotAddable) ||
    (capabilitiesAnswer.includes(MessageExtensionItem.id) && !isMEAddable)
  ) {
    const error = new UserError(
      SolutionError.FailedToAddCapability,
      getStrings().solution.addCapability.ExceedMaxLimit,
      SolutionSource
    );
    return err(
      sendErrorTelemetryThenReturnError(
        SolutionTelemetryEvent.AddCapability,
        error,
        ctx.telemetryReporter
      )
    );
  }

  const capabilitiesToAddManifest: (
    | { name: "staticTab"; snippet?: IStaticTab }
    | { name: "configurableTab"; snippet?: IConfigurableTab }
    | { name: "Bot"; snippet?: IBot }
    | { name: "MessageExtension"; snippet?: IComposeExtension }
  )[] = [];
  const pluginNamesToScaffold: Set<string> = new Set<string>();
  const pluginNamesToArm: Set<string> = new Set<string>();
  const newCapabilitySet = new Set<string>();
  solutionSettings.capabilities.forEach((c) => newCapabilitySet.add(c));
  // 4. check Tab
  if (capabilitiesAnswer.includes(TabOptionItem.id)) {
    const firstAdd = solutionSettings.capabilities.includes(TabOptionItem.id) ? false : true;
    if (inputs.platform === Platform.VS) {
      pluginNamesToScaffold.add(ResourcePluginsV2.FrontendPlugin);
      if (firstAdd) {
        pluginNamesToArm.add(ResourcePluginsV2.FrontendPlugin);
      }
    } else {
      if (firstAdd) {
        pluginNamesToScaffold.add(ResourcePluginsV2.FrontendPlugin);
        pluginNamesToArm.add(ResourcePluginsV2.FrontendPlugin);
      }
    }
    capabilitiesToAddManifest.push({ name: "staticTab" });
    newCapabilitySet.add(TabOptionItem.id);
  }
  // 5. check Bot
  if (capabilitiesAnswer.includes(BotOptionItem.id)) {
    const firstAdd =
      solutionSettings.capabilities.includes(BotOptionItem.id) ||
      solutionSettings.capabilities.includes(MessageExtensionItem.id)
        ? false
        : true;
    if (inputs.platform === Platform.VS) {
      pluginNamesToScaffold.add(ResourcePluginsV2.FrontendPlugin);
      if (firstAdd) {
        pluginNamesToArm.add(ResourcePluginsV2.BotPlugin);
      }
    } else {
      if (firstAdd) {
        pluginNamesToScaffold.add(ResourcePluginsV2.BotPlugin);
        pluginNamesToArm.add(ResourcePluginsV2.BotPlugin);
      }
    }
    capabilitiesToAddManifest.push({ name: "Bot" });
    newCapabilitySet.add(BotOptionItem.id);
  }
  // 6. check MessageExtension
  if (capabilitiesAnswer.includes(MessageExtensionItem.id)) {
    const firstAdd =
      solutionSettings.capabilities.includes(BotOptionItem.id) ||
      solutionSettings.capabilities.includes(MessageExtensionItem.id)
        ? false
        : true;
    if (inputs.platform === Platform.VS) {
      pluginNamesToScaffold.add(ResourcePluginsV2.FrontendPlugin);
      if (firstAdd) {
        pluginNamesToArm.add(ResourcePluginsV2.BotPlugin);
      }
    } else {
      if (firstAdd) {
        pluginNamesToScaffold.add(ResourcePluginsV2.BotPlugin);
        pluginNamesToArm.add(ResourcePluginsV2.BotPlugin);
      }
    }
    capabilitiesToAddManifest.push({ name: "MessageExtension" });
    newCapabilitySet.add(MessageExtensionItem.id);
  }

  // 7. update solution settings
  solutionSettings.capabilities = Array.from(newCapabilitySet);
  reloadV2Plugins(solutionSettings);

  // 8. scaffold and update arm
  const pluginsToScaffold = Array.from(pluginNamesToScaffold).map((name) =>
    Container.get<v2.ResourcePlugin>(name)
  );
  const pluginsToArm = Array.from(pluginNamesToArm).map((name) =>
    Container.get<v2.ResourcePlugin>(name)
  );
  if (pluginsToScaffold.length > 0) {
    const scaffoldRes = await scaffoldCodeAndResourceTemplate(
      ctx,
      inputs,
      localSettings,
      pluginsToScaffold,
      pluginsToArm
    );
    if (scaffoldRes.isErr()) {
      ctx.projectSetting.solutionSettings = originalSettings;
      return err(
        sendErrorTelemetryThenReturnError(
          SolutionTelemetryEvent.AddCapability,
          scaffoldRes.error,
          ctx.telemetryReporter
        )
      );
    }
    const addNames = capabilitiesAnswer.map((c) => `'${c}'`).join(" and ");
    const single = capabilitiesAnswer.length === 1;
    const template =
      inputs.platform === Platform.CLI
        ? single
          ? getStrings().solution.addCapability.AddCapabilityNoticeForCli
          : getStrings().solution.addCapability.AddCapabilitiesNoticeForCli
        : single
        ? getStrings().solution.addCapability.AddCapabilityNotice
        : getStrings().solution.addCapability.AddCapabilitiesNotice;
    const msg = util.format(template, addNames);
    ctx.userInteraction.showMessage("info", msg, false);

    ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddCapability, {
      [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
      [SolutionTelemetryProperty.Success]: SolutionTelemetrySuccess.Yes,
      [SolutionTelemetryProperty.Capabilities]: capabilitiesAnswer.join(";"),
    });
  }
  // 4. update manifest
  if (capabilitiesToAddManifest.length > 0) {
    await appStudioPlugin.addCapabilities(ctx, inputsWithProjectPath, capabilitiesToAddManifest);
  }
  return ok({
    solutionSettings: solutionSettings,
    solutionConfig: { provisionSucceeded: false },
  });
}

export function showUpdateArmTemplateNotice(ui?: UserInteraction) {
  const msg: string = util.format(getStrings().solution.UpdateArmTemplateNotice);
  ui?.showMessage("info", msg, false);
}

async function scaffoldCodeAndResourceTemplate(
  ctx: v2.Context,
  inputs: Inputs,
  localSettings: Json,
  pluginsToScaffold: v2.ResourcePlugin[],
  pluginsToDoArm?: v2.ResourcePlugin[]
): Promise<Result<unknown, FxError>> {
  const result = await scaffoldByPlugins(ctx, inputs, localSettings, pluginsToScaffold);
  if (result.isErr()) {
    return result;
  }
  const scaffoldLocalDebugSettingsResult = await scaffoldLocalDebugSettings(
    ctx,
    inputs,
    localSettings
  );
  if (scaffoldLocalDebugSettingsResult.isErr()) {
    return scaffoldLocalDebugSettingsResult;
  }
  const pluginsToUpdateArm = pluginsToDoArm ? pluginsToDoArm : pluginsToScaffold;
  if (pluginsToUpdateArm.length > 0) {
    return generateResourceTemplateForPlugins(ctx, inputs, pluginsToUpdateArm);
  }
  return ok(undefined);
}

export async function addResource(
  ctx: v2.Context,
  inputs: Inputs,
  localSettings: Json,
  func: Func,
  envInfo: v2.EnvInfoV2,
  tokenProvider: TokenProvider
): Promise<Result<unknown, FxError>> {
  ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddResourceStart, {
    [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
  });

  // 1. checking addable
  const solutionSettings: AzureSolutionSettings = getAzureSolutionSettings(ctx);
  const originalSettings = cloneDeep(solutionSettings);
  const canProceed = canAddResource(solutionSettings, ctx.telemetryReporter);
  if (canProceed.isErr()) {
    return err(canProceed.error);
  }

  // 2. check answer
  const addResourcesAnswer = inputs[AzureSolutionQuestionNames.AddResources] as string[];
  if (!addResourcesAnswer || addResourcesAnswer.length === 0) {
    ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddResource, {
      [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
      [SolutionTelemetryProperty.Success]: SolutionTelemetrySuccess.Yes,
      [SolutionTelemetryProperty.Resources]: [].join(";"),
    });
    return ok({});
  }

  const alreadyHaveFunction = solutionSettings.azureResources.includes(AzureResourceFunction.id);
  const alreadyHaveApim = solutionSettings.azureResources.includes(AzureResourceApim.id);
  const alreadyHaveKeyVault = solutionSettings.azureResources.includes(AzureResourceKeyVault.id);
  const addSQL = addResourcesAnswer.includes(AzureResourceSQL.id);
  const addApim = addResourcesAnswer.includes(AzureResourceApim.id);
  const addKeyVault = addResourcesAnswer.includes(AzureResourceKeyVault.id);
  const addFunc =
    addResourcesAnswer.includes(AzureResourceFunction.id) || (addApim && !alreadyHaveFunction);

  // 3. check APIM and KeyVault addable
  if ((alreadyHaveApim && addApim) || (alreadyHaveKeyVault && addKeyVault)) {
    const e = new UserError(
      new Error("APIM/KeyVault is already added."),
      SolutionSource,
      SolutionError.AddResourceNotSupport
    );
    return err(
      sendErrorTelemetryThenReturnError(
        SolutionTelemetryEvent.AddResource,
        e,
        ctx.telemetryReporter
      )
    );
  }

  const newResourceSet = new Set<string>();
  solutionSettings.azureResources.forEach((r) => newResourceSet.add(r));
  const addedResources: string[] = [];
  const pluginsToScaffold: v2.ResourcePlugin[] = [];
  const pluginsToDoArm: v2.ResourcePlugin[] = [];
  let scaffoldApim = false;
  // 4. check Function
  if (addFunc) {
    const functionPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.FunctionPlugin);
    pluginsToScaffold.push(functionPlugin);
    if (!alreadyHaveFunction) {
      pluginsToDoArm.push(functionPlugin);
    }
    addedResources.push(AzureResourceFunction.id);
  }
  // 5. check SQL
  if (addSQL) {
    const sqlPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.SqlPlugin);
    const identityPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.IdentityPlugin);
    pluginsToDoArm.push(sqlPlugin);
    if (!solutionSettings.activeResourcePlugins.includes(identityPlugin.name)) {
      // add identity for first time
      pluginsToDoArm.push(identityPlugin);
    }
    addedResources.push(AzureResourceSQL.id);
  }
  // 6. check APIM
  const apimPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.ApimPlugin);
  if (addApim) {
    // We don't add apimPlugin into pluginsToScaffold because
    // apim plugin needs to modify config output during scaffolding,
    // which is not supported by the scaffoldSourceCode API.
    // The scaffolding will run later as a userTask as a work around.
    addedResources.push(AzureResourceApim.id);
    pluginsToDoArm.push(apimPlugin);
    scaffoldApim = true;
  }
  if (addKeyVault) {
    const keyVaultPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.KeyVaultPlugin);
    pluginsToDoArm.push(keyVaultPlugin);
    addedResources.push(AzureResourceKeyVault.id);
  }

  // 7. update solution settings
  addedResources.forEach((r) => newResourceSet.add(r));
  solutionSettings.azureResources = Array.from(newResourceSet);
  reloadV2Plugins(solutionSettings);

  // 8. scaffold and update arm
  if (pluginsToScaffold.length > 0 || pluginsToDoArm.length > 0) {
    let scaffoldRes = await scaffoldCodeAndResourceTemplate(
      ctx,
      inputs,
      localSettings,
      pluginsToScaffold,
      pluginsToDoArm
    );
    if (scaffoldApim) {
      if (apimPlugin && apimPlugin.executeUserTask) {
        const result = await apimPlugin.executeUserTask(
          ctx,
          inputs,
          func,
          {},
          envInfo,
          tokenProvider
        );
        if (result.isErr()) {
          scaffoldRes = combine([scaffoldRes, result]);
        }
      }
    }
    if (scaffoldRes.isErr()) {
      ctx.projectSetting.solutionSettings = originalSettings;
      return err(
        sendErrorTelemetryThenReturnError(
          SolutionTelemetryEvent.AddResource,
          scaffoldRes.error,
          ctx.telemetryReporter
        )
      );
    }
    const addNames = addedResources.map((c) => `'${c}'`).join(" and ");
    const single = addedResources.length === 1;
    const template =
      inputs.platform === Platform.CLI
        ? single
          ? getStrings().solution.addResource.AddResourceNoticeForCli
          : getStrings().solution.addResource.AddResourcesNoticeForCli
        : single
        ? getStrings().solution.addResource.AddResourceNotice
        : getStrings().solution.addResource.AddResourcesNotice;
    ctx.userInteraction.showMessage("info", util.format(template, addNames), false);
  }

  ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddResource, {
    [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
    [SolutionTelemetryProperty.Success]: SolutionTelemetrySuccess.Yes,
    [SolutionTelemetryProperty.Resources]: addResourcesAnswer.join(";"),
  });
  return ok(
    pluginsToDoArm.length > 0
      ? { solutionSettings: solutionSettings, solutionConfig: { provisionSucceeded: false } }
      : Void
  );
}

export function extractParamForRegisterTeamsAppAndAad(
  answers?: Inputs
): Result<ParamForRegisterTeamsAppAndAad, FxError> {
  if (answers == undefined) {
    return err(
      returnSystemError(
        new Error("Input is undefined"),
        SolutionSource,
        SolutionError.FailedToGetParamForRegisterTeamsAppAndAad
      )
    );
  }

  const param: ParamForRegisterTeamsAppAndAad = {
    "app-name": "",
    endpoint: "",
    environment: "local",
    "root-path": "",
  };
  for (const key of Object.keys(param)) {
    const value = answers[key];
    if (value == undefined) {
      return err(
        returnSystemError(
          new Error(`${key} not found`),
          SolutionSource,
          SolutionError.FailedToGetParamForRegisterTeamsAppAndAad
        )
      );
    }
    (param as any)[key] = value;
  }

  return ok(param);
}

export type ParamForRegisterTeamsAppAndAad = {
  "app-name": string;
  environment: "local" | "remote";
  endpoint: string;
  "root-path": string;
};
