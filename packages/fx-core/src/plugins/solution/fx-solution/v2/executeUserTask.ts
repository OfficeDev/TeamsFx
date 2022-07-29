import {
  AzureSolutionSettings,
  combine,
  Err,
  err,
  Func,
  FxError,
  Inputs,
  Json,
  ok,
  TelemetryReporter,
  Void,
  Platform,
  ProjectSettings,
  Result,
  SolutionSettings,
  SystemError,
  TokenProvider,
  UserError,
  UserInteraction,
  v2,
  v3,
  Stage,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import { cloneDeep } from "lodash";
import path from "path";
import { Container } from "typedi";
import * as util from "util";
import {
  BotHostTypeName,
  BotHostTypes,
  canAddSso,
  isAADEnabled,
  isAadManifestEnabled,
} from "../../../../common";
import { ResourcePlugins } from "../../../../common/constants";
import { isExistingTabApp, isVSProject } from "../../../../common/projectSettingsHelper";
import {
  InvalidInputError,
  NoCapabilityFoundError,
  OperationNotPermittedError,
} from "../../../../core/error";
import { CoreQuestionNames, validateCapabilities } from "../../../../core/question";
import { AppStudioPluginV3 } from "../../../resource/appstudio/v3";
import {
  AddSsoParameters,
  DEFAULT_PERMISSION_REQUEST,
  Language,
  PluginNames,
  SolutionError,
  SolutionSource,
  SolutionTelemetryComponentName,
  SolutionTelemetryEvent,
  SolutionTelemetryProperty,
  SolutionTelemetrySuccess,
  SOLUTION_PROVISION_SUCCEEDED,
} from "../constants";
import { scaffoldLocalDebugSettings } from "../debug/scaffolding";
import {
  ApiConnectionOptionItem,
  AzureResourceApim,
  AzureResourceFunction,
  AzureResourceKeyVault,
  AzureResourceSQL,
  AzureSolutionQuestionNames,
  BotOptionItem,
  BotScenario,
  BotSsoItem,
  CicdOptionItem,
  CommandAndResponseOptionItem,
  HostTypeOptionAzure,
  HostTypeOptionSPFx,
  MessageExtensionItem,
  NotificationOptionItem,
  SingleSignOnOptionItem,
  TabNonSsoItem,
  TabOptionItem,
  TabSPFxItem,
  TabSsoItem,
} from "../question";
import { getAllV2ResourcePluginMap, ResourcePluginsV2 } from "../ResourcePluginContainer";
import { sendErrorTelemetryThenReturnError } from "../utils/util";
import { BuiltInFeaturePluginNames } from "../v3/constants";
import { TeamsAppSolutionNameV2 } from "./constants";
import { generateResourceTemplateForPlugins } from "./generateResourceTemplate";
import { scaffoldByPlugins } from "./scaffolding";
import { getAzureSolutionSettings, setActivatedResourcePluginsV2 } from "./utils";
import { Certificate } from "crypto";
import { getLocalAppName } from "../../../resource/appstudio/utils/utils";
import { getDefaultString, getLocalizedString } from "../../../../common/localizeUtils";
import { getTemplatesFolder } from "../../../../folder";
import AdmZip from "adm-zip";
import { unzip } from "../../../../common/template-utils/templatesUtils";
import { InputsWithProjectPath } from "@microsoft/teamsfx-api/build/v2";
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
  if (method === "addFeature") {
    return addFeature(ctx, inputs, localSettings, func, envInfo, tokenProvider);
  }
  if (method === "addSso") {
    return addSso(ctx, inputs, localSettings);
  }
  if (namespace.includes("solution")) {
    if (method === "registerTeamsAppAndAad") {
      // not implemented for now
      return err(
        new SystemError(SolutionSource, SolutionError.FeatureNotSupported, "Not implemented")
      );
    } else if (method === "VSpublish") {
      // VSpublish means VS calling cli to do publish. It is different than normal cli work flow
      // It's teamsfx init followed by teamsfx  publish without running provision.
      // Using executeUserTask here could bypass the fx project check.
      if (inputs.platform !== "vs") {
        return err(
          new SystemError(
            SolutionSource,
            SolutionError.UnsupportedPlatform,
            getDefaultString("error.UnsupportedPlatformVS"),
            getLocalizedString("error.UnsupportedPlatformVS")
          )
        );
      }
      const appStudioPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.AppStudioPlugin);
      if (appStudioPlugin.publishApplication) {
        return appStudioPlugin.publishApplication(
          ctx,
          inputs,
          envInfo,
          tokenProvider.m365TokenProvider
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
    } else if (method === "buildAadManifest") {
      const aadPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.AadPlugin);
      if (aadPlugin.executeUserTask) {
        return await aadPlugin.executeUserTask(
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
    new UserError(
      SolutionSource,
      "executeUserTaskRouteFailed",
      getDefaultString("error.appstudio.executeUserTaskRouteFailed", JSON.stringify(func)),
      getLocalizedString("error.appstudio.executeUserTaskRouteFailed", JSON.stringify(func))
    )
  );
}

export function canAddCapability(
  settings: AzureSolutionSettings | undefined,
  telemetryReporter: TelemetryReporter
): Result<Void, FxError> {
  if (settings && !(settings.hostType === HostTypeOptionAzure.id)) {
    const e = new UserError(
      SolutionSource,
      SolutionError.AddCapabilityNotSupport,
      getDefaultString("core.addCapability.onlySupportAzure"),
      getLocalizedString("core.addCapability.onlySupportAzure")
    );
    return err(
      sendErrorTelemetryThenReturnError(SolutionTelemetryEvent.AddCapability, e, telemetryReporter)
    );
  }
  return ok(Void);
}

export function canAddResource(
  projectSetting: ProjectSettings,
  telemetryReporter: TelemetryReporter
): Result<Void, FxError> {
  const isVS = isVSProject(projectSetting);
  if (isVS) {
    const e = new UserError(
      SolutionSource,
      SolutionError.AddResourceNotSupport,
      getDefaultString("core.addResource.notSupportForVSProject"),
      getLocalizedString("core.addResource.notSupportForVSProject")
    );
    return err(
      sendErrorTelemetryThenReturnError(SolutionTelemetryEvent.AddResource, e, telemetryReporter)
    );
  }
  if (isExistingTabApp(projectSetting)) {
    return err(new NoCapabilityFoundError(Stage.addResource));
  }
  const solutionSettings = projectSetting.solutionSettings as AzureSolutionSettings;
  if (!(solutionSettings.hostType === HostTypeOptionAzure.id)) {
    const e = new UserError(
      SolutionSource,
      SolutionError.AddResourceNotSupport,
      getDefaultString("core.addResource.onlySupportAzure"),
      getLocalizedString("core.addResource.onlySupportAzure")
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

  // 0. set programming language if it is empty
  const programmingLanguageInputs = inputs[CoreQuestionNames.ProgrammingLanguage];
  if (!ctx.projectSetting.programmingLanguage && programmingLanguageInputs) {
    ctx.projectSetting.programmingLanguage = programmingLanguageInputs;
  }

  // 1. checking addable
  let solutionSettings = getAzureSolutionSettings(ctx);
  let isMiniApp = false;
  if (!solutionSettings) {
    // pure existing app
    solutionSettings = {
      name: TeamsAppSolutionNameV2,
      version: "1.0.0",
      hostType: "Azure",
      capabilities: [],
      azureResources: [],
      activeResourcePlugins: [],
    };
    ctx.projectSetting.solutionSettings = solutionSettings;
    if (!isAadManifestEnabled()) {
      //aad need this file
      await fs.writeJSON(`${inputs.projectPath}/permissions.json`, DEFAULT_PERMISSION_REQUEST, {
        spaces: 4,
      });
    }
    isMiniApp = true;
  }
  const originalSettings = cloneDeep(solutionSettings);
  const inputsNew: InputsWithProjectPath = {
    ...inputs,
    projectPath: inputs.projectPath!,
    existingResources: originalSettings.activeResourcePlugins,
    existingCapabilities: originalSettings.capabilities,
  };
  const canProceed = canAddCapability(solutionSettings, ctx.telemetryReporter);
  if (canProceed.isErr()) {
    return err(canProceed.error);
  }

  // 2. check answer
  let capabilitiesAnswer = inputs[AzureSolutionQuestionNames.Capabilities] as string[];
  if (!capabilitiesAnswer || capabilitiesAnswer.length === 0) {
    ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddCapability, {
      [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
      [SolutionTelemetryProperty.Success]: SolutionTelemetrySuccess.Yes,
      [SolutionTelemetryProperty.Capabilities]: [].join(";"),
    });
    return ok({});
  }
  const validateRes = validateCapabilities(capabilitiesAnswer);
  if (validateRes) {
    return err(InvalidInputError(validateRes));
  }
  // add spfx tab is not permitted for non-mini app
  if (!isMiniApp && capabilitiesAnswer.includes(TabSPFxItem.id)) {
    return err(InvalidInputError(getLocalizedString("core.capability.validation.spfx")));
  }

  // normalize capability answer
  const scenarios: BotScenario[] = [];
  const notificationIndex = capabilitiesAnswer.indexOf(NotificationOptionItem.id);
  if (notificationIndex !== -1) {
    capabilitiesAnswer[notificationIndex] = BotOptionItem.id;
    scenarios.push(BotScenario.NotificationBot);
  }
  const commandAndResponseIndex = capabilitiesAnswer.indexOf(CommandAndResponseOptionItem.id);
  if (commandAndResponseIndex !== -1) {
    capabilitiesAnswer[commandAndResponseIndex] = BotOptionItem.id;
    scenarios.push(BotScenario.CommandAndResponseBot);
  }
  inputsNew[AzureSolutionQuestionNames.Scenarios] = scenarios;
  capabilitiesAnswer = [...new Set(capabilitiesAnswer)];

  // 3. check capability limit
  const alreadyHasTab = solutionSettings.capabilities.includes(TabOptionItem.id);
  const alreadyHasBot = solutionSettings.capabilities.includes(BotOptionItem.id);
  const alreadyHasME = solutionSettings.capabilities.includes(MessageExtensionItem.id);
  const alreadyHasTabSso =
    isAadManifestEnabled() && solutionSettings.capabilities.includes(TabSsoItem.id);
  const toAddTab = capabilitiesAnswer.includes(TabOptionItem.id);
  const toAddBot = capabilitiesAnswer.includes(BotOptionItem.id);
  const toAddME = capabilitiesAnswer.includes(MessageExtensionItem.id);
  const toAddTabNonSso = isAadManifestEnabled() && capabilitiesAnswer.includes(TabNonSsoItem.id);
  const toAddSpfx = false; //capabilitiesAnswer.includes(TabSPFxItem.id);
  if (isAadManifestEnabled()) {
    if (alreadyHasTabSso && toAddTabNonSso) {
      const e = new SystemError(
        SolutionSource,
        SolutionError.InvalidInput,
        getLocalizedString("core.addSsoFiles.canNotAddNonSsoTabWhenSsoEnabled")
      );
      return err(e);
    }

    if (alreadyHasTab && !alreadyHasTabSso && toAddTab) {
      const e = new SystemError(
        SolutionSource,
        SolutionError.InvalidInput,
        getLocalizedString("core.addSsoFiles.canNotAddTabWhenSsoNotEnabled")
      );
      return err(e);
    }
  }

  const appStudioPlugin = Container.get<AppStudioPluginV3>(BuiltInFeaturePluginNames.appStudio);
  const inputsWithProjectPath = inputs as v2.InputsWithProjectPath;
  const tabExceedRes = await appStudioPlugin.capabilityExceedLimit(
    ctx,
    inputs as v2.InputsWithProjectPath,
    "staticTab"
  );
  if (tabExceedRes.isErr()) {
    return err(tabExceedRes.error);
  }
  const isTabAddable = !tabExceedRes.value;
  const botExceedRes = await appStudioPlugin.capabilityExceedLimit(
    ctx,
    inputs as v2.InputsWithProjectPath,
    "Bot"
  );
  if (botExceedRes.isErr()) {
    return err(botExceedRes.error);
  }
  const isBotAddable = !botExceedRes.value;
  const meExceedRes = await appStudioPlugin.capabilityExceedLimit(
    ctx,
    inputs as v2.InputsWithProjectPath,
    "MessageExtension"
  );
  if (meExceedRes.isErr()) {
    return err(meExceedRes.error);
  }
  const isMEAddable = !meExceedRes.value;
  if (
    ((toAddTab || toAddTabNonSso) && !isTabAddable) ||
    (toAddBot && !isBotAddable) ||
    (toAddME && !isMEAddable)
  ) {
    const error = new UserError(
      SolutionSource,
      SolutionError.FailedToAddCapability,
      getDefaultString("core.addCapability.exceedMaxLimit"),
      getLocalizedString("core.addCapability.exceedMaxLimit")
    );
    return err(
      sendErrorTelemetryThenReturnError(
        SolutionTelemetryEvent.AddCapability,
        error,
        ctx.telemetryReporter
      )
    );
  }

  const capabilitiesToAddManifest: v3.ManifestCapability[] = [];
  const pluginNamesToScaffold: Set<string> = new Set<string>();
  const pluginNamesToArm: Set<string> = new Set<string>();
  const newCapabilitySet = new Set<string>();
  solutionSettings.capabilities.forEach((c) => newCapabilitySet.add(c));
  const vsProject = isVSProject(ctx.projectSetting);

  // check SPFx
  if (toAddSpfx) {
    pluginNamesToScaffold.add(ResourcePluginsV2.SpfxPlugin);
    capabilitiesToAddManifest.push({ name: "staticTab" });
    capabilitiesToAddManifest.push({ name: "configurableTab" });
    newCapabilitySet.add(TabSPFxItem.id);
    solutionSettings.hostType = HostTypeOptionSPFx.id;
  } else {
    if (!originalSettings.activeResourcePlugins.includes(BuiltInFeaturePluginNames.identity)) {
      pluginNamesToArm.add(ResourcePluginsV2.IdentityPlugin);
    }
    if (
      !isAadManifestEnabled() &&
      !originalSettings.activeResourcePlugins.includes(BuiltInFeaturePluginNames.aad)
    ) {
      pluginNamesToArm.add(ResourcePluginsV2.AadPlugin);
    }

    // 4. check Tab
    if (toAddTab || toAddTabNonSso) {
      if (vsProject) {
        pluginNamesToScaffold.add(ResourcePluginsV2.FrontendPlugin);
        if (!alreadyHasTab) {
          pluginNamesToArm.add(ResourcePluginsV2.FrontendPlugin);
        }
      } else {
        if (!alreadyHasTab) {
          pluginNamesToScaffold.add(ResourcePluginsV2.FrontendPlugin);
          pluginNamesToArm.add(ResourcePluginsV2.FrontendPlugin);
        }
      }
      capabilitiesToAddManifest.push({ name: "staticTab" });
      newCapabilitySet.add(TabOptionItem.id);

      if (toAddTab && !alreadyHasTabSso) {
        newCapabilitySet.add(TabSsoItem.id);
        pluginNamesToScaffold.add(ResourcePluginsV2.AadPlugin);
        pluginNamesToArm.add(ResourcePluginsV2.AadPlugin);

        // Add webapplicationInfo in teams app manifest
        const appStudioPlugin = Container.get<AppStudioPluginV3>(
          BuiltInFeaturePluginNames.appStudio
        );
        await appStudioPlugin.addCapabilities(ctx, inputs as v2.InputsWithProjectPath, [
          { name: "WebApplicationInfo" },
        ]);
      }
    }

    // 5. check Bot
    if (toAddBot) {
      if (vsProject) {
        pluginNamesToScaffold.add(ResourcePluginsV2.FrontendPlugin);
        if (!alreadyHasBot && !alreadyHasME) {
          pluginNamesToArm.add(ResourcePluginsV2.BotPlugin);
        }
      } else {
        if (!alreadyHasBot && !alreadyHasME) {
          pluginNamesToScaffold.add(ResourcePluginsV2.BotPlugin);
          pluginNamesToArm.add(ResourcePluginsV2.BotPlugin);
        }
      }
      capabilitiesToAddManifest.push({ name: "Bot" });
      newCapabilitySet.add(BotOptionItem.id);
    }

    // 6. check MessageExtension
    if (toAddME) {
      if (vsProject) {
        pluginNamesToScaffold.add(ResourcePluginsV2.FrontendPlugin);
        if (!alreadyHasBot && !alreadyHasME) {
          pluginNamesToArm.add(ResourcePluginsV2.BotPlugin);
        }
      } else {
        if (!alreadyHasBot && !alreadyHasME) {
          pluginNamesToScaffold.add(ResourcePluginsV2.BotPlugin);
          pluginNamesToArm.add(ResourcePluginsV2.BotPlugin);
        }
      }
      capabilitiesToAddManifest.push({ name: "MessageExtension" });
      newCapabilitySet.add(MessageExtensionItem.id);
    }
  }

  // 7. update solution settings
  solutionSettings.capabilities = Array.from(newCapabilitySet);
  setActivatedResourcePluginsV2(ctx.projectSetting);

  if (
    !toAddSpfx &&
    !isAadManifestEnabled() &&
    !solutionSettings.activeResourcePlugins.includes(BuiltInFeaturePluginNames.aad)
  ) {
    solutionSettings.activeResourcePlugins.push(BuiltInFeaturePluginNames.aad);
  }

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
      inputsNew,
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
  }
  // 4. update manifest
  if (capabilitiesToAddManifest.length > 0) {
    await appStudioPlugin.addCapabilities(ctx, inputsNew, capabilitiesToAddManifest);
  }
  if (capabilitiesAnswer.length > 0) {
    const addNames = capabilitiesAnswer.map((c) => `'${c}'`).join(" and ");
    const single = capabilitiesAnswer.length === 1;
    const template =
      inputs.platform === Platform.CLI
        ? single
          ? getLocalizedString("core.addCapability.addCapabilityNoticeForCli")
          : getLocalizedString("core.addCapability.addCapabilitiesNoticeForCli")
        : single
        ? getLocalizedString("core.addCapability.addCapabilityNotice")
        : getLocalizedString("core.addCapability.addCapabilitiesNotice");
    const msg = util.format(template, addNames);
    ctx.userInteraction.showMessage("info", msg, false);
    ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddCapability, {
      [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
      [SolutionTelemetryProperty.Success]: SolutionTelemetrySuccess.Yes,
      [SolutionTelemetryProperty.Capabilities]: capabilitiesAnswer.join(";"),
    });
  }
  return ok({
    solutionSettings: solutionSettings,
    solutionConfig: { provisionSucceeded: false },
  });
}

export function showUpdateArmTemplateNotice(ui?: UserInteraction) {
  const msg: string = getLocalizedString("core.updateArmTemplate.successNotice");
  ui?.showMessage("info", msg, false);
}

async function scaffoldCodeAndResourceTemplate(
  ctx: v2.Context,
  inputs: Inputs,
  localSettings: Json,
  pluginsToScaffold: v2.ResourcePlugin[],
  pluginsToDoArm?: v2.ResourcePlugin[],
  concurrent = true
): Promise<Result<unknown, FxError>> {
  const result = await scaffoldByPlugins(ctx, inputs, localSettings, pluginsToScaffold, concurrent);
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
  const solutionSettings = getAzureSolutionSettings(ctx);
  if (!solutionSettings) {
    return err(new OperationNotPermittedError("addResource"));
  }
  const originalSettings = cloneDeep(solutionSettings);
  const inputsNew: v2.InputsWithProjectPath & { existingResources: string[] } = {
    ...inputs,
    projectPath: inputs.projectPath!,
    existingResources: originalSettings.activeResourcePlugins,
  };
  const canProceed = canAddResource(ctx.projectSetting, ctx.telemetryReporter);
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
    addResourcesAnswer.includes(AzureResourceFunction.id) ||
    (addApim && !alreadyHaveFunction) ||
    (addSQL && !alreadyHaveFunction);

  // 3. check APIM and KeyVault addable
  if ((alreadyHaveApim && addApim) || (alreadyHaveKeyVault && addKeyVault)) {
    const e = new UserError(
      SolutionSource,
      SolutionError.AddResourceNotSupport,
      "APIM/KeyVault is already added."
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
  let addSsoRes = {};
  // 4. check Function
  if (addFunc) {
    // AAD plugin needs to be activated when adding function.
    // Since APIM also have dependency on Function, will only add depenedency here.
    if (!isAADEnabled(solutionSettings)) {
      if (isAadManifestEnabled()) {
        const res = await addSso(ctx, inputs, localSettings);
        if (res.isErr()) {
          ctx.projectSetting.solutionSettings = originalSettings;
          return err(
            sendErrorTelemetryThenReturnError(
              SolutionTelemetryEvent.AddResource,
              res.error,
              ctx.telemetryReporter
            )
          );
        }
        addSsoRes = res.value as any;
      } else {
        solutionSettings.activeResourcePlugins?.push(PluginNames.AAD);
      }
    }
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
  setActivatedResourcePluginsV2(ctx.projectSetting);

  // 8. scaffold and update arm
  if (pluginsToScaffold.length > 0 || pluginsToDoArm.length > 0) {
    let scaffoldRes = await scaffoldCodeAndResourceTemplate(
      ctx,
      inputsNew,
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
          ? getLocalizedString("core.addResource.addResourceNoticeForCli")
          : getLocalizedString("core.addResource.addResourcesNoticeForCli")
        : single
        ? getLocalizedString("core.addResource.addResourceNotice")
        : getLocalizedString("core.addResource.addResourcesNotice");
    ctx.userInteraction.showMessage("info", util.format(template, addNames), false);
  }

  ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddResource, {
    [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
    [SolutionTelemetryProperty.Success]: SolutionTelemetrySuccess.Yes,
    [SolutionTelemetryProperty.Resources]: addResourcesAnswer.join(";"),
  });
  return ok(
    pluginsToDoArm.length > 0
      ? {
          solutionSettings: solutionSettings,
          solutionConfig: { provisionSucceeded: false },
          ...addSsoRes,
        }
      : Void
  );
}

export async function addFeature(
  ctx: v2.Context,
  inputs: Inputs,
  localSettings: Json,
  func: Func,
  envInfo: v2.EnvInfoV2,
  tokenProvider: TokenProvider
): Promise<Result<unknown, FxError>> {
  const featureAnswer = inputs[AzureSolutionQuestionNames.Features] as string;
  const capabilityAnswers = new Set([
    TabOptionItem.id,
    BotOptionItem.id,
    CommandAndResponseOptionItem.id,
    NotificationOptionItem.id,
    TabNonSsoItem.id,
    MessageExtensionItem.id,
  ]);
  const resourceAnswers = new Set([
    AzureResourceFunction.id,
    AzureResourceSQL.id,
    AzureResourceApim.id,
    AzureResourceKeyVault.id,
  ]);
  if (capabilityAnswers.has(featureAnswer)) {
    inputs[AzureSolutionQuestionNames.Capabilities] = [featureAnswer];
    return addCapability(ctx, inputs, localSettings);
  }
  const settings = ctx.projectSetting.solutionSettings as AzureSolutionSettings | undefined;
  const alreadyHaveFunction = settings?.azureResources.includes(AzureResourceFunction.id);
  if (resourceAnswers.has(featureAnswer)) {
    inputs[AzureSolutionQuestionNames.AddResources] = [featureAnswer];
    if (
      (featureAnswer === AzureResourceSQL.id || featureAnswer === AzureResourceApim.id) &&
      !alreadyHaveFunction
    ) {
      inputs[AzureSolutionQuestionNames.AddResources].push(AzureResourceFunction.id);
    }
    return addResource(ctx, inputs, localSettings, func, envInfo, tokenProvider);
  }
  if (featureAnswer === SingleSignOnOptionItem.id) {
    return addSso(ctx, inputs, localSettings);
  } else if (featureAnswer === ApiConnectionOptionItem.id) {
    const apiFunction: Func = {
      namespace: "fx-solution-azure/fx-resource-api-connector",
      method: "connectExistingApi",
      params: {},
    };
    return executeUserTask(ctx, inputs, apiFunction, localSettings, envInfo, tokenProvider);
  } else if (featureAnswer === CicdOptionItem.id) {
    const cicdFunction: Func = {
      namespace: "fx-solution-azure/fx-resource-cicd",
      method: "addCICDWorkflows",
      params: {},
    };
    return executeUserTask(ctx, inputs, cicdFunction, localSettings, envInfo, tokenProvider);
  }
  return ok({});
}

export function extractParamForRegisterTeamsAppAndAad(
  answers?: Inputs
): Result<ParamForRegisterTeamsAppAndAad, FxError> {
  if (answers == undefined) {
    return err(
      new SystemError(
        SolutionSource,
        SolutionError.FailedToGetParamForRegisterTeamsAppAndAad,
        "Input is undefined"
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
        new SystemError(
          SolutionSource,
          SolutionError.FailedToGetParamForRegisterTeamsAppAndAad,
          `${key} not found`
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

export async function addSso(
  ctx: v2.Context,
  inputs: Inputs,
  localSettings: Json
): Promise<Result<unknown, FxError>> {
  ctx.telemetryReporter.sendTelemetryEvent(SolutionTelemetryEvent.AddSsoStart, {
    [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
  });

  let solutionSettings = getAzureSolutionSettings(ctx);
  let existingApp = false;
  if (!solutionSettings) {
    // pure existing app
    solutionSettings = {
      name: TeamsAppSolutionNameV2,
      version: "1.0.0",
      hostType: "Azure",
      capabilities: [],
      azureResources: [],
      activeResourcePlugins: [],
    };
    ctx.projectSetting.solutionSettings = solutionSettings;
    existingApp = true;
  }

  // Check whether can add sso
  const canProceed = canAddSso(ctx.projectSetting, true) as Result<Void, FxError>;
  if (canProceed.isErr()) {
    return err(
      sendErrorTelemetryThenReturnError(
        SolutionTelemetryEvent.AddSso,
        canProceed.error,
        ctx.telemetryReporter
      )
    );
  }

  const needsTab =
    solutionSettings.capabilities.length == 0 ||
    (solutionSettings.capabilities.includes(TabOptionItem.id) &&
      !solutionSettings.capabilities.includes(TabSsoItem.id));
  const needsBot =
    solutionSettings.capabilities.includes(BotOptionItem.id) &&
    !(
      ctx.projectSetting.pluginSettings?.[ResourcePlugins.Bot]?.[BotHostTypeName] ===
      BotHostTypes.AzureFunctions
    ) &&
    !solutionSettings.capabilities.includes(BotSsoItem.id);

  // Update project settings
  if (!solutionSettings.activeResourcePlugins.includes(PluginNames.AAD)) {
    solutionSettings.activeResourcePlugins.push(PluginNames.AAD);
  }
  if (solutionSettings.capabilities.length == 0) {
    solutionSettings.capabilities.push(TabSsoItem.id);
  }
  if (
    solutionSettings.capabilities.includes(TabOptionItem.id) &&
    !solutionSettings.capabilities.includes(TabSsoItem.id)
  ) {
    solutionSettings.capabilities.push(TabSsoItem.id);
  }
  if (needsBot) {
    solutionSettings.capabilities.push(BotSsoItem.id);
  }

  const originalSettings = cloneDeep(solutionSettings);
  const inputsNew = {
    ...inputs,
    projectPath: inputs.projectPath!,
    existingResources: originalSettings.activeResourcePlugins,
    existingCapabilities: originalSettings.capabilities,
  };

  const createAuthFilesRes = await createAuthFiles(
    inputsNew,
    ctx,
    needsTab,
    needsBot,
    isVSProject(ctx.projectSetting)
  );
  if (createAuthFilesRes.isErr()) {
    return err(
      sendErrorTelemetryThenReturnError(
        SolutionTelemetryEvent.AddSso,
        createAuthFilesRes.error,
        ctx.telemetryReporter
      )
    );
  }

  // Scaffold aad plugin and arm template
  const scaffoldRes = await scaffoldCodeAndResourceTemplate(
    ctx,
    inputsNew,
    localSettings,
    [Container.get<v2.ResourcePlugin>(ResourcePluginsV2.AadPlugin)],
    existingApp ? [] : [Container.get<v2.ResourcePlugin>(ResourcePluginsV2.AadPlugin)]
  );
  if (scaffoldRes.isErr()) {
    ctx.projectSetting.solutionSettings = originalSettings;
    return err(
      sendErrorTelemetryThenReturnError(
        SolutionTelemetryEvent.AddSso,
        scaffoldRes.error,
        ctx.telemetryReporter
      )
    );
  }

  // Update manifest
  const appStudioPlugin = Container.get<AppStudioPluginV3>(BuiltInFeaturePluginNames.appStudio);
  await appStudioPlugin.addCapabilities(ctx, inputs as v2.InputsWithProjectPath, [
    { name: "WebApplicationInfo" },
  ]);

  if (inputs.platform == Platform.VSCode) {
    ctx.userInteraction
      .showMessage(
        "info",
        getLocalizedString("core.addSso.learnMore", AddSsoParameters.LearnMore),
        false,
        AddSsoParameters.LearnMore
      )
      .then((result) => {
        const userSelected = result.isOk() ? result.value : undefined;
        if (userSelected === AddSsoParameters.LearnMore) {
          ctx.userInteraction?.openUrl(AddSsoParameters.LearnMoreUrl);
          ctx.telemetryReporter.sendTelemetryEvent(SolutionTelemetryEvent.AddSsoReadme, {
            [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
          });
        }
      });
  } else if (inputs.platform == Platform.CLI) {
    await ctx.userInteraction.showMessage(
      "info",
      getLocalizedString("core.addSso.learnMore", AddSsoParameters.LearnMoreUrl),
      false
    );
  }

  ctx.telemetryReporter.sendTelemetryEvent(SolutionTelemetryEvent.AddSso, {
    [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
    [SolutionTelemetryProperty.Success]: SolutionTelemetrySuccess.Yes,
    [SolutionTelemetryProperty.AddTabSso]: needsTab
      ? SolutionTelemetrySuccess.Yes
      : SolutionTelemetrySuccess.No,
    [SolutionTelemetryProperty.AddBotSso]: needsBot
      ? SolutionTelemetrySuccess.Yes
      : SolutionTelemetrySuccess.No,
  });

  return ok({
    func: AddSsoParameters.AddSso,
    capabilities: [
      ...(needsTab ? [AddSsoParameters.Tab] : []),
      ...(needsBot ? [AddSsoParameters.Bot] : []),
    ],
  });
}

// TODO: use 'isVsProject' for changes in VS
export async function createAuthFiles(
  input: Inputs,
  ctx: v2.Context,
  needTab: boolean,
  needBot: boolean,
  isVsProject = false
): Promise<Result<unknown, FxError>> {
  const projectPath = input.projectPath;
  if (!projectPath) {
    const e = new SystemError(
      SolutionSource,
      SolutionError.InvalidProjectPath,
      getLocalizedString("core.addSsoFiles.emptyProjectPath")
    );
    return err(e);
  }

  const language = (ctx.projectSetting.programmingLanguage as string) ?? Language.JavaScript;
  const languageFolderResult = validateAndParseLanguage(language);
  if (languageFolderResult.isErr()) {
    return err(languageFolderResult.error);
  }
  const languageFolderName = languageFolderResult.value;

  const projectFolderExists = await fs.pathExists(projectPath!);
  if (!projectFolderExists) {
    const e = new SystemError(
      SolutionSource,
      SolutionError.InvalidProjectPath,
      getLocalizedString("core.addSsoFiles.projectPathNotExists")
    );
    return err(e);
  }

  const authFolder = path.join(projectPath!, "auth");
  const tabFolder = path.join(authFolder, AddSsoParameters.Tab);
  const botFolder = path.join(authFolder, AddSsoParameters.Bot);
  try {
    const authFolderExists = await fs.pathExists(authFolder);
    if (!authFolderExists) {
      await fs.ensureDir(authFolder);
    }

    if (needTab) {
      const tabFolderExists = await fs.pathExists(tabFolder);
      if (!tabFolderExists) {
        await fs.ensureDir(tabFolder);
      }

      const templateFolder = getTemplatesFolder();
      const tabTemplateFolder = path.join(
        templateFolder,
        AddSsoParameters.filePath,
        AddSsoParameters.Tab
      );
      if (isVsProject) {
        // TODO: add steps for VS
      } else {
        // README.md
        const readmeSourcePath = path.join(tabTemplateFolder, AddSsoParameters.Readme);
        const readmeTargetPath = path.join(tabFolder, AddSsoParameters.Readme);
        const readme = await fs.readFile(readmeSourcePath);
        fs.writeFile(readmeTargetPath, readme);

        // Sample Code
        const sampleSourceFolder = path.join(tabTemplateFolder, languageFolderName);
        const sampleZip = new AdmZip();
        sampleZip.addLocalFolder(sampleSourceFolder);
        await unzip(sampleZip, tabFolder);
      }
    }

    if (needBot) {
      const botFolderExists = await fs.pathExists(botFolder);
      if (!botFolderExists) {
        await fs.ensureDir(botFolder);
      }

      const templateFolder = getTemplatesFolder();
      const botTemplateFolder = path.join(
        templateFolder,
        AddSsoParameters.filePath,
        AddSsoParameters.Bot
      );
      if (isVsProject) {
        // TODO: add steps for VS
      } else {
        // README.md
        const readmeSourcePath = path.join(botTemplateFolder, AddSsoParameters.Readme);
        const readmeTargetPath = path.join(botFolder, AddSsoParameters.Readme);
        const readme = await fs.readFile(readmeSourcePath);
        fs.writeFile(readmeTargetPath, readme);

        // Sample Code
        const sampleSourceFolder = path.join(botTemplateFolder, languageFolderName);
        const sampleZip = new AdmZip();
        sampleZip.addLocalFolder(sampleSourceFolder);
        await unzip(sampleZip, botFolder);
      }
    }
  } catch (error) {
    if (needTab && (await fs.pathExists(tabFolder))) {
      await fs.remove(tabFolder);
    }
    if (needBot && (await fs.pathExists(botFolder))) {
      await fs.remove(botFolder);
    }
    const e = new SystemError(
      SolutionSource,
      SolutionError.FailedToCreateAuthFiles,
      getLocalizedString("core.addSsoFiles.FailedToCreateAuthFiles", error.message)
    );
    return err(e);
  }

  return ok(undefined);
}

export function validateAndParseLanguage(language: string): Result<string, FxError> {
  if (language.toLowerCase() == Language.TypeScript) {
    return ok("ts");
  }

  if (language.toLowerCase() == Language.JavaScript) {
    return ok("js");
  }

  if (language.toLocaleLowerCase() == Language.CSharp) {
    return ok("csharp");
  }

  const e = new SystemError(
    SolutionSource,
    SolutionError.InvalidInput,
    getLocalizedString("core.addSsoFiles.invalidLanguage")
  );
  return err(e);
}
