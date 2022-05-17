// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  AzureSolutionSettings,
  FxError,
  Inputs,
  Json,
  ok,
  Platform,
  Result,
  TokenProvider,
  v2,
  Void,
} from "@microsoft/teamsfx-api";
import {
  Context,
  DeepReadonly,
  DeploymentInputs,
  EnvInfoV2,
  ProvisionInputs,
  ResourceTemplate,
} from "@microsoft/teamsfx-api/build/v2";
import { scaffold } from "./scaffold";
import * as utils from "../utils/common";
import { HostTypeTriggerOptions } from "../question";
import path from "path";
import { AzureHostingFactory } from "../../../../common/azure-hosting/hostingFactory";
import {
  AzureSolutionQuestionNames,
  BotOptionItem,
  BotScenario,
  MessageExtensionItem,
} from "../../../solution/fx-solution/question";
import {
  QuestionNames,
  TemplateProjectsConstants,
  TemplateProjectsScenarios,
  TriggerTemplateScenarioMappings,
} from "../constants";
import {
  BotTrigger,
  PluginBot,
  CommonStrings,
  Commands,
  HostType,
  HostTypes,
} from "../resources/strings";
import { CodeTemplateInfo } from "./interface/codeTemplateInfo";
import { CommandExecutionError } from "../errors";
import { BicepConfigs, ServiceType } from "../../../../common/azure-hosting/interfaces";
import { mergeTemplates } from "../../../../common/azure-hosting/utils";
import { getActivatedV2ResourcePlugins } from "../../../solution/fx-solution/ResourcePluginContainer";
import { NamedArmResourcePluginAdaptor } from "../../../solution/fx-solution/v2/adaptor";
import { ResourcePlugins } from "../../../../common/constants";
import { hostServiceTypeMapping, languageMapping } from "./constants";

export class TeamsBotV2Impl {
  async scaffoldSourceCode(ctx: Context, inputs: Inputs): Promise<Result<Void, FxError>> {
    let workingPath = inputs.projectPath ?? "";
    if (inputs.platform !== Platform.VS) {
      workingPath = path.join(workingPath, "bot");
    }
    const hostType = this.resolveHostType(inputs);
    utils.checkAndSavePluginSettingV2(ctx, PluginBot.HOST_TYPE, hostType);

    const templateInfos = this.getTemplateInfos(ctx, inputs);
    await Promise.all(
      templateInfos.map(async (templateInfo) => {
        await scaffold(templateInfo, workingPath);
      })
    );

    return ok(Void);
  }

  private resolveHostType(inputs: Inputs): HostType {
    const hostTypeTriggers = inputs[QuestionNames.BOT_HOST_TYPE_TRIGGER];
    let hostType;
    if (Array.isArray(hostTypeTriggers)) {
      const hostTypes = hostTypeTriggers.map(
        (item) => HostTypeTriggerOptions.find((option) => option.id === item)?.hostType
      );
      hostType = hostTypes ? hostTypes[0] : undefined;
    }
    return hostType ? hostType : HostTypes.APP_SERVICE;
  }

  async generateResourceTemplate(
    ctx: Context,
    inputs: Inputs
  ): Promise<Result<ResourceTemplate, FxError>> {
    const plugins = getActivatedV2ResourcePlugins(ctx.projectSetting).map(
      (p) => new NamedArmResourcePluginAdaptor(p)
    );
    const bicepConfigs = this.getBicepConfigs(ctx, inputs);
    const bicepContext = {
      plugins: plugins.map((obj) => obj.name),
      configs: bicepConfigs,
    };

    const serviceTypes = [this.resolveServiceType(ctx), ServiceType.BotService];
    const templates = await Promise.all(
      serviceTypes.map((serviceType) => {
        const hosting = AzureHostingFactory.createHosting(serviceType);
        return hosting.generateBicep(bicepContext, ResourcePlugins.Bot);
      })
    );
    const result = mergeTemplates(templates);

    return ok({ kind: "bicep", template: result });
  }

  async updateResourceTemplate(
    ctx: Context,
    inputs: Inputs
  ): Promise<Result<ResourceTemplate, FxError>> {
    const plugins = getActivatedV2ResourcePlugins(ctx.projectSetting).map(
      (p) => new NamedArmResourcePluginAdaptor(p)
    );
    const bicepConfigs = this.getBicepConfigs(ctx, inputs);
    const bicepContext = {
      plugins: plugins.map((obj) => obj.name),
      configs: bicepConfigs,
    };

    const serviceTypes = [this.resolveServiceType(ctx), ServiceType.BotService];
    const templates = await Promise.all(
      serviceTypes.map((serviceType) => {
        const hosting = AzureHostingFactory.createHosting(serviceType);
        return hosting.updateBicep(bicepContext, ResourcePlugins.Bot);
      })
    );
    const result = mergeTemplates(templates);

    return ok({ kind: "bicep", template: result });
  }

  async configureResource(
    ctx: Context,
    inputs: ProvisionInputs,
    envInfo: v2.EnvInfoV2,
    tokenProvider: TokenProvider
  ): Promise<Result<Void, FxError>> {
    return ok(Void);
  }

  async deploy(
    ctx: Context,
    inputs: DeploymentInputs,
    envInfo: DeepReadonly<v2.EnvInfoV2>,
    tokenProvider: TokenProvider
  ): Promise<Result<Void, FxError>> {
    const packDir = await this.localBuild(ctx, inputs);
    // TODO: zip packDir and upload to Azure Web App or Azure Function
    return ok(Void);
  }

  async provisionLocalResource(
    ctx: Context,
    inputs: Inputs,
    localSettings: Json,
    tokenProvider: TokenProvider,
    envInfo?: EnvInfoV2
  ): Promise<Result<Void, FxError>> {
    return ok(Void);
  }

  async configureLocalResource(
    ctx: Context,
    inputs: Inputs,
    localSettings: Json,
    tokenProvider: TokenProvider,
    envInfo?: v2.EnvInfoV2 | undefined
  ): Promise<Result<Void, FxError>> {
    return ok(Void);
  }

  private getTemplateInfos(ctx: Context, inputs: Inputs): CodeTemplateInfo[] {
    const lang = this.resolveProgrammingLanguage(ctx);
    const templateInfos: CodeTemplateInfo[] = [];
    const botGroupScenarios = new Set<string>();
    const msgextGroupScenarios = new Set<string>();
    const solutionSettings = ctx.projectSetting.solutionSettings as AzureSolutionSettings;
    const capabilities = solutionSettings.capabilities;
    capabilities.map((capability: string) => {
      switch (capability) {
        case MessageExtensionItem.id:
          this.resolveScenariosForMessageExtension(
            ctx,
            inputs,
            botGroupScenarios,
            msgextGroupScenarios
          );
          break;
        case BotOptionItem.id:
          this.resolveScenariosForBot(ctx, inputs, botGroupScenarios, msgextGroupScenarios);
          break;
      }
    });
    botGroupScenarios.forEach((scenario) => {
      templateInfos.push({
        group: TemplateProjectsConstants.GROUP_NAME_BOT,
        language: lang,
        scenario: scenario,
        variables: {},
      });
    });
    msgextGroupScenarios.forEach((scenario) => {
      templateInfos.push({
        group: TemplateProjectsConstants.GROUP_NAME_MSGEXT,
        language: lang,
        scenario: scenario,
        variables: {},
      });
    });
    return templateInfos;
  }

  private getBicepConfigs(ctx: Context, inputs: Inputs): BicepConfigs {
    const bicepConfigs: BicepConfigs = [];
    const lang = this.resolveProgrammingLanguage(ctx);
    bicepConfigs.push(languageMapping[lang]);
    bicepConfigs.push("running-on-azure");
    return bicepConfigs;
  }

  private async localBuild(ctx: Context, inputs: Inputs): Promise<string> {
    // Return the folder path to be zipped and uploaded

    const lang = this.resolveProgrammingLanguage(ctx);
    const packDir = path.join(inputs.projectPath!, CommonStrings.BOT_WORKING_DIR_NAME);
    if (lang === "ts") {
      //Typescript needs tsc build before deploy because of windows app server. other languages don't need it.
      try {
        await utils.execute("npm install", packDir);
        await utils.execute("npm run build", packDir);
        return packDir;
      } catch (e) {
        throw new CommandExecutionError(`${Commands.NPM_INSTALL},${Commands.NPM_BUILD}`, e);
      }
    }

    if (lang === "js") {
      try {
        // fail to npm install @microsoft/teamsfx on azure web app, so pack it locally.
        await utils.execute("npm install", packDir);
        return packDir;
      } catch (e) {
        throw new CommandExecutionError(`${Commands.NPM_INSTALL}`, e);
      }
    }

    if (lang === "csharp") {
      try {
        // TODO: build csharp project
        await utils.execute("dotnet publish", packDir);
        return packDir;
      } catch (e) {
        throw new CommandExecutionError(`dotnet publish`, e);
      }
    }

    throw new Error("Invalid programming language");
  }

  private resolveServiceType(ctx: Context): ServiceType {
    const rawHostType = ctx.projectSetting?.pluginSettings?.[PluginBot.PLUGIN_NAME]?.[
      PluginBot.HOST_TYPE
    ] as string;
    return hostServiceTypeMapping[rawHostType];
  }

  private resolveProgrammingLanguage(ctx: Context): string {
    const lang = ctx.projectSetting.programmingLanguage;
    switch (lang?.toLocaleLowerCase()) {
      case "javascript":
        return "js";
      case "typescript":
        return "ts";
      case "csharp":
        return "csharp";
    }
    throw new Error("Invalid programming language");
  }

  private resolveScenariosForMessageExtension(
    ctx: Context,
    inputs: Inputs,
    botGroupScenarios: Set<string>,
    msgextGroupScenarios: Set<string>
  ): void {
    const isM365 = ctx.projectSetting?.isM365;
    if (isM365) {
      botGroupScenarios.add(TemplateProjectsScenarios.M365_SCENARIO_NAME);
    } else {
      if (inputs.platform === Platform.VS) {
        msgextGroupScenarios.add(TemplateProjectsScenarios.DEFAULT_SCENARIO_NAME);
      } else {
        botGroupScenarios.add(TemplateProjectsScenarios.DEFAULT_SCENARIO_NAME);
      }
    }
  }

  private resolveScenariosForBot(
    ctx: Context,
    inputs: Inputs,
    botGroupScenarios: Set<string>,
    msgextGroupScenarios: Set<string>
  ): void {
    const botScenarios = inputs?.[AzureSolutionQuestionNames.Scenarios];
    if (!botScenarios) {
      botGroupScenarios.add(TemplateProjectsScenarios.DEFAULT_SCENARIO_NAME);
    } else {
      botScenarios.map((scenario: string) => {
        switch (scenario) {
          case BotScenario.CommandAndResponseBot:
            botGroupScenarios.add(TemplateProjectsScenarios.COMMAND_AND_RESPONSE_SCENARIO_NAME);
            break;
          case BotScenario.NotificationBot:
            const hostType = this.resolveHostType(inputs);
            if (hostType === HostTypes.AZURE_FUNCTIONS) {
              botGroupScenarios.add(
                TemplateProjectsScenarios.NOTIFICATION_FUNCTION_BASE_SCENARIO_NAME
              );
              const triggers = this.resolveTriggers(inputs);
              triggers.map((trigger) =>
                botGroupScenarios.add(TriggerTemplateScenarioMappings[trigger])
              );
            }
            if (hostType === HostTypes.APP_SERVICE) {
              botGroupScenarios.add(TemplateProjectsScenarios.NOTIFICATION_RESTIFY_SCENARIO_NAME);
            }
            break;
        }
      });
    }
  }

  private resolveTriggers(inputs: Inputs): BotTrigger[] {
    const rawHostTypeTriggers = inputs?.[QuestionNames.BOT_HOST_TYPE_TRIGGER];
    return rawHostTypeTriggers.map((hostTypeTrigger: string) => {
      const option = HostTypeTriggerOptions.find((option) => option.id === hostTypeTrigger);
      return option?.trigger;
    });
  }
}

export default new TeamsBotV2Impl();
