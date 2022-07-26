// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Action,
  ContextV3,
  Effect,
  FxError,
  InputsWithProjectPath,
  IProgressHandler,
  MaybePromise,
  ok,
  ProjectSettingsV3,
  ProvisionContextV3,
  Result,
  SourceCodeProvider,
} from "@microsoft/teamsfx-api";
import { merge } from "lodash";
import * as path from "path";
import "reflect-metadata";
import { Service } from "typedi";
import {
  genTemplateRenderReplaceFn,
  ScaffoldAction,
  ScaffoldActionName,
  ScaffoldContext,
  scaffoldFromTemplates,
} from "../../common/template-utils/templatesActions";
import { convertToAlphanumericOnly } from "../../common/utils";
import { CoreQuestionNames } from "../../core/question";
import { TemplateZipFallbackError } from "../../plugins/resource/bot/v3/error";
import {
  Constants,
  FrontendPathInfo,
  DependentPluginInfo,
  FrontendPluginInfo,
} from "../../plugins/resource/frontend/constants";
import { FrontendDeployment } from "../../plugins/resource/frontend/ops/deploy";
import {
  UnknownScaffoldError,
  UnzipTemplateError,
} from "../../plugins/resource/frontend/resources/errors";
import { Messages } from "../../plugins/resource/frontend/resources/messages";
import { Scenario, TemplateInfo } from "../../plugins/resource/frontend/resources/templateInfo";
import { ComponentNames } from "../constants";
import { ActionExecutionMW, getComponent } from "../workflow";
import { convertToLangKey } from "./botCode";
import { envFilePath, EnvKeys, saveEnvFile } from "../../plugins/resource/frontend/env";
import { isVSProject } from "../../common/projectSettingsHelper";
import { DotnetCommands } from "../../plugins/resource/frontend/dotnet/constants";
import { Utils } from "../../plugins/resource/frontend/utils";
import { CommandExecutionError } from "../../plugins/resource/bot/errors";
import { isAadManifestEnabled } from "../../common/tools";
import { hasAAD, hasApi } from "../../common/projectSettingsHelperV3";
import { ScaffoldProgress } from "../../plugins/resource/frontend/resources/steps";
import { Plans, ProgressMessages, ProgressTitles } from "../messages";
import { hooks } from "@feathersjs/hooks/lib";
/**
 * tab scaffold
 */
@Service("tab-code")
export class TabCodeProvider implements SourceCodeProvider {
  name = "tab-code";
  @hooks([
    ActionExecutionMW({
      componentName: "tab-code",
      enableTelemetry: true,
      telemetryComponentName: FrontendPluginInfo.PluginName,
      telemetryEventName: "scaffold",
      errorSource: FrontendPluginInfo.ShortName,
      errorIssueLink: FrontendPluginInfo.IssueLink,
      errorHelpLink: FrontendPluginInfo.HelpLink,
      enableProgressBar: true,
      progressTitle: ProgressTitles.scaffoldTab,
      progressSteps: Object.keys(ScaffoldProgress.steps).length,
    }),
  ])
  async generateNew(
    ctx: ContextV3,
    projectPath: string,
    folder: string,
    language: string,
    safeProjectName?: string,
    actionContext?: {
      progress?: IProgressHandler;
      telemetryProps?: Record<string, string>;
    }
  ): Promise<Result<string, FxError>> {
    folder = folder ?? (language === "csharp" ? "" : FrontendPathInfo.WorkingDir);
    const langKey = convertToLangKey(language);
    const workingDir = path.join(projectPath, folder);
    const hasFunction = hasApi(ctx.projectSetting);
    safeProjectName = safeProjectName ?? convertToAlphanumericOnly(ctx.projectSetting.appName);
    const variables = {
      showFunction: hasFunction.toString(),
      ProjectName: ctx.projectSetting.appName,
      SafeProjectName: safeProjectName,
    };

    const scenario = ctx.projectSetting.isM365
      ? Scenario.M365
      : isAadManifestEnabled() && !hasAAD(ctx.projectSetting)
      ? Scenario.NonSso
      : Scenario.Default;
    await actionContext?.progress?.next(ProgressMessages.scaffoldTab);
    await scaffoldFromTemplates({
      group: TemplateInfo.TemplateGroupName,
      lang: langKey,
      scenario: scenario,
      dst: workingDir,
      fileNameReplaceFn: (name: string, data: Buffer) =>
        name.replace(/ProjectName/, ctx.projectSetting.appName).replace(/\.tpl/, ""),
      fileDataReplaceFn: genTemplateRenderReplaceFn(variables),
      onActionEnd: async (action: ScaffoldAction, context: ScaffoldContext) => {
        if (action.name === ScaffoldActionName.FetchTemplatesUrlWithTag) {
          ctx.logProvider.info(Messages.getTemplateFrom(context.zipUrl ?? Constants.EmptyString));
        }
      },
      onActionError: async (action: ScaffoldAction, context: ScaffoldContext, error: Error) => {
        ctx.logProvider.info(error.toString());
        switch (action.name) {
          case ScaffoldActionName.FetchTemplatesUrlWithTag:
          case ScaffoldActionName.FetchTemplatesZipFromUrl:
            ctx.logProvider.info(Messages.FailedFetchTemplate);
            break;
          case ScaffoldActionName.FetchTemplateZipFromLocal:
            throw new TemplateZipFallbackError();
          case ScaffoldActionName.Unzip:
            throw new UnzipTemplateError();
          default:
            throw new UnknownScaffoldError();
        }
      },
    });

    return ok(folder);
  }
  generate(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): MaybePromise<Result<Action | undefined, FxError>> {
    const action: Action = {
      name: "tab-code.generate",
      type: "function",
      enableTelemetry: true,
      telemetryComponentName: FrontendPluginInfo.PluginName,
      telemetryEventName: "scaffold",
      errorSource: FrontendPluginInfo.ShortName,
      errorIssueLink: FrontendPluginInfo.IssueLink,
      errorHelpLink: FrontendPluginInfo.HelpLink,
      enableProgressBar: true,
      progressTitle: ProgressTitles.scaffoldTab,
      progressSteps: Object.keys(ScaffoldProgress.steps).length,
      plan: (context: ContextV3, inputs: InputsWithProjectPath) => {
        const language =
          inputs?.["programming-language"] ||
          context.projectSetting.programmingLanguage ||
          "javascript";
        const folder = inputs.folder || (language === "csharp" ? "" : FrontendPathInfo.WorkingDir);
        return ok([Plans.scaffold("tab", path.join(inputs.projectPath, folder))]);
      },
      execute: async (
        ctx: ContextV3,
        inputs: InputsWithProjectPath,
        progress?: IProgressHandler
      ) => {
        const projectSettings = ctx.projectSetting as ProjectSettingsV3;
        const appName = projectSettings.appName;
        const language = inputs[CoreQuestionNames.ProgrammingLanguage];
        const folder = inputs.folder ?? (language === "csharp" ? "" : FrontendPathInfo.WorkingDir);
        const langKey = convertToLangKey(language);
        const workingDir = path.join(inputs.projectPath, folder);
        const hasFunction = hasApi(ctx.projectSetting);
        const safeProjectName =
          inputs[CoreQuestionNames.SafeProjectName] ?? convertToAlphanumericOnly(appName);
        const variables = {
          showFunction: hasFunction.toString(),
          ProjectName: appName,
          SafeProjectName: safeProjectName,
        };
        const scenario = ctx.projectSetting.isM365
          ? Scenario.M365
          : isAadManifestEnabled() && !hasAAD(ctx.projectSetting)
          ? Scenario.NonSso
          : Scenario.Default;
        await progress?.next(ProgressMessages.scaffoldTab);
        await scaffoldFromTemplates({
          group: TemplateInfo.TemplateGroupName,
          lang: langKey,
          scenario: scenario,
          dst: workingDir,
          fileNameReplaceFn: (name: string, data: Buffer) =>
            name.replace(/ProjectName/, appName).replace(/\.tpl/, ""),
          fileDataReplaceFn: genTemplateRenderReplaceFn(variables),
          onActionEnd: async (action: ScaffoldAction, context: ScaffoldContext) => {
            if (action.name === ScaffoldActionName.FetchTemplatesUrlWithTag) {
              ctx.logProvider.info(
                Messages.getTemplateFrom(context.zipUrl ?? Constants.EmptyString)
              );
            }
          },
          onActionError: async (action: ScaffoldAction, context: ScaffoldContext, error: Error) => {
            ctx.logProvider.info(error.toString());
            switch (action.name) {
              case ScaffoldActionName.FetchTemplatesUrlWithTag:
              case ScaffoldActionName.FetchTemplatesZipFromUrl:
                ctx.logProvider.info(Messages.FailedFetchTemplate);
                break;
              case ScaffoldActionName.FetchTemplateZipFromLocal:
                throw new TemplateZipFallbackError();
              case ScaffoldActionName.Unzip:
                throw new UnzipTemplateError();
              default:
                throw new UnknownScaffoldError();
            }
          },
        });
        return ok([Plans.scaffold("tab", workingDir)]);
      },
    };
    return ok(action);
  }
  configure(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): MaybePromise<Result<Action | undefined, FxError>> {
    const action: Action = {
      name: "tab-code.configure",
      type: "function",
      plan: (context: ContextV3, inputs: InputsWithProjectPath) => {
        const teamsTab = getComponent(context.projectSetting, ComponentNames.TeamsTab);
        if (!teamsTab) return ok([]);
        const tabDir = teamsTab?.folder;
        if (!tabDir || !inputs.env) return ok([]);
        return ok([
          {
            type: "file",
            filePath: envFilePath(inputs.env, path.join(inputs.projectPath, tabDir)),
            operate: "create",
          },
        ]);
      },
      execute: async (
        context: ContextV3,
        inputs: InputsWithProjectPath
      ): Promise<Result<Effect[], FxError>> => {
        const teamsTab = getComponent(context.projectSetting, ComponentNames.TeamsTab);
        const tabDir = teamsTab?.folder;
        if (!tabDir || !inputs.env) return ok([]);
        const envFile = envFilePath(inputs.env, path.join(inputs.projectPath, tabDir));
        const envs = this.collectEnvs(context);
        await saveEnvFile(envFile, { teamsfxRemoteEnvs: envs, customizedRemoteEnvs: {} });

        return ok([
          {
            type: "file",
            filePath: envFile,
            operate: "create",
          },
        ]);
      },
    };
    return ok(action);
  }
  build(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): MaybePromise<Result<Action | undefined, FxError>> {
    const action: Action = {
      name: "tab-code.build",
      type: "function",
      enableProgressBar: true,
      progressTitle: ProgressTitles.buildingTab,
      progressSteps: 1,
      enableTelemetry: true,
      telemetryComponentName: "fx-resource-frontend",
      telemetryEventName: "build",
      plan: (context: ContextV3, inputs: InputsWithProjectPath) => {
        const teamsTab = getComponent(context.projectSetting, ComponentNames.TeamsTab);
        if (!teamsTab) return ok([]);
        const tabDir = teamsTab?.folder;
        if (!tabDir) return ok([]);
        return ok([Plans.buildProject(tabDir)]);
      },
      execute: async (
        context: ContextV3,
        inputs: InputsWithProjectPath,
        progress?: IProgressHandler
      ) => {
        const ctx = context as ProvisionContextV3;
        const teamsTab = getComponent(context.projectSetting, ComponentNames.TeamsTab);
        if (!teamsTab) return ok([]);
        if (teamsTab.folder == undefined) throw new Error("path not found");
        progress?.next(ProgressMessages.buildingTab);
        const tabPath = path.resolve(inputs.projectPath, teamsTab.folder);
        const artifactFolder = isVSProject(context.projectSetting)
          ? await this.doBlazorBuild(tabPath)
          : await this.doReactBuild(tabPath, ctx.envInfo.envName);
        merge(teamsTab, {
          build: true,
          artifactFolder: path.join(teamsTab.folder, artifactFolder),
        });
        return ok([Plans.buildProject(tabPath)]);
      },
    };
    return ok(action);
  }
  private collectEnvs(ctx: ContextV3): { [key: string]: string } {
    const envs: { [key: string]: string } = {};
    const addToEnvs = (key: string, value: string | undefined) => {
      // Check for both null and undefined, add to envs when value is "", 0 or false.
      if (value != null) {
        envs[key] = value;
      }
    };

    const connections = getComponent(ctx.projectSetting, ComponentNames.TeamsTab)?.connections;
    if (connections?.includes(ComponentNames.TeamsApi)) {
      const teamsApi = getComponent(ctx.projectSetting, ComponentNames.TeamsApi);
      addToEnvs(EnvKeys.FuncName, teamsApi?.functionNames[0]);
      addToEnvs(
        EnvKeys.FuncEndpoint,
        ctx.envInfo?.state?.[ComponentNames.TeamsApi]?.functionEndpoint as string
      );
    }
    if (connections?.includes(ComponentNames.AadApp)) {
      addToEnvs(EnvKeys.ClientID, ctx.envInfo?.state?.[ComponentNames.AadApp]?.clientId as string);
      addToEnvs(EnvKeys.StartLoginPage, DependentPluginInfo.StartLoginPageURL);
    }

    return envs;
  }
  private async doBlazorBuild(tabPath: string): Promise<string> {
    const command = DotnetCommands.buildRelease("win-x86");
    try {
      await Utils.execute(command, tabPath);
    } catch (e) {
      throw new CommandExecutionError(command, tabPath, e);
    }
    return path.join("bin", "Release", "net6.0", "win-x86", "publish");
  }
  private async doReactBuild(tabPath: string, envName: string): Promise<string> {
    await FrontendDeployment.doFrontendBuildV3(tabPath, envName);
    return "build";
  }
}
