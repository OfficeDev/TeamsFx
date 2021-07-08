// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ConfigFolderName,
  FxError,
  ok,
  err,
  LogProvider,
  Platform,
  Plugin,
  PluginContext,
  QTreeNode,
  Result,
  Stage,
  TeamsAppManifest,
  DialogMsg,
  DialogType,
  MsgLevel,
  QuestionType,
  SystemError,
  UserError,
  ProjectSettings,
  Colors,
} from "@microsoft/teamsfx-api";
import { AppStudioPluginImpl } from "./plugin";
import { Constants } from "./constants";
import { IAppDefinition } from "../../solution/fx-solution/appstudio/interface";
import { AppStudioError } from "./errors";
import { AppStudioResultFactory } from "./results";
import { manuallySubmitOption, autoPublishOption } from "./questions";
import { TelemetryUtils, TelemetryEventName, TelemetryPropertyKey } from "./utils/telemetry";

export class AppStudioPlugin implements Plugin {
  private appStudioPluginImpl = new AppStudioPluginImpl();

  async getQuestions(
    stage: Stage,
    ctx: PluginContext
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    const appStudioQuestions = new QTreeNode({
      type: "group",
    });

    if (stage === Stage.publish) {
      if (ctx.answers?.platform === Platform.VS) {
        const appPath = new QTreeNode({
          type: "folder",
          name: Constants.PUBLISH_PATH_QUESTION,
          title: "Please select the folder contains manifest.json and icons",
          default: `${ctx.root}/.${ConfigFolderName}`,
        });
        appStudioQuestions.addChild(appPath);

        const remoteTeamsAppId = new QTreeNode({
          type: "text",
          name: Constants.REMOTE_TEAMS_APP_ID,
          title: "Please input the teams app id in App Studio",
        });
        appStudioQuestions.addChild(remoteTeamsAppId);
      } else if (ctx.answers?.platform === Platform.VSCode) {
        const buildOrPublish = new QTreeNode({
          name: Constants.BUILD_OR_PUBLISH_QUESTION,
          type: "singleSelect",
          staticOptions: [manuallySubmitOption, autoPublishOption],
          title: "Teams Toolkit: Publish to Teams",
          default: autoPublishOption.id,
        });
        appStudioQuestions.addChild(buildOrPublish);
      }
    }

    return ok(appStudioQuestions);
  }

  public async createApp(
    appDefinition: IAppDefinition,
    appStudioToken?: string,
    logProvider?: LogProvider,
    colorIconContent?: string, // base64 encoded
    outlineIconContent?: string // base64 encoded
  ): Promise<IAppDefinition | undefined> {
    return await this.appStudioPluginImpl.createApp(
      appDefinition,
      appStudioToken!,
      logProvider,
      colorIconContent,
      outlineIconContent
    );
  }

  public async updateApp(
    appDefinition: IAppDefinition,
    type: "localDebug" | "remote",
    createIfNotExist: boolean,
    teamsAppId?: string,
    appStudioToken?: string,
    logProvider?: LogProvider,
    projectRoot?: string
  ): Promise<Result<string, FxError>> {
    return await this.appStudioPluginImpl.updateApp(
      appDefinition,
      appStudioToken!,
      type,
      createIfNotExist,
      teamsAppId,
      logProvider,
      projectRoot
    );
  }

  public async createManifest(settings: ProjectSettings): Promise<TeamsAppManifest | undefined> {
    return await this.appStudioPluginImpl.createManifest(settings);
  }

  public async reloadManifestAndCheckRequiredFields(
    ctxRoot: string
  ): Promise<Result<TeamsAppManifest, FxError>> {
    return await this.appStudioPluginImpl.reloadManifestAndCheckRequiredFields(ctxRoot);
  }

  /**
   * Validate manifest string against schema
   * @param {string} manifestString - the string of manifest.json file
   * @returns {string[]} an array of errors
   */
  public async validateManifest(
    ctx: PluginContext,
    manifestString: string
  ): Promise<Result<string[], FxError>> {
    TelemetryUtils.init(ctx);
    TelemetryUtils.sendStartEvent(TelemetryEventName.validateManifest);
    const validationResult = await this.appStudioPluginImpl.validateManifest(ctx, manifestString);
    if (validationResult.length > 0) {
      const errMessage = AppStudioError.ValidationFailedError.message(validationResult);
      ctx.logProvider?.error("Manifest Validation failed!");
      ctx.ui?.showMessage("error", errMessage, false);
      const properties: { [key: string]: string } = {};
      properties[TelemetryPropertyKey.validationResult] = validationResult.join("\n");
      const validationFailed = AppStudioResultFactory.UserError(
        AppStudioError.ValidationFailedError.name,
        errMessage
      );
      TelemetryUtils.sendErrorEvent(
        TelemetryEventName.validateManifest,
        validationFailed,
        properties
      );
      return err(validationFailed);
    }
    const validationSuccess = "Manifest Validation succeed!";
    ctx.ui?.showMessage("info", validationSuccess, false);
    TelemetryUtils.sendSuccessEvent(TelemetryEventName.validateManifest);
    return ok(validationResult);
  }

  public getDevAppDefinition(
    manifest: string,
    appId: string,
    domains: string[],
    webApplicationInfoResource: string,
    ignoreIcon: boolean,
    tabEndpoint?: string,
    appName?: string,
    version?: string,
    botId?: string,
    appNameSuffix?: string
  ): [IAppDefinition, TeamsAppManifest] {
    return this.appStudioPluginImpl.getDevAppDefinition(
      manifest,
      appId,
      domains,
      webApplicationInfoResource,
      ignoreIcon,
      tabEndpoint,
      appName,
      version,
      botId,
      appNameSuffix
    );
  }

  public convertToAppDefinition(
    appManifest: TeamsAppManifest,
    ignoreIcon: boolean
  ): IAppDefinition {
    return this.appStudioPluginImpl.convertToAppDefinition(appManifest, ignoreIcon);
  }

  /**
   * Build Teams Package
   * @param {string} appDirectory - The directory contains manifest.source.json and two images
   * @returns {string} - Path of built appPackage.zip
   */
  public async buildTeamsPackage(
    ctx: PluginContext,
    appDirectory: string,
    manifestString: string
  ): Promise<Result<string, FxError>> {
    TelemetryUtils.init(ctx);
    TelemetryUtils.sendStartEvent(TelemetryEventName.buildTeamsPackage);
    try {
      const appPackagePath = await this.appStudioPluginImpl.buildTeamsAppPackage(
        ctx,
        appDirectory,
        manifestString
      );
      const builtSuccess = [
        { content: "(√)Done: ", color: Colors.BRIGHT_GREEN },
        { content: "Teams Package ", color: Colors.BRIGHT_WHITE },
        { content: appPackagePath, color: Colors.BRIGHT_MAGENTA },
        { content: " built successfully!", color: Colors.BRIGHT_WHITE },
      ];
      ctx.ui?.showMessage("info", builtSuccess, false);
      const properties: { [key: string]: string } = {};
      properties[TelemetryPropertyKey.buildOnly] = "true";
      TelemetryUtils.sendSuccessEvent(TelemetryEventName.buildTeamsPackage, properties);
      return ok(appPackagePath);
    } catch (error) {
      TelemetryUtils.sendErrorEvent(TelemetryEventName.buildTeamsPackage, error);
      return err(
        AppStudioResultFactory.SystemError(
          AppStudioError.TeamsPackageBuildError.name,
          AppStudioError.TeamsPackageBuildError.message(error)
        )
      );
    }
  }

  /**
   * Publish the app to Teams App Catalog
   * @param {PluginContext} ctx
   * @returns {string[]} - Teams App ID in Teams app catalog
   */
  public async publish(ctx: PluginContext): Promise<Result<string | undefined, FxError>> {
    TelemetryUtils.init(ctx);
    TelemetryUtils.sendStartEvent(TelemetryEventName.publish);
    if (ctx.answers?.platform === Platform.VSCode) {
      const answer = ctx.answers![Constants.BUILD_OR_PUBLISH_QUESTION] as string;
      if (answer === manuallySubmitOption.id) {
        const appDirectory = `${ctx.root}/.${ConfigFolderName}`;
        const manifestString = JSON.stringify(ctx.app);
        try {
          const appPackagePath = await this.appStudioPluginImpl.buildTeamsAppPackage(
            ctx,
            appDirectory,
            manifestString
          );
          const msg = `Successfully created ${ctx.app.name.short} app package file at ${appPackagePath}. Send this to your administrator for approval.`;
          ctx.ui?.showMessage("info", msg, false, "OK", Constants.READ_MORE).then((value) => {
            if (value.isOk() && value.value === Constants.READ_MORE) {
              ctx.ui?.openUrl(Constants.PUBLISH_GUIDE);
            }
          });
          TelemetryUtils.sendSuccessEvent(TelemetryEventName.publish);
          return ok(appPackagePath);
        } catch (error) {
          TelemetryUtils.sendErrorEvent(TelemetryEventName.publish, error);
          return err(
            AppStudioResultFactory.SystemError(
              AppStudioError.TeamsPackageBuildError.name,
              AppStudioError.TeamsPackageBuildError.message(error)
            )
          );
        }
      }
    }

    try {
      const result = await this.appStudioPluginImpl.publish(ctx);
      ctx.logProvider?.info(`Publish success!`);
      ctx.ui?.showMessage(
        "info",
        `${result.name} successfully published to the admin portal. Once approved, your app will be available for your organization.`,
        false
      );
      const properties: { [key: string]: string } = {};
      properties[TelemetryPropertyKey.updateExistingApp] = String(result.update);
      TelemetryUtils.sendSuccessEvent(TelemetryEventName.publish);
      return ok(result.id);
    } catch (error) {
      if (error instanceof SystemError || error instanceof UserError) {
        if (error.name === AppStudioError.TeamsAppPublishCancelError.name) {
          TelemetryUtils.sendSuccessEvent(TelemetryEventName.publish);
          return ok(undefined);
        }
        const innerError = error.innerError ? `innerError: ${error.innerError}` : "";
        error.message = `${error.message} ${innerError}`;
        TelemetryUtils.sendErrorEvent(TelemetryEventName.publish, error);
        return err(error);
      } else {
        const publishFailed = new SystemError(
          AppStudioError.TeamsAppPublishFailedError.name,
          error.message,
          Constants.PLUGIN_NAME,
          undefined,
          undefined,
          error
        );
        TelemetryUtils.sendErrorEvent(TelemetryEventName.publish, publishFailed);
        return err(publishFailed);
      }
    }
  }
}
