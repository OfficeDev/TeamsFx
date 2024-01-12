// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks";
import {
  ApiOperation,
  AppPackageFolderName,
  BuildFolderName,
  Context,
  CoreCallbackEvent,
  CreateProjectResult,
  CryptoProvider,
  Func,
  FxError,
  IQTreeNode,
  Inputs,
  InputsWithProjectPath,
  ManifestUtil,
  OpenAIPluginManifest,
  Platform,
  ResponseTemplatesFolderName,
  Result,
  Stage,
  TeamsAppInputs,
  Tools,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import { DotenvParseOutput } from "dotenv";
import fs from "fs-extra";
import * as os from "os";
import * as path from "path";
import "reflect-metadata";
import { Container } from "typedi";
import { pathToFileURL } from "url";
import { parse, parseDocument } from "yaml";
import { VSCodeExtensionCommand } from "../common/constants";
import { isApiKeyEnabled, isMultipleParametersEnabled } from "../common/featureFlags";
import { getLocalizedString } from "../common/localizeUtils";
import { LaunchHelper } from "../common/m365/launchHelper";
import { ListCollaboratorResult, PermissionsResult } from "../common/permissionInterface";
import { isValidProjectV2, isValidProjectV3 } from "../common/projectSettingsHelper";
import { ProjectTypeResult, projectTypeChecker } from "../common/projectTypeChecker";
import { SpecParser, SpecParserError } from "../common/spec-parser";
import { TelemetryEvent, fillinProjectTypeProperties } from "../common/telemetry";
import { MetadataV3, VersionSource, VersionState } from "../common/versionMetadata";
import { ILifecycle, LifecycleName } from "../component/configManager/interface";
import { YamlParser } from "../component/configManager/parser";
import {
  AadConstants,
  SPFxQuestionNames,
  SingleSignOnOptionItem,
  ViewAadAppHelpLinkV5,
} from "../component/constants";
import { coordinator } from "../component/coordinator";
import { UpdateAadAppArgs } from "../component/driver/aad/interface/updateAadAppArgs";
import { UpdateAadAppDriver } from "../component/driver/aad/update";
import { buildAadManifest } from "../component/driver/aad/utility/buildAadManifest";
import { AddWebPartDriver } from "../component/driver/add/addWebPart";
import { AddWebPartArgs } from "../component/driver/add/interface/AddWebPartArgs";
import "../component/driver/index";
import { DriverContext } from "../component/driver/interface/commonArgs";
import "../component/driver/script/scriptDriver";
import { updateManifestV3 } from "../component/driver/teamsApp/appStudio";
import { CreateAppPackageDriver } from "../component/driver/teamsApp/createAppPackage";
import { CreateAppPackageArgs } from "../component/driver/teamsApp/interfaces/CreateAppPackageArgs";
import { ValidateAppPackageArgs } from "../component/driver/teamsApp/interfaces/ValidateAppPackageArgs";
import { ValidateManifestArgs } from "../component/driver/teamsApp/interfaces/ValidateManifestArgs";
import { teamsappMgr } from "../component/driver/teamsApp/teamsappMgr";
import { manifestUtils } from "../component/driver/teamsApp/utils/ManifestUtils";
import {
  containsUnsupportedFeature,
  getFeaturesFromAppDefinition,
} from "../component/driver/teamsApp/utils/utils";
import { ValidateManifestDriver } from "../component/driver/teamsApp/validate";
import { ValidateAppPackageDriver } from "../component/driver/teamsApp/validateAppPackage";
import { SSO } from "../component/feature/sso";
import {
  ErrorResult,
  OpenAIPluginManifestHelper,
  convertSpecParserErrorToFxError,
  generateScaffoldingSummary,
  listOperations,
  specParserGenerateResultAllSuccessTelemetryProperty,
  specParserGenerateResultTelemetryEvent,
  specParserGenerateResultWarningsTelemetryProperty,
} from "../component/generator/copilotPlugin/helper";
import { EnvLoaderMW, EnvWriterMW } from "../component/middleware/envMW";
import { QuestionMW } from "../component/middleware/questionMW";
import { createContextV3, createDriverContext } from "../component/utils";
import { expandEnvironmentVariable } from "../component/utils/common";
import { envUtil } from "../component/utils/envUtil";
import { metadataUtil } from "../component/utils/metadataUtil";
import { pathUtils } from "../component/utils/pathUtils";
import { settingsUtil } from "../component/utils/settingsUtil";
import {
  FileNotFoundError,
  InjectAPIKeyActionFailedError,
  InvalidProjectError,
  MultipleAuthError,
  MultipleServerError,
  assembleError,
} from "../error/common";
import { NoNeedUpgradeError } from "../error/upgrade";
import { YamlFieldMissingError } from "../error/yml";
import { ValidateTeamsAppInputs } from "../question";
import { SPFxVersionOptionIds, ScratchOptions, createProjectCliHelpNode } from "../question/create";
import { HubTypes, isAadMainifestContainsPlaceholder } from "../question/other";
import { QuestionNames } from "../question/questionNames";
import { CallbackRegistry } from "./callback";
import { checkPermission, grantPermission, listCollaborator } from "./collaborator";
import { LocalCrypto } from "./crypto";
import { environmentNameManager } from "./environmentName";
import { InvalidInputError } from "./error";
import { ErrorContextMW, TOOLS, setErrorContext, setTools } from "./globalVars";
import { ConcurrentLockerMW } from "./middleware/concurrentLocker";
import { ContextInjectorMW } from "./middleware/contextInjector";
import { ErrorHandlerMW } from "./middleware/errorHandler";
import { ProjectMigratorMWV3, checkActiveResourcePlugins } from "./middleware/projectMigratorV3";
import {
  getProjectVersionFromPath,
  getTrackingIdFromPath,
  getVersionState,
} from "./middleware/utils/v3MigrationUtils";
import { CoreTelemetryComponentName, CoreTelemetryEvent, CoreTelemetryProperty } from "./telemetry";
import { CoreHookContext, PreProvisionResForVS, VersionCheckRes } from "./types";
import "../component/feature/sso";
import { AppStudioScopes } from "../component/driver/teamsApp/constants";
import { AppStudioClient } from "../component/driver/teamsApp/clients/appStudioClient";
import { AppDefinition } from "../component/driver/teamsApp/interfaces/appdefinitions/appDefinition";
import { TelemetryUtils } from "../component/driver/teamsApp/utils/telemetry";

export type CoreCallbackFunc = (name: string, err?: FxError, data?: any) => void | Promise<void>;

export class FxCore {
  constructor(tools: Tools) {
    setTools(tools);
  }

  /**
   * @todo this's a really primitive implement. Maybe could use Subscription Model to
   * refactor later.
   */
  public on(event: CoreCallbackEvent, callback: CoreCallbackFunc): void {
    return CallbackRegistry.set(event, callback);
  }

  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "createProject", reset: true }),
    ErrorHandlerMW,
    QuestionMW("createProject"),
  ])
  async createProject(inputs: Inputs): Promise<Result<CreateProjectResult, FxError>> {
    const context = createContextV3();
    inputs[QuestionNames.Scratch] = ScratchOptions.yes().id;
    if (inputs.teamsAppFromTdp) {
      // should never happen as we do same check on Developer Portal.
      if (containsUnsupportedFeature(inputs.teamsAppFromTdp)) {
        return err(InvalidInputError("Teams app contains unsupported features"));
      } else {
        context.telemetryReporter.sendTelemetryEvent(CoreTelemetryEvent.CreateFromTdpStart, {
          [CoreTelemetryProperty.TdpTeamsAppFeatures]: getFeaturesFromAppDefinition(
            inputs.teamsAppFromTdp
          ).join(","),
          [CoreTelemetryProperty.TdpTeamsAppId]: inputs.teamsAppFromTdp.teamsAppId,
        });
      }
    }
    const res = await coordinator.create(context, inputs);
    inputs.projectPath = context.projectPath;
    return res;
  }

  /**
   * lifecycle command: create new sample project
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "createSampleProject", reset: true }),
    ErrorHandlerMW,
    QuestionMW("createSampleProject"),
  ])
  async createSampleProject(inputs: Inputs): Promise<Result<CreateProjectResult, FxError>> {
    const context = createContextV3();
    inputs[QuestionNames.Scratch] = ScratchOptions.no().id;
    const res = await coordinator.create(context, inputs);
    inputs.projectPath = context.projectPath;
    return res;
  }
  /**
   * lifecycle commands: provision
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "provision", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    EnvLoaderMW(false),
    ConcurrentLockerMW,
    ContextInjectorMW,
    EnvWriterMW,
  ])
  async provisionResources(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<undefined, FxError>> {
    inputs.stage = Stage.provision;
    const context = createDriverContext(inputs);
    try {
      const res = await coordinator.provision(context, inputs as InputsWithProjectPath);
      if (res.isOk()) {
        ctx!.envVars = res.value;
        return ok(undefined);
      } else {
        // for partial success scenario, output is set in inputs object
        ctx!.envVars = inputs.envVars;
        return err(res.error);
      }
    } finally {
      //reset subscription
      try {
        await TOOLS.tokenProvider.azureAccountProvider.setSubscription("");
      } catch (e) {}
    }
  }
  /**
   * lifecycle commands: deploy
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "deploy", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    EnvLoaderMW(false),
    ConcurrentLockerMW,
    ContextInjectorMW,
    EnvWriterMW,
  ])
  async deployArtifacts(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<undefined, FxError>> {
    inputs.stage = Stage.deploy;
    const context = createDriverContext(inputs);
    const res = await coordinator.deploy(context, inputs as InputsWithProjectPath);
    if (res.isOk()) {
      ctx!.envVars = res.value;
      return ok(undefined);
    } else {
      // for partial success scenario, output is set in inputs object
      ctx!.envVars = inputs.envVars;
      return err(res.error);
    }
  }
  @hooks([ErrorContextMW({ component: "FxCore", stage: "localDebug", reset: true })])
  async localDebug(inputs: Inputs): Promise<Result<undefined, FxError>> {
    inputs.env = environmentNameManager.getLocalEnvName();
    return this.provisionResources(inputs);
  }

  /**
   * none lifecycle command, v3 only
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "deployAadManifest", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    QuestionMW("deployAadManifest"),
    EnvLoaderMW(true, true),
    ConcurrentLockerMW,
    ContextInjectorMW,
  ])
  async deployAadManifest(inputs: Inputs): Promise<Result<undefined, FxError>> {
    inputs.stage = Stage.deployAad;
    const updateAadClient = Container.get<UpdateAadAppDriver>("aadApp/update");
    // In V3, the aad.template.json exist at .fx folder, and output to root build folder.
    const manifestTemplatePath: string = inputs[QuestionNames.AadAppManifestFilePath];
    if (!(await fs.pathExists(manifestTemplatePath))) {
      return err(new FileNotFoundError("deployAadManifest", manifestTemplatePath));
    }
    let manifestOutputPath: string = manifestTemplatePath;
    if (inputs.env && (await isAadMainifestContainsPlaceholder(inputs))) {
      await fs.ensureDir(path.join(inputs.projectPath!, "build"));
      manifestOutputPath = path.join(
        inputs.projectPath!,
        "build",
        `aad.manifest.${inputs.env as string}.json`
      );
    }
    const inputArgs: UpdateAadAppArgs = {
      manifestPath: manifestTemplatePath,
      outputFilePath: manifestOutputPath,
    };
    const Context: DriverContext = createDriverContext(inputs);
    setErrorContext({ component: "aadAppUpdate" });
    const res = await updateAadClient.execute(inputArgs, Context);
    if (res.result.isErr()) {
      return err(res.result.error);
    }
    if (Context.platform === Platform.CLI) {
      const msg = getLocalizedString("core.deploy.aadManifestOnCLISuccessNotice");
      void Context.ui!.showMessage("info", msg, false);
    } else {
      const msg = getLocalizedString("core.deploy.aadManifestSuccessNotice");
      void Context.ui!.showMessage(
        "info",
        msg,
        false,
        getLocalizedString("core.deploy.aadManifestLearnMore")
      ).then((result) => {
        const userSelected = result.isOk() ? result.value : undefined;
        if (userSelected === getLocalizedString("core.deploy.aadManifestLearnMore")) {
          void Context.ui!.openUrl(ViewAadAppHelpLinkV5);
        }
      });
    }
    return ok(undefined);
  }
  /**
   * none lifecycle command, v3 only
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "addWebpart", reset: true }),
    ErrorHandlerMW,
    QuestionMW("addWebpart"),
    ProjectMigratorMWV3,
    ConcurrentLockerMW,
  ])
  async addWebpart(inputs: Inputs): Promise<Result<undefined, FxError>> {
    setErrorContext({ component: "spfxAdd", method: "run" });
    const driver: AddWebPartDriver = Container.get<AddWebPartDriver>("spfx/add");
    const args: AddWebPartArgs = {
      manifestPath: inputs[SPFxQuestionNames.ManifestPath],
      localManifestPath: inputs[SPFxQuestionNames.LocalManifestPath],
      spfxFolder: inputs[SPFxQuestionNames.SPFxFolder],
      webpartName: inputs[SPFxQuestionNames.WebPartName],
      framework: inputs[QuestionNames.SPFxFramework],
      spfxPackage: SPFxVersionOptionIds.installLocally,
    };
    const Context: DriverContext = createDriverContext(inputs);
    const res = await driver.run(args, Context);
    if (res.isErr()) {
      return err(res.error);
    }
    return ok(undefined);
  }
  /**
   * lifecycle command: publish
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "publish", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    EnvLoaderMW(false),
    ConcurrentLockerMW,
    ContextInjectorMW,
    EnvWriterMW,
  ])
  async publishApplication(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<undefined, FxError>> {
    inputs.stage = Stage.publish;
    const context = createDriverContext(inputs);
    const res = await coordinator.publish(context, inputs as InputsWithProjectPath);
    if (res.isOk()) {
      ctx!.envVars = res.value;
      return ok(undefined);
    } else {
      // for partial success scenario, output is set in inputs object
      ctx!.envVars = inputs.envVars;
      return err(res.error);
    }
  }
  /**
   * most commands will be deprecated in V3
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "executeUserTask", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    EnvLoaderMW(false),
    ConcurrentLockerMW,
  ])
  async executeUserTask(func: Func, inputs: Inputs): Promise<Result<any, FxError>> {
    let res: Result<any, FxError> = ok(undefined);
    const context = createDriverContext(inputs);
    if (func.method === "addSso") {
      // used in v3 only in VS
      inputs.stage = Stage.addFeature;
      inputs[QuestionNames.Features] = SingleSignOnOptionItem.id;
      const component = Container.get<SSO>("sso");
      setErrorContext({ component: "sso", method: "add" });
      res = await component.add(context as unknown as Context, inputs as InputsWithProjectPath);
    }
    return res;
  }
  /**
   * v3 only none lifecycle command
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "buildAadManifest", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    EnvLoaderMW(false),
    ConcurrentLockerMW,
  ])
  async buildAadManifest(inputs: Inputs): Promise<Result<undefined, FxError>> {
    const manifestTemplatePath: string = inputs.AAD_MANIFEST_FILE
      ? inputs.AAD_MANIFEST_FILE
      : path.join(inputs.projectPath!, AadConstants.DefaultTemplateFileName);
    if (!(await fs.pathExists(manifestTemplatePath))) {
      return err(new FileNotFoundError("buildAadManifest", manifestTemplatePath));
    }
    await fs.ensureDir(path.join(inputs.projectPath!, "build"));
    const manifestOutputPath: string = path.join(
      inputs.projectPath!,
      "build",
      // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
      `aad.${inputs.env}.json`
    );
    const Context: DriverContext = createDriverContext(inputs);
    await buildAadManifest(Context, manifestTemplatePath, manifestOutputPath);
    return ok(undefined);
  }
  /**
   * v3 only none lifecycle command
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "deployTeamsManifest", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    QuestionMW("selectTeamsAppManifest"),
    EnvLoaderMW(true),
    ConcurrentLockerMW,
    ContextInjectorMW,
    EnvWriterMW,
  ])
  async deployTeamsManifest(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<undefined, FxError>> {
    inputs.manifestTemplatePath = inputs[QuestionNames.TeamsAppManifestFilePath] as string;
    const context = createContextV3();
    const res = await updateManifestV3(context, inputs as InputsWithProjectPath);
    if (res.isOk()) {
      ctx!.envVars = envUtil.map2object(res.value);
      return ok(undefined);
    }
    return err(res.error);
  }

  /******
   * CLI v3 commands
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "updateTeamsAppCLIV3", reset: true }),
    ErrorHandlerMW,
  ])
  async updateTeamsAppCLIV3(inputs: TeamsAppInputs): Promise<Result<undefined, FxError>> {
    const res = await teamsappMgr.updateTeamsApp(inputs);
    return res;
  }
  /******
   * CLI v3 commands
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "validateTeamsAppCLIV3", reset: true }),
    ErrorHandlerMW,
  ])
  async validateTeamsAppCLIV3(inputs: TeamsAppInputs): Promise<Result<undefined, FxError>> {
    const res = await teamsappMgr.validateTeamsApp(inputs);
    return res;
  }
  /******
   * CLI v3 commands
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "packageTeamsAppCLIV3", reset: true }),
    ErrorHandlerMW,
  ])
  async packageTeamsAppCLIV3(inputs: TeamsAppInputs): Promise<Result<undefined, FxError>> {
    const res = await teamsappMgr.packageTeamsApp(inputs);
    if (res.isErr()) {
      return err(res.error);
    }
    return ok(undefined);
  }
  /******
   * CLI v3 commands
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "publishTeamsAppCLIV3", reset: true }),
    ErrorHandlerMW,
  ])
  async publishTeamsAppCLIV3(inputs: TeamsAppInputs): Promise<Result<undefined, FxError>> {
    const res = await teamsappMgr.publishTeamsApp(inputs);
    return res;
  }

  /******
   * CLI v3 commands
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "listTeamsApps", reset: true }),
    ErrorHandlerMW,
  ])
  async listTeamsApps(inputs: Inputs): Promise<Result<AppDefinition[], FxError>> {
    const appStudioTokenRes = await TOOLS.tokenProvider.m365TokenProvider.getAccessToken({
      scopes: AppStudioScopes,
    });
    if (appStudioTokenRes.isErr()) {
      return err(appStudioTokenRes.error);
    }
    const appStudioToken = appStudioTokenRes.value;
    TelemetryUtils.init(createDriverContext(inputs));
    const apps = await AppStudioClient.listApps(appStudioToken, TOOLS.logProvider);
    return ok(apps);
  }

  /******
   * CLI v3 commands
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "deleteTeamsApps", reset: true }),
    ErrorHandlerMW,
  ])
  async deleteTeamsApps(inputs: Inputs): Promise<Result<undefined, FxError>> {
    const appStudioTokenRes = await TOOLS.tokenProvider.m365TokenProvider.getAccessToken({
      scopes: AppStudioScopes,
    });
    if (appStudioTokenRes.isErr()) {
      return err(appStudioTokenRes.error);
    }
    const appStudioToken = appStudioTokenRes.value;
    TelemetryUtils.init(createDriverContext(inputs));
    for (const appId of inputs.teamsAppIds as string[]) {
      await AppStudioClient.deleteApp(appId, appStudioToken, TOOLS.logProvider);
      TOOLS.logProvider?.info(`Delete Teams app ${appId} succeeded.`);
    }
    return ok(undefined);
  }

  /**
   * v3 only none lifecycle command
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "validateApplication", reset: true }),
    QuestionMW("validateTeamsApp"),
  ])
  async validateApplication(inputs: ValidateTeamsAppInputs): Promise<Result<any, FxError>> {
    if (inputs["manifest-path"]) {
      return await this.validateManifest(inputs);
    } else {
      return await this.validateAppPackage(inputs);
    }
  }
  /**
   * v3 only none lifecycle command
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "validateManifest", reset: true }),
    ErrorHandlerMW,
    EnvLoaderMW(true),
    ConcurrentLockerMW,
  ])
  async validateManifest(inputs: ValidateTeamsAppInputs): Promise<Result<any, FxError>> {
    inputs.stage = Stage.validateApplication;
    const context: DriverContext = createDriverContext(inputs);
    const teamsAppManifestFilePath = inputs["manifest-path"] as string;
    const args: ValidateManifestArgs = {
      manifestPath: teamsAppManifestFilePath,
      showMessage: inputs?.showMessage != undefined ? inputs.showMessage : true,
    };
    const driver: ValidateManifestDriver = Container.get("teamsApp/validateManifest");
    const result = await driver.execute(args, context);
    return result.result;
  }
  /**
   * v3 only none lifecycle command
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "validateAppPackage", reset: true }),
    ErrorHandlerMW,
    ConcurrentLockerMW,
  ])
  async validateAppPackage(inputs: ValidateTeamsAppInputs): Promise<Result<any, FxError>> {
    inputs.stage = Stage.validateApplication;
    const context: DriverContext = createDriverContext(inputs);
    const teamsAppPackageFilePath = inputs["app-package-file-path"] as string;
    const args: ValidateAppPackageArgs = {
      appPackagePath: teamsAppPackageFilePath,
      showMessage: true,
    };
    const driver: ValidateAppPackageDriver = Container.get("teamsApp/validateAppPackage");
    return (await driver.execute(args, context)).result;
  }
  /**
   * v3 only none lifecycle command
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "createAppPackage", reset: true }),
    ErrorHandlerMW,
    QuestionMW("selectTeamsAppManifest"),
    EnvLoaderMW(true),
    ConcurrentLockerMW,
  ])
  async createAppPackage(inputs: Inputs): Promise<Result<any, FxError>> {
    inputs.stage = Stage.createAppPackage;

    const context: DriverContext = createDriverContext(inputs);

    const teamsAppManifestFilePath = inputs?.[QuestionNames.TeamsAppManifestFilePath] as string;

    const driver: CreateAppPackageDriver = Container.get("teamsApp/zipAppPackage");
    const args: CreateAppPackageArgs = {
      manifestPath: teamsAppManifestFilePath,
      outputZipPath:
        inputs[QuestionNames.OutputZipPathParamName] ??
        // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
        `${inputs.projectPath}/${AppPackageFolderName}/${BuildFolderName}/appPackage.${process.env
          .TEAMSFX_ENV!}.zip`,
      outputJsonPath:
        inputs[QuestionNames.OutputManifestParamName] ??
        // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
        `${inputs.projectPath}/${AppPackageFolderName}/${BuildFolderName}/manifest.${process.env
          .TEAMSFX_ENV!}.json`,
    };
    const result = (await driver.execute(args, context)).result;
    if (context.platform === Platform.VSCode) {
      if (result.isOk()) {
        const isWindows = process.platform === "win32";
        let zipFileName = args.outputZipPath;
        if (!path.isAbsolute(zipFileName)) {
          zipFileName = path.join(context.projectPath, zipFileName);
        }
        let builtSuccess = getLocalizedString(
          "plugins.appstudio.buildSucceedNotice.fallback",
          zipFileName
        );
        if (isWindows) {
          const folderLink = pathToFileURL(path.dirname(zipFileName));
          const appPackageLink = `${
            VSCodeExtensionCommand.openFolder
          }?%5B%22${folderLink.toString()}%22%5D`;
          builtSuccess = getLocalizedString("plugins.appstudio.buildSucceedNotice", appPackageLink);
        }
        context.ui?.showMessage("info", builtSuccess, false);
      }
    }
    return result;
  }
  /**
   * get url to preview the app, may prompt to select env, hub and Teams manifest
   * v3 only none lifecycle command
   * @param {Inputs} inputs
   * @returns the url to preview the app
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "preview", reset: true }),
    ErrorHandlerMW,
    QuestionMW("previewWithTeamsAppManifest"),
    EnvLoaderMW(false),
    ConcurrentLockerMW,
  ])
  async previewWithManifest(inputs: Inputs): Promise<Result<string, FxError>> {
    inputs.stage = Stage.previewWithManifest;

    const hub = inputs[QuestionNames.M365Host] as HubTypes;
    const manifestFilePath = inputs[QuestionNames.TeamsAppManifestFilePath] as string;

    const manifestRes = await manifestUtils.getManifestV3(manifestFilePath, undefined, false);
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }

    const teamsAppId = manifestRes.value.id;
    const properties = ManifestUtil.parseCommonProperties(manifestRes.value);

    const launchHelper = new LaunchHelper(TOOLS.tokenProvider.m365TokenProvider, TOOLS.logProvider);
    const result = await launchHelper.getLaunchUrl(
      hub,
      teamsAppId,
      properties.capabilities,
      true,
      properties.isApiME
    );
    return result;
  }
  /**
   * Warning: this API only works for CLI_HELP, it has no business with interactive run for CLI!
   */
  getQuestions(stage: Stage, inputs: Inputs): Result<IQTreeNode | undefined, FxError> {
    if (stage === Stage.create) {
      return ok(createProjectCliHelpNode());
    }
    return ok(undefined);
  }

  /**
   * get all dot envs
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "getDotEnvs", reset: true }),
    ErrorHandlerMW,
  ])
  async getDotEnvs(
    inputs: InputsWithProjectPath
  ): Promise<Result<{ [name: string]: DotenvParseOutput }, FxError>> {
    const envListRes = await envUtil.listEnv(inputs.projectPath);
    if (envListRes.isErr()) {
      return err(envListRes.error);
    }
    const res: { [name: string]: DotenvParseOutput } = {};
    for (const env of envListRes.value) {
      const envRes = await envUtil.readEnv(inputs.projectPath, env, false, false);
      if (envRes.isErr()) {
        return err(envRes.error);
      }
      res[env] = envRes.value as DotenvParseOutput;
    }
    return ok(res);
  }
  /**
   * given projectPath and filePath, return whether the filePath is a env file
   */
  @hooks([ErrorContextMW({ component: "FxCore", stage: "isEnvFile", reset: true })])
  async isEnvFile(projectPath: string, inputFile: string): Promise<Result<boolean, FxError>> {
    const inputFileName = path.basename(inputFile);
    const envName = envUtil.extractEnvNameFromFileName(inputFileName);
    if (!envName) return ok(false);
    const folderRes = await pathUtils.getEnvFolderPath(projectPath);
    if (folderRes.isErr()) return err(folderRes.error);
    const envFolderPath = folderRes.value;
    if (!envFolderPath) return ok(false);
    const inputFileDir = path.dirname(inputFile);
    if (path.resolve(inputFileDir) !== path.resolve(envFolderPath)) return ok(false);
    return ok(true);
  }

  /**
   * get projectId
   */
  @hooks([ErrorContextMW({ component: "FxCore", stage: "getProjectId", reset: true })])
  async getProjectId(projectPath: string): Promise<Result<string, FxError>> {
    const res = await this.getProjectMetadata(projectPath);
    if (res.isErr()) {
      return err(res.error);
    }
    return ok(res.value.projectId || "");
  }

  /**
   * @description get projectId and version from yml
   */
  @hooks([ErrorContextMW({ component: "FxCore", stage: "getProjectMetadata", reset: true })])
  async getProjectMetadata(
    projectPath: string
  ): Promise<Result<{ version?: string; projectId?: string }, FxError>> {
    try {
      const ymlPath = pathUtils.getYmlFilePath(projectPath, "dev");
      if (!ymlPath || !(await fs.pathExists(ymlPath))) {
        return ok({});
      }
      const ymlContent = await fs.readFile(ymlPath, "utf-8");
      const ymlObject = parse(ymlContent);
      return ok({
        projectId: ymlObject?.projectId ? ymlObject.projectId.toString() : "",
        version: ymlObject?.version ? ymlObject.version.toString() : "",
      });
    } catch {
      return ok({});
    }
  }

  /**
   * get Teams App Name from yml
   */
  @hooks([ErrorContextMW({ component: "FxCore", stage: "getTeamsAppName", reset: true })])
  async getTeamsAppName(projectPath: string): Promise<Result<string, FxError>> {
    const ymlPath = pathUtils.getYmlFilePath(projectPath, "dev");
    const maybeProjectModel = await metadataUtil.parse(ymlPath, "dev");
    if (maybeProjectModel.isErr()) {
      return err(maybeProjectModel.error);
    }
    const projectModel = maybeProjectModel.value as any;
    if (projectModel.provision) {
      const teamsAppCreate = projectModel.provision?.driverDefs.find(
        (d: any) => d.uses === "teamsApp/create"
      );
      if (teamsAppCreate) {
        let name = teamsAppCreate.with.name as string;
        if (name) {
          name = expandEnvironmentVariable(name, { APP_NAME_SUFFIX: "", TEAMSFX_ENV: " " }).trim();
          return ok(name);
        }
      }
    }
    return ok("");
  }

  /**
   * get project info
   */
  @hooks([ErrorContextMW({ component: "FxCore", stage: "getProjectInfo", reset: true })])
  async getProjectInfo(
    projectPath: string,
    env: string
  ): Promise<
    Result<
      {
        projectId: string;
        teamsAppId: string;
        teamsAppName: string;
        m365TenantId: string;
      },
      FxError
    >
  > {
    const ymlPath = pathUtils.getYmlFilePath(projectPath, env);
    const maybeProjectModel = await metadataUtil.parse(ymlPath, env);
    if (maybeProjectModel.isErr()) {
      return err(maybeProjectModel.error);
    }
    const projectModel = maybeProjectModel.value;
    const readEnvRes = await envUtil.readEnv(projectPath, env, false, true);
    if (readEnvRes.isErr()) {
      return err(readEnvRes.error);
    }
    const envObject = readEnvRes.value;
    const res: {
      projectId: string;
      teamsAppId: string;
      teamsAppName: string;
      m365TenantId: string;
    } = {
      projectId: (projectModel as any).projectId || "",
      teamsAppId: "",
      teamsAppName: "",
      m365TenantId: envObject.TEAMS_APP_TENANT_ID || "",
    };
    if (projectModel.provision) {
      const teamsAppCreate = projectModel.provision.driverDefs.find(
        (d) => d.uses === "teamsApp/create"
      );
      if (teamsAppCreate) {
        const teamsAppIdEnvName = teamsAppCreate.writeToEnvironmentFile?.teamsAppId;
        if (teamsAppIdEnvName) {
          const teamsAppId = envObject[teamsAppIdEnvName];
          res.teamsAppId = teamsAppId;
        }
        const name = (teamsAppCreate.with as any).name;
        if (name) {
          res.teamsAppName = name.replace("-${{TEAMSFX_ENV}}", "") || "";
        }
      }
    }
    return ok(res);
  }

  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "grantPermission", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    QuestionMW("grantPermission"),
    EnvLoaderMW(false, true),
    ConcurrentLockerMW,
    EnvWriterMW,
  ])
  async grantPermission(inputs: Inputs): Promise<Result<PermissionsResult, FxError>> {
    inputs.stage = Stage.grantPermission;
    const context = createContextV3();
    setErrorContext({ component: "collaborator" });
    const res = await grantPermission(
      context,
      inputs as InputsWithProjectPath,
      TOOLS.tokenProvider
    );
    return res;
  }
  /**
   * none lifecycle command
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "checkPermission", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    QuestionMW("listCollaborator"),
    EnvLoaderMW(false, true),
    ConcurrentLockerMW,
    EnvWriterMW,
  ])
  async checkPermission(inputs: Inputs): Promise<Result<PermissionsResult, FxError>> {
    inputs.stage = Stage.checkPermission;
    const context = createContextV3();
    const res = await checkPermission(
      context,
      inputs as InputsWithProjectPath,
      TOOLS.tokenProvider
    );
    return res;
  }
  /**
   * none lifecycle command
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "listCollaborator", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    QuestionMW("listCollaborator"),
    EnvLoaderMW(false, true),
    ConcurrentLockerMW,
    EnvWriterMW,
  ])
  async listCollaborator(inputs: Inputs): Promise<Result<ListCollaboratorResult, FxError>> {
    inputs.stage = Stage.listCollaborator;
    const context = createContextV3();
    const res = await listCollaborator(
      context,
      inputs as InputsWithProjectPath,
      TOOLS.tokenProvider
    );
    return res;
  }
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "getSelectedEnv", reset: true }),
    ErrorHandlerMW,
    EnvLoaderMW(false),
  ])
  getSelectedEnv(inputs: Inputs): Promise<Result<string | undefined, FxError>> {
    return Promise.resolve(ok(inputs.env)); //work for both v2 and v3
  }

  @hooks([ErrorContextMW({ component: "FxCore", stage: "createLocalCrypto" })])
  async createLocalCrypto(projectPath: string): Promise<Result<CryptoProvider, FxError>> {
    const settingsRes = await settingsUtil.readSettings(projectPath);
    if (settingsRes.isErr()) {
      return err(settingsRes.error);
    }
    const projectId = settingsRes.value.trackingId;
    const cryptoProvider = new LocalCrypto(projectId);
    return ok(cryptoProvider);
  }

  /**
   * only for vs code extension
   */
  @hooks([ErrorContextMW({ component: "FxCore", stage: "encrypt", reset: true }), ErrorHandlerMW])
  async encrypt(plaintext: string, inputs: Inputs): Promise<Result<string, FxError>> {
    const res = await this.createLocalCrypto(inputs.projectPath!);
    if (res.isErr()) {
      return err(res.error);
    }
    return res.value.encrypt(plaintext);
  }
  /**
   * only for vs code extension
   */
  @hooks([ErrorContextMW({ component: "FxCore", stage: "decrypt", reset: true }), ErrorHandlerMW])
  async decrypt(ciphertext: string, inputs: Inputs): Promise<Result<string, FxError>> {
    const res = await this.createLocalCrypto(inputs.projectPath!);
    if (res.isErr()) {
      return err(res.error);
    }
    return res.value.decrypt(ciphertext);
  }

  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "createEnv", reset: true }),
    ErrorHandlerMW,
    QuestionMW("createNewEnv"),
    ConcurrentLockerMW,
  ])
  async createEnv(inputs: Inputs): Promise<Result<undefined, FxError>> {
    return this.createEnvCopyV3(
      inputs[QuestionNames.NewTargetEnvName]!,
      inputs[QuestionNames.SourceEnvName]!,
      inputs.projectPath!
    );
  }
  async createEnvCopyV3(
    targetEnvName: string,
    sourceEnvName: string,
    projectPath: string
  ): Promise<Result<undefined, FxError>> {
    let res = await pathUtils.getEnvFilePath(projectPath, sourceEnvName);
    if (res.isErr()) return err(res.error);
    const sourceDotEnvFile = res.value;

    res = await pathUtils.getEnvFilePath(projectPath, targetEnvName);
    if (res.isErr()) return err(res.error);
    const targetDotEnvFile = res.value;
    if (!sourceDotEnvFile || !targetDotEnvFile)
      return err(new YamlFieldMissingError("environmentFolderPath"));
    if (!(await fs.pathExists(sourceDotEnvFile)))
      return err(new FileNotFoundError("createEnvCopyV3", sourceDotEnvFile));
    const source = await fs.readFile(sourceDotEnvFile);
    const writeStream = fs.createWriteStream(targetDotEnvFile);
    source
      .toString()
      .split(/\r?\n/)
      .forEach((line) => {
        const reg = /^([a-zA-Z_][a-zA-Z0-9_]*=)/g;
        const match = reg.exec(line);
        if (match) {
          if (match[1].startsWith("TEAMSFX_ENV=")) {
            writeStream.write(`TEAMSFX_ENV=${targetEnvName}${os.EOL}`);
          } else {
            writeStream.write(`${match[1]}${os.EOL}`);
          }
        } else {
          writeStream.write(`${line.trim()}${os.EOL}`);
        }
      });

    writeStream.end();
    TOOLS.logProvider.info(`env file created: ${targetDotEnvFile}`);
    return ok(undefined);
  }

  // a phantom migration method for V3
  @hooks([ErrorContextMW({ component: "FxCore", stage: "phantomMigrationV3", reset: true })])
  async phantomMigrationV3(inputs: Inputs): Promise<Result<undefined, FxError>> {
    // If the project is invalid or upgraded, the ProjectMigratorMWV3 will not take action.
    // Check invaliad/upgraded project here before call ProjectMigratorMWV3
    const projectPath = (inputs.projectPath as string) || "";
    const version = await getProjectVersionFromPath(projectPath);

    if (version.source === VersionSource.teamsapp) {
      return err(new NoNeedUpgradeError());
    } else if (version.source === VersionSource.projectSettings) {
      const isValid = await checkActiveResourcePlugins(projectPath);
      if (!isValid) {
        return err(new InvalidProjectError());
      }
    }
    if (version.source === VersionSource.unknown) {
      return err(new InvalidProjectError());
    }
    return this.innerMigrationV3(inputs);
  }

  @hooks([ErrorHandlerMW, ProjectMigratorMWV3])
  innerMigrationV3(inputs: Inputs): Result<undefined, FxError> {
    return ok(undefined);
  }

  // a project version check
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "projectVersionCheck", reset: true }),
    ErrorHandlerMW,
  ])
  async projectVersionCheck(inputs: Inputs): Promise<Result<VersionCheckRes, FxError>> {
    const projectPath = (inputs.projectPath as string) || "";
    if (isValidProjectV3(projectPath) || isValidProjectV2(projectPath)) {
      const versionInfo = await getProjectVersionFromPath(projectPath);
      if (!versionInfo.version) {
        return err(new InvalidProjectError());
      }
      const trackingId = await getTrackingIdFromPath(projectPath);
      const isSupport = getVersionState(versionInfo);
      // if the project is upgradeable, check whether the project is valid and invalid project should not show upgrade option.
      if (isSupport === VersionState.upgradeable) {
        if (!(await checkActiveResourcePlugins(projectPath))) {
          return err(new InvalidProjectError());
        }
      }
      return ok({
        currentVersion: versionInfo.version,
        trackingId,
        isSupport,
        versionSource: VersionSource[versionInfo.source],
      });
    } else {
      return err(new InvalidProjectError());
    }
  }

  // apply the given yaml template to current project.
  async apply(
    inputs: Inputs,
    templatePath: string,
    lifecycleName: string
  ): Promise<Result<undefined, FxError>> {
    if (!inputs.projectPath) {
      return err(InvalidInputError("invalid projectPath", inputs));
    }
    const projectPath = inputs.projectPath;
    if (!inputs.env) {
      return err(InvalidInputError("invalid env", inputs));
    }
    const env = inputs.env;
    const lifecycleName_: LifecycleName = lifecycleName as LifecycleName;
    const result = await envUtil.readEnv(projectPath, env);
    if (result.isErr()) {
      return err(result.error);
    }

    const parser = new YamlParser();
    const maybeProjectModel = await parser.parse(templatePath);
    if (maybeProjectModel.isErr()) {
      return err(maybeProjectModel.error);
    }

    const projectModel = maybeProjectModel.value;
    const driverContext: DriverContext = {
      azureAccountProvider: TOOLS.tokenProvider.azureAccountProvider,
      m365TokenProvider: TOOLS.tokenProvider.m365TokenProvider,
      ui: TOOLS.ui,
      progressBar: undefined,
      logProvider: TOOLS.logProvider,
      telemetryReporter: TOOLS.telemetryReporter!,
      projectPath: projectPath,
      platform: inputs.platform,
    };
    const lifecycle = projectModel[lifecycleName_];
    if (lifecycle) {
      return this.runLifecycle(lifecycle, driverContext, env);
    } else {
      driverContext.logProvider.warning(`No definition found for ${lifecycleName}`);
      return ok(undefined);
    }
  }

  async runLifecycle(
    lifecycle: ILifecycle,
    driverContext: DriverContext,
    env: string
  ): Promise<Result<undefined, FxError>> {
    const r = await lifecycle.execute(driverContext);
    const runResult = r.result;
    if (runResult.isOk()) {
      driverContext.logProvider.info(`Lifecycle ${lifecycle.name} succeeded`);
      const writeResult = await envUtil.writeEnv(
        driverContext.projectPath,
        env,
        envUtil.map2object(runResult.value)
      );
      return writeResult.map(() => undefined);
    } else {
      const error = runResult.error;
      if (error.kind === "Failure") {
        driverContext.logProvider.error(
          `Failed to run ${lifecycle.name} due to ${error.error.name}: ${error.error.message}`
        );
        return err(error.error);
      } else {
        try {
          const failedDriver = error.reason.failedDriver;
          if (error.reason.kind === "UnresolvedPlaceholders") {
            const unresolved = error.reason.unresolvedPlaceHolders;
            driverContext.logProvider.warning(
              `Unresolved placeholders: ${unresolved.join(",")} for driver ${failedDriver.uses}`
            );
            return ok(undefined);
          } else {
            driverContext.logProvider.error(
              `Failed to run ${lifecycle.name} due to ${error.reason.error.name}: ${error.reason.error.message}. Failed driver: ${failedDriver.uses}`
            );
            return err(error.reason.error);
          }
        } finally {
          await envUtil.writeEnv(driverContext.projectPath, env, envUtil.map2object(error.env));
        }
      }
    }
  }

  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "preProvisionForVS" }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    EnvLoaderMW(false),
    ConcurrentLockerMW,
    ContextInjectorMW,
  ])
  async preProvisionForVS(inputs: Inputs): Promise<Result<PreProvisionResForVS, FxError>> {
    const context = createDriverContext(inputs);
    return coordinator.preProvisionForVS(context, inputs as InputsWithProjectPath);
  }
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "preCheckYmlAndEnvForVS" }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    EnvLoaderMW(false),
    ConcurrentLockerMW,
    ContextInjectorMW,
  ])
  async preCheckYmlAndEnvForVS(inputs: Inputs): Promise<Result<undefined, FxError>> {
    const context = createDriverContext(inputs);
    const result = await coordinator.preCheckYmlAndEnvForVS(
      context,
      inputs as InputsWithProjectPath
    );
    return result;
  }
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "publishInDeveloperPortal" }),
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ContextInjectorMW,
  ])
  async publishInDeveloperPortal(inputs: Inputs): Promise<Result<undefined, FxError>> {
    inputs.stage = Stage.publishInDeveloperPortal;
    const context = createContextV3();
    return await coordinator.publishInDeveloperPortal(context, inputs as InputsWithProjectPath);
  }

  async injectCreateAPIKeyAction(
    ymlPath: string,
    authName: string,
    specRelativePath: string
  ): Promise<void> {
    const ymlContent = await fs.readFile(ymlPath, "utf-8");

    const document = parseDocument(ymlContent);
    const provisionNode = document.get("provision") as any;

    if (provisionNode) {
      const hasApiKeyAction = provisionNode.items.some(
        (item: any) =>
          item.get("uses") === "apiKey/register" && item.get("with")?.get("name") === authName
      );

      if (!hasApiKeyAction) {
        provisionNode.items = provisionNode.items.filter(
          (item: any) => item.get("uses") !== "apiKey/register"
        );
        let added = false;
        for (let i = 0; i < provisionNode.items.length; i++) {
          const item = provisionNode.items[i];
          if (item.get("uses") === "teamsApp/create") {
            const teamsAppId = item.get("writeToEnvironmentFile")?.get("teamsAppId") as string;
            if (teamsAppId) {
              provisionNode.items.splice(i + 1, 0, {
                uses: "apiKey/register",
                with: {
                  name: `${authName}`,
                  appId: `\${{${teamsAppId}}}`,
                  apiSpecPath: specRelativePath,
                },
                writeToEnvironmentFile: {
                  registrationId: `${authName.toUpperCase()}_REGISTRATION_ID`,
                },
              });
              added = true;
              break;
            }
          }
        }

        if (!added) {
          throw new InjectAPIKeyActionFailedError();
        }

        await fs.writeFile(ymlPath, document.toString(), "utf8");
      }
    } else {
      throw new InjectAPIKeyActionFailedError();
    }
  }

  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "copilotPluginAddAPI" }),
    ErrorHandlerMW,
    QuestionMW("copilotPluginAddAPI"),
    ConcurrentLockerMW,
  ])
  async copilotPluginAddAPI(inputs: Inputs): Promise<Result<undefined, FxError>> {
    const newOperations = inputs[QuestionNames.ApiOperation] as string[];
    const url = inputs[QuestionNames.ApiSpecLocation] ?? inputs.openAIPluginManifest?.api.url;
    const manifestPath = inputs[QuestionNames.ManifestPath];

    // Get API spec file path from manifest
    const manifestRes = await manifestUtils._readAppManifest(manifestPath);
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }
    const apiSpecificationFile = manifestRes.value.composeExtensions![0].apiSpecificationFile;
    const outputAPISpecPath = path.join(path.dirname(manifestPath), apiSpecificationFile!);

    // Merge existing operations in manifest.json
    const specParser = new SpecParser(url, {
      allowAPIKeyAuth: isApiKeyEnabled(),
      allowMultipleParameters: isMultipleParametersEnabled(),
    });
    const existingOperationIds = manifestUtils.getOperationIds(manifestRes.value);
    const apiResultList = await specParser.list();

    const existingOperations: string[] = [];
    for (const id of existingOperationIds) {
      const operation = apiResultList.find((item) => item.operationId === id);
      if (operation) {
        existingOperations.push(operation.api);
      }
    }

    const operations = [...existingOperations, ...newOperations];

    const adaptiveCardFolder = path.join(
      inputs.projectPath!,
      AppPackageFolderName,
      ResponseTemplatesFolderName
    );

    const context = createContextV3();

    try {
      if (isApiKeyEnabled()) {
        const authNames: Set<string> = new Set();
        const serverUrls: Set<string> = new Set();
        for (const api of operations) {
          const operation = apiResultList.find((op) => op.api === api);
          if (operation) {
            if (operation.auth && operation.auth.type === "apiKey") {
              authNames.add(operation.auth.name);
              serverUrls.add(operation.server);
            }
          }
        }

        if (authNames.size > 1) {
          throw new MultipleAuthError(authNames);
        }

        if (serverUrls.size > 1) {
          throw new MultipleServerError(serverUrls);
        }

        if (authNames.size === 1) {
          const ymlPath = path.join(inputs.projectPath!, MetadataV3.configFile);
          const localYamlPath = path.join(inputs.projectPath!, MetadataV3.localConfigFile);
          const authName = [...authNames][0];

          const relativeSpecPath =
            "./" + path.relative(inputs.projectPath!, outputAPISpecPath).replace(/\\/g, "/");

          await this.injectCreateAPIKeyAction(ymlPath, authName, relativeSpecPath);

          if (await fs.pathExists(localYamlPath)) {
            await this.injectCreateAPIKeyAction(localYamlPath, authName, relativeSpecPath);
          }
        }
      }

      const generateResult = await specParser.generate(
        manifestPath,
        operations,
        outputAPISpecPath,
        adaptiveCardFolder
      );

      // Send SpecParser.generate() warnings
      context.telemetryReporter.sendTelemetryEvent(specParserGenerateResultTelemetryEvent, {
        [specParserGenerateResultAllSuccessTelemetryProperty]: generateResult.allSuccess.toString(),
        [specParserGenerateResultWarningsTelemetryProperty]: generateResult.warnings
          .map((w) => w.type.toString() + ": " + w.content)
          .join(";"),
        [CoreTelemetryProperty.Component]: CoreTelemetryComponentName,
      });

      if (generateResult.warnings && generateResult.warnings.length > 0) {
        const warnSummary = generateScaffoldingSummary(
          generateResult.warnings,
          manifestRes.value,
          inputs.projectPath!
        );
        context.logProvider.info(warnSummary);
      }
    } catch (e) {
      let error: FxError;
      if (e instanceof SpecParserError) {
        error = convertSpecParserErrorToFxError(e);
      } else {
        error = assembleError(e);
      }
      return err(error);
    }

    const message = getLocalizedString(
      "core.copilot.addAPI.success",
      newOperations,
      inputs.projectPath
    );
    void context.userInteraction.showMessage("info", message, false);
    return ok(undefined);
  }

  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "copilotPluginLoadOpenAIManifest" }),
    ErrorHandlerMW,
  ])
  async copilotPluginLoadOpenAIManifest(
    inputs: Inputs
  ): Promise<Result<OpenAIPluginManifest, FxError>> {
    try {
      return ok(await OpenAIPluginManifestHelper.loadOpenAIPluginManifest(inputs.domain));
    } catch (error) {
      return err(error as FxError);
    }
  }
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "copilotPluginListOperations" }),
    ErrorHandlerMW,
  ])
  async copilotPluginListOperations(
    inputs: Inputs
  ): Promise<Result<ApiOperation[], ErrorResult[]>> {
    return await listOperations(
      createContextV3(),
      inputs.manifest,
      inputs.apiSpecUrl,
      inputs[QuestionNames.ManifestPath],
      inputs.includeExistingAPIs,
      inputs.shouldLogWarning
    );
  }

  /**
   * check project type info
   */
  @hooks([ErrorContextMW({ component: "FxCore", stage: "checkProjectType" }), ErrorHandlerMW])
  async checkProjectType(projectPath: string): Promise<Result<ProjectTypeResult, FxError>> {
    const projectTypeRes = await projectTypeChecker.checkProjectType(projectPath);
    const props: Record<string, string> = {};
    fillinProjectTypeProperties(props, projectTypeRes);
    TOOLS.telemetryReporter?.sendTelemetryEvent(TelemetryEvent.ProjectType, props);
    return ok(projectTypeRes);
  }
}
