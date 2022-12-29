// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as os from "os";
import fs from "fs-extra";
import * as path from "path";
import { Container } from "typedi";
import { hooks } from "@feathersjs/hooks";
import {
  err,
  Func,
  FxError,
  Inputs,
  InputsWithProjectPath,
  ok,
  Platform,
  ProjectSettingsV3,
  Result,
  Settings,
  Stage,
  Tools,
  UserCancelError,
  Void,
} from "@microsoft/teamsfx-api";

import { AadConstants } from "../component/constants";
import { environmentManager } from "./environment";
import {
  ObjectIsUndefinedError,
  NoAadManifestExistError,
  InvalidInputError,
  InvalidProjectError,
} from "./error";
import { setCurrentStage, TOOLS } from "./globalVars";
import { ConcurrentLockerMW } from "./middleware/concurrentLocker";
import { ProjectConsolidateMW } from "./middleware/consolidateLocalRemote";
import { ContextInjectorMW } from "./middleware/contextInjector";
import { askNewEnvironment } from "./middleware/envInfoLoaderV3";
import { ErrorHandlerMW } from "./middleware/errorHandler";
import { CoreHookContext, PreProvisionResForVS, VersionCheckRes } from "./types";
import { createContextV3, createDriverContext } from "../component/utils";
import { manifestUtils } from "../component/resource/appManifest/utils/ManifestUtils";
import "../component/driver/index";
import { UpdateAadAppDriver } from "../component/driver/aad/update";
import { UpdateAadAppArgs } from "../component/driver/aad/interface/updateAadAppArgs";
import { ValidateTeamsAppDriver } from "../component/driver/teamsApp/validate";
import { ValidateTeamsAppArgs } from "../component/driver/teamsApp/interfaces/ValidateTeamsAppArgs";
import { DriverContext } from "../component/driver/interface/commonArgs";
import { coordinator } from "../component/coordinator";
import { CreateAppPackageDriver } from "../component/driver/teamsApp/createAppPackage";
import { CreateAppPackageArgs } from "../component/driver/teamsApp/interfaces/CreateAppPackageArgs";
import { EnvLoaderMW, EnvWriterMW } from "../component/middleware/envMW";
import { envUtil } from "../component/utils/envUtil";
import { settingsUtil } from "../component/utils/settingsUtil";
import { DotenvParseOutput } from "dotenv";
import { ProjectMigratorMWV3 } from "./middleware/projectMigratorV3";
import {
  containsUnsupportedFeature,
  getFeaturesFromAppDefinition,
} from "../component/resource/appManifest/utils/utils";
import { CoreTelemetryEvent, CoreTelemetryProperty } from "./telemetry";
import { isValidProjectV2, isValidProjectV3 } from "../common/projectSettingsHelper";
import {
  getVersionState,
  getProjectVersionFromPath,
  getTrackingIdFromPath,
} from "./middleware/utils/v3MigrationUtils";
import { QuestionMW } from "../component/middleware/questionMW";
import { getQuestionsForCreateProjectV2 } from "./middleware/questionModel";
import { getQuestionsForInit, getQuestionsForProvisionV3 } from "../component/question";
import { isFromDevPortalInVSC } from "../component/developerPortalScaffoldUtils";

export class FxCoreV3Implement {
  tools: Tools;
  isFromSample?: boolean;
  settingsVersion?: string;

  constructor(tools: Tools) {
    this.tools = tools;
  }

  async dispatch<Inputs, ExecuteRes>(
    exec: (inputs: Inputs) => Promise<ExecuteRes>,
    inputs: Inputs
  ): Promise<ExecuteRes> {
    const methodName = exec.name as keyof FxCoreV3Implement;
    if (!this[methodName]) {
      throw new Error("no implement");
    }
    const method = this[methodName] as any as typeof exec;
    return await method.call(this, inputs);
  }

  async dispatchUserTask<Inputs, ExecuteRes>(
    exec: (func: Func, inputs: Inputs) => Promise<ExecuteRes>,
    func: Func,
    inputs: Inputs
  ): Promise<ExecuteRes> {
    const methodName = exec.name as keyof FxCoreV3Implement;
    if (!this[methodName]) {
      throw new Error("no implement");
    }
    const method = this[methodName] as any as typeof exec;
    return await method.call(this, func, inputs);
  }

  @hooks([ErrorHandlerMW, QuestionMW(getQuestionsForCreateProjectV2), ContextInjectorMW])
  async createProject(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<string, FxError>> {
    if (!ctx) {
      return err(new ObjectIsUndefinedError("ctx for createProject"));
    }
    setCurrentStage(Stage.create);
    inputs.stage = Stage.create;
    const context = createContextV3();
    if (isFromDevPortalInVSC(inputs)) {
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
    const res = await coordinator.create(context, inputs as InputsWithProjectPath);
    if (res.isErr()) return err(res.error);
    ctx.projectSettings = context.projectSetting;
    inputs.projectPath = context.projectPath;
    return ok(inputs.projectPath!);
  }

  @hooks([
    ErrorHandlerMW,
    QuestionMW((inputs) => {
      return getQuestionsForInit("infra", inputs);
    }),
  ])
  async initInfra(inputs: Inputs): Promise<Result<undefined, FxError>> {
    const res = await coordinator.initInfra(createContextV3(), inputs);
    return res;
  }

  @hooks([
    ErrorHandlerMW,
    QuestionMW((inputs) => {
      return getQuestionsForInit("debug", inputs);
    }),
  ])
  async initDebug(inputs: Inputs): Promise<Result<undefined, FxError>> {
    const res = await coordinator.initDebug(createContextV3(), inputs);
    return res;
  }

  @hooks([
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    QuestionMW(getQuestionsForProvisionV3),
    ConcurrentLockerMW,
    EnvLoaderMW(false),
    ContextInjectorMW,
    EnvWriterMW,
  ])
  async provisionResources(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    setCurrentStage(Stage.provision);
    inputs.stage = Stage.provision;
    const context = createDriverContext(inputs);
    try {
      const res = await coordinator.provision(context, inputs as InputsWithProjectPath);
      if (res.isOk()) {
        ctx!.envVars = res.value;
        return ok(Void);
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

  @hooks([
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    ConcurrentLockerMW,
    EnvLoaderMW(false),
    ContextInjectorMW,
    EnvWriterMW,
  ])
  async deployArtifacts(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    setCurrentStage(Stage.deploy);
    inputs.stage = Stage.deploy;
    const context = createDriverContext(inputs);
    const res = await coordinator.deploy(context, inputs as InputsWithProjectPath);
    if (res.isOk()) {
      ctx!.envVars = res.value;
      return ok(Void);
    } else {
      // for partial success scenario, output is set in inputs object
      ctx!.envVars = inputs.envVars;
      return err(res.error);
    }
  }

  @hooks([
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    ConcurrentLockerMW,
    ProjectConsolidateMW,
    EnvLoaderMW(false),
    ContextInjectorMW,
    EnvWriterMW,
  ])
  async deployAadManifest(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    setCurrentStage(Stage.deployAad);
    inputs.stage = Stage.deployAad;
    const updateAadClient = Container.get<UpdateAadAppDriver>("aadApp/update");
    // In V3, the aad.template.json exist at .fx folder, and output to root build folder.
    const manifestTemplatePath: string = inputs.AAD_MANIFEST_FILE
      ? inputs.AAD_MANIFEST_FILE
      : path.join(inputs.projectPath!, AadConstants.DefaultTemplateFileName);
    if (!(await fs.pathExists(manifestTemplatePath))) {
      return err(new NoAadManifestExistError(manifestTemplatePath));
    }
    await fs.ensureDir(path.join(inputs.projectPath!, "build"));
    const manifestOutputPath: string = path.join(
      inputs.projectPath!,
      "build",
      `aad.${inputs.env}.json`
    );
    const inputArgs: UpdateAadAppArgs = {
      manifestTemplatePath: manifestTemplatePath,
      outputFilePath: manifestOutputPath,
    };
    const contextV3: DriverContext = {
      azureAccountProvider: TOOLS.tokenProvider.azureAccountProvider,
      m365TokenProvider: TOOLS.tokenProvider.m365TokenProvider,
      ui: TOOLS.ui,
      logProvider: TOOLS.logProvider,
      telemetryReporter: TOOLS.telemetryReporter!,
      projectPath: inputs.projectPath as string,
      platform: Platform.VSCode,
    };
    const res = await updateAadClient.run(inputArgs, contextV3);
    if (res.isErr()) return err(res.error);
    return ok(Void);
  }

  @hooks([ErrorHandlerMW, ProjectMigratorMWV3, ConcurrentLockerMW, EnvLoaderMW(false)])
  async publishApplication(inputs: Inputs): Promise<Result<Void, FxError>> {
    setCurrentStage(Stage.publish);
    inputs.stage = Stage.publish;
    const context = createDriverContext(inputs);
    const res = await coordinator.publish(context, inputs as InputsWithProjectPath);
    if (res.isErr()) return err(res.error);
    return ok(Void);
  }

  @hooks([
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    ConcurrentLockerMW,
    EnvLoaderMW(true),
    ContextInjectorMW,
    EnvWriterMW,
  ])
  async deployTeamsManifest(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    const context = createContextV3(ctx?.projectSettings as ProjectSettingsV3);
    const component = Container.get("app-manifest") as any;
    const res = await component.deployV3(context, inputs as InputsWithProjectPath);
    if (res.isOk()) {
      ctx!.envVars = envUtil.map2object(res.value);
    }
    return res;
  }

  @hooks([ErrorHandlerMW, ProjectMigratorMWV3, ConcurrentLockerMW, EnvLoaderMW(false)])
  async executeUserTask(
    func: Func,
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<any, FxError>> {
    let res: Result<any, FxError> = ok(undefined);
    const context = createDriverContext(inputs);
    if (func.method === "getManifestTemplatePath") {
      const path = await manifestUtils.getTeamsAppManifestPath(
        (inputs as InputsWithProjectPath).projectPath
      );
      res = ok(path);
    } else if (func.method === "validateManifest") {
      const driver: ValidateTeamsAppDriver = Container.get("teamsApp/validate");
      const args: ValidateTeamsAppArgs = {
        manifestPath: func.params.manifestTemplatePath,
      };
      res = await driver.run(args, context);
    } else if (func.method === "buildPackage") {
      const driver: CreateAppPackageDriver = Container.get("teamsApp/zipAppPackage");
      const args: CreateAppPackageArgs = {
        manifestPath: func.params.manifestTemplatePath,
        outputZipPath: func.params.outputZipPath,
        outputJsonPath: func.params.outputJsonPath,
      };
      res = await driver.run(args, context);
    }
    return res;
  }

  @hooks([ErrorHandlerMW, ConcurrentLockerMW, ContextInjectorMW])
  async publishInDeveloperPortal(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<Void, FxError>> {
    setCurrentStage(Stage.publishInDeveloperPortal);
    inputs.stage = Stage.publishInDeveloperPortal;
    const context = createContextV3();
    return await coordinator.publishInDeveloperPortal(context, inputs as InputsWithProjectPath);
  }

  async getSettings(inputs: InputsWithProjectPath): Promise<Result<Settings, FxError>> {
    return settingsUtil.readSettings(inputs.projectPath);
  }

  @hooks([ErrorHandlerMW, EnvLoaderMW(true), ContextInjectorMW])
  async getDotEnv(
    inputs: InputsWithProjectPath,
    ctx?: CoreHookContext
  ): Promise<Result<DotenvParseOutput | undefined, FxError>> {
    return ok(ctx?.envVars);
  }

  @hooks([ErrorHandlerMW, ProjectMigratorMWV3])
  async phantomMigrationV3(inputs: Inputs): Promise<Result<Void, FxError>> {
    return ok(Void);
  }

  @hooks([ErrorHandlerMW])
  async projectVersionCheck(inputs: Inputs): Promise<Result<VersionCheckRes, FxError>> {
    const projectPath = (inputs.projectPath as string) || "";
    if (isValidProjectV3(projectPath) || isValidProjectV2(projectPath)) {
      const currentVersion = await getProjectVersionFromPath(projectPath);
      if (!currentVersion) {
        return err(new InvalidProjectError());
      }
      const trackingId = await getTrackingIdFromPath(projectPath);
      const isSupport = getVersionState(currentVersion);
      return ok({
        currentVersion,
        trackingId,
        isSupport,
      });
    } else {
      return err(new InvalidProjectError());
    }
  }

  @hooks([
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    ConcurrentLockerMW,
    EnvLoaderMW(false),
    ContextInjectorMW,
  ])
  async preProvisionForVS(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<PreProvisionResForVS, FxError>> {
    const context = createDriverContext(inputs);
    return coordinator.preProvisionForVS(context, inputs as InputsWithProjectPath);
  }

  @hooks([ErrorHandlerMW, ConcurrentLockerMW, ContextInjectorMW])
  async createEnv(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    if (!ctx || !inputs.projectPath)
      return err(new ObjectIsUndefinedError("createEnv input stuff"));

    const createEnvCopyInput = await askNewEnvironment(ctx!, inputs);
    if (
      !createEnvCopyInput ||
      !createEnvCopyInput.targetEnvName ||
      !createEnvCopyInput.sourceEnvName
    ) {
      return err(UserCancelError);
    }

    return this.createEnvCopyV3(
      createEnvCopyInput.targetEnvName,
      createEnvCopyInput.sourceEnvName,
      inputs.projectPath
    );
  }

  async createEnvCopyV3(
    targetEnvName: string,
    sourceEnvName: string,
    projectPath: string
  ): Promise<Result<Void, FxError>> {
    const sourceDotEnvFile = environmentManager.getDotEnvPath(sourceEnvName, projectPath);
    const source = await fs.readFile(sourceDotEnvFile);
    const targetDotEnvFile = environmentManager.getDotEnvPath(targetEnvName, projectPath);
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
    return ok(Void);
  }
}
