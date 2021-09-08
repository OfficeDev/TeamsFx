// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { HookContext, hooks } from "@feathersjs/hooks";
import {
  AppPackageFolderName,
  ArchiveFolderName,
  ArchiveLogFileName,
  assembleError,
  ConfigFolderName,
  Core,
  err,
  Func,
  FunctionRouter,
  FxError,
  GroupOfTasks,
  Inputs,
  LogProvider,
  ok,
  OptionItem,
  Platform,
  ProjectConfig,
  ProjectSettings,
  QTreeNode,
  Result,
  RunnableTask,
  SingleSelectQuestion,
  Solution,
  SolutionConfig,
  SolutionContext,
  Stage,
  TelemetryReporter,
  Tools,
  Void,
} from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import { AxiosResponse } from "axios";
import * as fs from "fs-extra";
import * as jsonschema from "jsonschema";
import * as path from "path";
import * as uuid from "uuid";
import { environmentManager } from "..";
import { globalStateUpdate } from "../common/globalState";
import {
  Component,
  sendTelemetryErrorEvent,
  sendTelemetryEvent,
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../common/telemetry";
import {
  downloadSampleHook,
  fetchCodeZip,
  isArmSupportEnabled,
  isMultiEnvEnabled,
  saveFilesRecursively,
} from "../common/tools";
import { getParameterJson } from "../plugins/solution/fx-solution/arm";
import { LocalCrypto } from "./crypto";
import {
  FetchSampleError,
  FunctionRouterError,
  InvalidInputError,
  MigrateNotImplementError,
  NonExistEnvNameError,
  ProjectFolderExistError,
  ProjectFolderNotExistError,
  TaskNotSupportError,
  WriteFileError,
} from "./error";
import { defaultSolutionLoader } from "./loader";
import { ConcurrentLockerMW } from "./middleware/concurrentLocker";
import { ContextInjecterMW } from "./middleware/contextInjecter";
import {
  askNewEnvironment,
  EnvInfoLoaderMW,
  loadSolutionContext,
  upgradeDefaultFunctionName,
  upgradeProgrammingLanguage,
} from "./middleware/envInfoLoader";
import { EnvInfoWriterMW } from "./middleware/envInfoWriter";
import { ErrorHandlerMW } from "./middleware/errorHandler";
import { LocalSettingsLoaderMW } from "./middleware/localSettingsLoader";
import { LocalSettingsWriterMW } from "./middleware/localSettingsWriter";
import { MigrateConditionHandlerMW } from "./middleware/migrateConditionHandler";
import { newSolutionContext, ProjectSettingsLoaderMW } from "./middleware/projectSettingsLoader";
import { ProjectSettingsWriterMW } from "./middleware/projectSettingsWriter";
import { ProjectUpgraderMW } from "./middleware/projectUpgrader";
import { QuestionModelMW } from "./middleware/questionModel";
import { SolutionLoaderMW } from "./middleware/solutionLoader";
import { PermissionRequestFileProvider } from "./permissionRequest";
import {
  CoreQuestionNames,
  DefaultAppNameFunc,
  getCreateNewOrFromSampleQuestion,
  ProjectNamePattern,
  QuestionAppName,
  QuestionRootFolder,
  QuestionSelectSolution,
  QuestionV1AppName,
  SampleSelect,
  ScratchOptionNo,
  ScratchOptionYes,
} from "./question";
import { newEnvInfo } from "./tools";
import { FxCoreV2 } from "./v2";

export interface CoreHookContext extends HookContext {
  projectSettings?: ProjectSettings;
  projectIdMissing?: boolean;
  solutionContext?: SolutionContext;
  solution?: Solution;
}

export let Logger: LogProvider;
export let telemetryReporter: TelemetryReporter | undefined;
export let currentStage: Stage;
export let TOOLS: Tools;
export class FxCore implements Core {
  version = "1";
  tools: Tools;

  constructor(tools: Tools) {
    this.tools = tools;
    TOOLS = tools;
    Logger = tools.logProvider;
    telemetryReporter = tools.telemetryReporter;
  }

  @hooks([
    ErrorHandlerMW,
    QuestionModelMW,
    ContextInjecterMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(isMultiEnvEnabled()),
  ])
  async createProject(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<string, FxError>> {
    currentStage = Stage.create;
    const folder = inputs[QuestionRootFolder.name] as string;
    const scratch = inputs[CoreQuestionNames.CreateFromScratch] as string;
    let projectPath: string;
    let globalStateDescription = "openReadme";

    if (scratch === ScratchOptionNo.id) {
      // create from sample
      const downloadRes = await downloadSample(this, inputs);
      if (downloadRes.isErr()) {
        return err(downloadRes.error);
      }
      projectPath = downloadRes.value;
      globalStateDescription = "openSampleReadme";
    } else {
      // create from new
      const appName = inputs[QuestionAppName.name] as string;
      if (undefined === appName) return err(InvalidInputError(`App Name is empty`, inputs));

      const validateResult = jsonschema.validate(appName, {
        pattern: ProjectNamePattern,
      });
      if (validateResult.errors && validateResult.errors.length > 0) {
        return err(InvalidInputError(`${validateResult.errors[0].message}`, inputs));
      }

      projectPath = path.join(folder, appName);
      const folderExist = await fs.pathExists(projectPath);
      if (folderExist) {
        return err(ProjectFolderExistError(projectPath));
      }

      inputs.projectPath = projectPath;
      const solution = await defaultSolutionLoader.loadSolution(inputs);
      const projectSettings: ProjectSettings = {
        appName: appName,
        projectId: uuid.v4(),
        solutionSettings: {
          name: solution.name,
          version: "1.0.0",
        },
      };

      if (isMultiEnvEnabled()) {
        projectSettings.activeEnvironment = environmentManager.getDefaultEnvName();
      }

      const solutionContext: SolutionContext = {
        projectSettings: projectSettings,
        envInfo: newEnvInfo(),
        root: projectPath,
        ...this.tools,
        ...this.tools.tokenProvider,
        answers: inputs,
      };

      await fs.ensureDir(projectPath);
      await fs.ensureDir(path.join(projectPath, `.${ConfigFolderName}`));
      await fs.ensureDir(
        path.join(
          projectPath,
          isMultiEnvEnabled()
            ? path.join("templates", `${AppPackageFolderName}`)
            : `${AppPackageFolderName}`
        )
      );

      const createResult = await this.createBasicFolderStructure(inputs);
      if (createResult.isErr()) {
        return err(createResult.error);
      }

      const createRes = await solution.create(solutionContext);
      if (createRes.isErr()) {
        return createRes;
      }

      const scaffoldRes = await solution.scaffold(solutionContext);
      if (scaffoldRes.isErr()) {
        return scaffoldRes;
      }

      if (isMultiEnvEnabled()) {
        const createEnvResult = await this.createEnvWithName(
          environmentManager.getDefaultEnvName(),
          projectSettings,
          inputs,
          ctx!.self as FxCore
        );
        if (createEnvResult.isErr()) {
          return err(createEnvResult.error);
        }
      }

      ctx!.solution = solution;
      ctx!.solutionContext = solutionContext;
    }

    if (inputs.platform === Platform.VSCode) {
      await globalStateUpdate(globalStateDescription, true);
    }

    return ok(projectPath);
  }

  @hooks([
    ErrorHandlerMW,
    MigrateConditionHandlerMW,
    QuestionModelMW,
    ContextInjecterMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(),
  ])
  async migrateV1Project(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<string, FxError>> {
    currentStage = Stage.migrateV1;
    const globalStateDescription = "openReadme";

    const appName = (inputs[DefaultAppNameFunc.name] ?? inputs[QuestionV1AppName.name]) as string;
    if (undefined === appName) return err(InvalidInputError(`App Name is empty`, inputs));

    const validateResult = jsonschema.validate(appName, {
      pattern: ProjectNamePattern,
    });
    if (validateResult.errors && validateResult.errors.length > 0) {
      return err(InvalidInputError(`${validateResult.errors[0].message}`, inputs));
    }

    const projectPath = inputs.projectPath;

    if (!projectPath || !(await fs.pathExists(projectPath))) {
      return err(ProjectFolderNotExistError(projectPath ?? ""));
    }

    const solution = await defaultSolutionLoader.loadSolution(inputs);
    const projectSettings: ProjectSettings = {
      appName: appName,
      projectId: uuid.v4(),
      solutionSettings: {
        name: solution.name,
        version: "1.0.0",
        migrateFromV1: true,
      },
    };

    const solutionContext: SolutionContext = {
      projectSettings: projectSettings,
      envInfo: newEnvInfo(),
      root: projectPath,
      ...this.tools,
      ...this.tools.tokenProvider,
      answers: inputs,
    };

    await this.archive(projectPath);
    await fs.ensureDir(projectPath);
    await fs.ensureDir(path.join(projectPath, `.${ConfigFolderName}`));

    const createResult = await this.createBasicFolderStructure(inputs);
    if (createResult.isErr()) {
      return err(createResult.error);
    }

    if (!solution.migrate) {
      return err(MigrateNotImplementError(projectPath));
    }
    const migrateV1Res = await solution.migrate(solutionContext);
    if (migrateV1Res.isErr()) {
      return migrateV1Res;
    }

    ctx!.solution = solution;
    ctx!.solutionContext = solutionContext;

    if (inputs.platform === Platform.VSCode) {
      await globalStateUpdate(globalStateDescription, true);
    }

    return ok(projectPath);
  }

  async archive(projectPath: string): Promise<void> {
    const archiveFolderPath = path.join(projectPath, ArchiveFolderName);
    await fs.ensureDir(archiveFolderPath);

    const fileNames = await fs.readdir(projectPath);
    const archiveLog = async (projectPath: string, message: string): Promise<void> => {
      await fs.appendFile(
        path.join(projectPath, ArchiveLogFileName),
        `[${new Date().toISOString()}] ${message}\n`
      );
    };

    await archiveLog(projectPath, `Start to move files into '${ArchiveFolderName}' folder.`);
    for (const fileName of fileNames) {
      if (fileName === ArchiveFolderName || fileName === ArchiveLogFileName) {
        continue;
      }

      try {
        await fs.move(path.join(projectPath, fileName), path.join(archiveFolderPath, fileName), {
          overwrite: true,
        });
      } catch (e: any) {
        await archiveLog(projectPath, `Failed to move '${fileName}'. ${e.message}`);
        throw e;
      }

      await archiveLog(
        projectPath,
        `'${fileName}' has been moved to '${ArchiveFolderName}' folder.`
      );
    }
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(isMultiEnvEnabled()),
    SolutionLoaderMW(defaultSolutionLoader),
    QuestionModelMW,
    ContextInjecterMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(),
  ])
  async provisionResources(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    currentStage = Stage.provision;
    return await ctx!.solution!.provision(ctx!.solutionContext!);
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(isMultiEnvEnabled()),
    SolutionLoaderMW(defaultSolutionLoader),
    QuestionModelMW,
    ContextInjecterMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(),
  ])
  async deployArtifacts(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    currentStage = Stage.deploy;
    return await ctx!.solution!.deploy(ctx!.solutionContext!);
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectUpgraderMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(false),
    LocalSettingsLoaderMW,
    SolutionLoaderMW(defaultSolutionLoader),
    QuestionModelMW,
    ContextInjecterMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(),
    LocalSettingsWriterMW,
  ])
  async localDebug(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    currentStage = Stage.debug;
    upgradeProgrammingLanguage(
      ctx!.solutionContext!.envInfo.profile as SolutionConfig,
      ctx!.projectSettings!
    );
    upgradeDefaultFunctionName(
      ctx!.solutionContext!.envInfo.profile as SolutionConfig,
      ctx!.projectSettings!
    );

    return await ctx!.solution!.localDebug(ctx!.solutionContext!);
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(isMultiEnvEnabled()),
    SolutionLoaderMW(defaultSolutionLoader),
    QuestionModelMW,
    ContextInjecterMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(),
  ])
  async publishApplication(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    currentStage = Stage.publish;
    return await ctx!.solution!.publish(ctx!.solutionContext!);
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(isMultiEnvEnabled()),
    LocalSettingsLoaderMW,
    SolutionLoaderMW(defaultSolutionLoader),
    QuestionModelMW,
    ContextInjecterMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(),
    LocalSettingsWriterMW,
  ])
  async executeUserTask(
    func: Func,
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<unknown, FxError>> {
    currentStage = Stage.userTask;
    if (ctx!.solutionContext === undefined)
      ctx!.solutionContext = await newSolutionContext(this.tools, inputs);
    const solution = ctx!.solution!;
    const namespace = func.namespace;
    const array = namespace ? namespace.split("/") : [];
    if ("" !== namespace && array.length > 0 && solution.executeUserTask) {
      return await solution.executeUserTask(func, ctx!.solutionContext!);
    }
    return err(FunctionRouterError(func));
  }

  @hooks([
    ErrorHandlerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(false),
    SolutionLoaderMW(defaultSolutionLoader),
    ContextInjecterMW,
    EnvInfoWriterMW(),
  ])
  async getQuestions(
    task: Stage,
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    if (task === Stage.create) {
      delete inputs.projectPath;
      return await this._getQuestionsForCreateProject(inputs);
    } else {
      const solutionContext =
        ctx!.solutionContext === undefined
          ? await newSolutionContext(this.tools, inputs)
          : ctx!.solutionContext;
      const solution =
        ctx!.solution === undefined
          ? await defaultSolutionLoader.loadSolution(inputs)
          : ctx!.solution;
      return await this._getQuestions(solutionContext, solution, task, inputs);
    }
  }

  @hooks([
    ErrorHandlerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(false),
    SolutionLoaderMW(defaultSolutionLoader),
    ContextInjecterMW,
    EnvInfoWriterMW(),
  ])
  async getQuestionsForUserTask(
    func: FunctionRouter,
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    const solutionContext =
      ctx!.solutionContext === undefined
        ? await newSolutionContext(this.tools, inputs)
        : ctx!.solutionContext;
    const solution =
      ctx!.solution === undefined
        ? await defaultSolutionLoader.loadSolution(inputs)
        : ctx!.solution;
    return await this._getQuestionsForUserTask(solutionContext, solution, func, inputs);
  }

  @hooks([
    ErrorHandlerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(isMultiEnvEnabled()),
    LocalSettingsLoaderMW,
    ContextInjecterMW,
  ])
  async getProjectConfig(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<ProjectConfig | undefined, FxError>> {
    return ok({
      settings: ctx!.projectSettings,
      config: ctx!.solutionContext?.envInfo.profile,
      localSettings: ctx!.solutionContext?.localSettings,
    });
  }

  @hooks([
    ErrorHandlerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(false),
    ContextInjecterMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(),
  ])
  async setSubscriptionInfo(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    const solutionContext = ctx!.solutionContext! as SolutionContext;
    if (inputs.tenantId)
      solutionContext.envInfo.profile.get("solution")?.set("tenantId", inputs.tenantId);
    else solutionContext.envInfo.profile.get("solution")?.delete("tenantId");
    if (inputs.subscriptionId)
      solutionContext.envInfo.profile.get("solution")?.set("subscriptionId", inputs.subscriptionId);
    else solutionContext.envInfo.profile.get("solution")?.delete("subscriptionId");
    return ok(Void);
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(isMultiEnvEnabled()),
    SolutionLoaderMW(defaultSolutionLoader),
    QuestionModelMW,
    ContextInjecterMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(),
  ])
  async grantPermission(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
    currentStage = Stage.grantPermission;
    return await ctx!.solution!.grantPermission!(ctx!.solutionContext!);
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(isMultiEnvEnabled()),
    SolutionLoaderMW(defaultSolutionLoader),
    QuestionModelMW,
    ContextInjecterMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(),
  ])
  async checkPermission(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
    currentStage = Stage.checkPermission;
    return await ctx!.solution!.checkPermission!(ctx!.solutionContext!);
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(isMultiEnvEnabled()),
    SolutionLoaderMW(defaultSolutionLoader),
    QuestionModelMW,
    ContextInjecterMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(),
  ])
  async listCollaborator(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
    currentStage = Stage.listCollaborator;
    return await ctx!.solution!.listCollaborator!(ctx!.solutionContext!);
  }

  async _getQuestionsForUserTask(
    ctx: SolutionContext,
    solution: Solution,
    func: FunctionRouter,
    inputs: Inputs
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    const namespace = func.namespace;
    const array = namespace ? namespace.split("/") : [];
    if (namespace && "" !== namespace && array.length > 0 && solution.getQuestionsForUserTask) {
      ctx!.answers = inputs;
      const res = await solution.getQuestionsForUserTask!(func, ctx!);
      if (res.isOk()) {
        if (res.value) {
          const node = res.value.trim();
          return ok(node);
        }
      }
      return res;
    }
    return err(FunctionRouterError(func));
  }

  async _getQuestionsForCreateProject(
    inputs: Inputs
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    const node = new QTreeNode(getCreateNewOrFromSampleQuestion(inputs.platform));
    // create new
    const createNew = new QTreeNode({ type: "group" });
    node.addChild(createNew);
    createNew.condition = { equals: ScratchOptionYes.id };
    const globalSolutions: Solution[] = await defaultSolutionLoader.loadGlobalSolutions(inputs);
    const solutionNames: string[] = globalSolutions.map((s) => s.name);
    const selectSolution: SingleSelectQuestion = QuestionSelectSolution;
    selectSolution.staticOptions = solutionNames;
    const solutionSelectNode = new QTreeNode(selectSolution);
    createNew.addChild(solutionSelectNode);
    const solutionContext = await newSolutionContext(this.tools, inputs);
    for (const v of globalSolutions) {
      if (v.getQuestions) {
        const res = await v.getQuestions(Stage.create, solutionContext);
        if (res.isErr()) return res;
        if (res.value) {
          const solutionNode = res.value as QTreeNode;
          solutionNode.condition = { equals: v.name };
          if (solutionNode.data) solutionSelectNode.addChild(solutionNode);
        }
      }
    }
    createNew.addChild(new QTreeNode(QuestionRootFolder));
    createNew.addChild(new QTreeNode(QuestionAppName));

    // create from sample
    const sampleNode = new QTreeNode(SampleSelect);
    node.addChild(sampleNode);
    sampleNode.condition = { equals: ScratchOptionNo.id };
    sampleNode.addChild(new QTreeNode(QuestionRootFolder));

    return ok(node.trim());
  }

  async _getQuestionsForMigrateV1Project(
    inputs: Inputs
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    const node = new QTreeNode({ type: "group" });
    const globalSolutions: Solution[] = await defaultSolutionLoader.loadGlobalSolutions(inputs);
    const solutionContext = await newSolutionContext(this.tools, inputs);

    for (const v of globalSolutions) {
      if (v.getQuestions) {
        const res = await v.getQuestions(Stage.migrateV1, solutionContext);
        if (res.isErr()) return res;
        if (res.value) {
          const solutionNode = res.value as QTreeNode;
          solutionNode.condition = { equals: v.name };
          if (solutionNode.data) node.addChild(solutionNode);
        }
      }
    }

    const defaultAppNameFunc = new QTreeNode(DefaultAppNameFunc);
    node.addChild(defaultAppNameFunc);

    const appNameQuestion = new QTreeNode(QuestionV1AppName);
    appNameQuestion.condition = {
      validFunc: (input: any) => (!input ? undefined : "App name is auto generated."),
    };
    defaultAppNameFunc.addChild(appNameQuestion);
    return ok(node.trim());
  }

  async _getQuestions(
    ctx: SolutionContext,
    solution: Solution,
    stage: Stage,
    inputs: Inputs
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    const node = new QTreeNode({ type: "group" });
    if (stage !== Stage.create) {
      const res = await solution.getQuestions(stage, ctx);
      if (res.isErr()) return res;
      if (res.value) {
        const child = res.value as QTreeNode;
        if (child.data) node.addChild(child);
      }
    }
    return ok(node.trim());
  }

  async createBasicFolderStructure(inputs: Inputs): Promise<Result<null, FxError>> {
    try {
      const appName = inputs[QuestionAppName.name] as string;
      await fs.writeFile(
        path.join(inputs.projectPath!, `package.json`),
        JSON.stringify(
          {
            name: appName,
            version: "0.0.1",
            description: "",
            author: "",
            scripts: {
              test: 'echo "Error: no test specified" && exit 1',
            },
            devDependencies: {
              "@microsoft/teamsfx-cli": "0.*",
            },
            license: "MIT",
          },
          null,
          4
        )
      );
      await fs.writeFile(
        path.join(inputs.projectPath!, `.gitignore`),
        `node_modules\n/.${ConfigFolderName}/*.env\n/.${ConfigFolderName}/*.userdata\n.DS_Store\n${ArchiveFolderName}\n${ArchiveLogFileName}`
      );
    } catch (e) {
      return err(WriteFileError(e));
    }
    return ok(null);
  }

  @hooks([
    ErrorHandlerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(false),
    ContextInjecterMW,
    EnvInfoWriterMW(),
  ])
  async encrypt(
    plaintext: string,
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<string, FxError>> {
    return ctx!.solutionContext!.cryptoProvider!.encrypt(plaintext);
  }

  @hooks([
    ErrorHandlerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(false),
    ContextInjecterMW,
    EnvInfoWriterMW(),
  ])
  async decrypt(
    ciphertext: string,
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<string, FxError>> {
    return ctx!.solutionContext!.cryptoProvider!.decrypt(ciphertext);
  }

  async buildArtifacts(inputs: Inputs): Promise<Result<Void, FxError>> {
    throw TaskNotSupportError(Stage.build);
  }

  @hooks([ErrorHandlerMW, ProjectSettingsLoaderMW, ContextInjecterMW])
  async createEnv(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    const projectSettings = ctx!.projectSettings;
    if (!isMultiEnvEnabled() || !projectSettings) {
      return ok(Void);
    }

    const core = ctx!.self as FxCore;
    const targetEnvName = await askNewEnvironment(ctx!, inputs);

    if (!targetEnvName) {
      return ok(Void);
    }

    if (targetEnvName) {
      const createEnvResult = await this.createEnvWithName(
        targetEnvName,
        projectSettings,
        inputs,
        core
      );
      if (createEnvResult.isErr()) {
        return createEnvResult;
      }
    }

    return ok(Void);
  }

  async createEnvWithName(
    targetEnvName: string,
    projectSettings: ProjectSettings,
    inputs: Inputs,
    core: FxCore
  ): Promise<Result<Void, FxError>> {
    const appName = projectSettings.appName;
    const newEnvConfig = environmentManager.newEnvConfigData(appName);
    const writeEnvResult = await environmentManager.writeEnvConfig(
      inputs.projectPath!,
      newEnvConfig,
      targetEnvName
    );
    if (writeEnvResult.isErr()) {
      return err(writeEnvResult.error);
    }
    core.tools.logProvider.debug(
      `[core] persist ${targetEnvName} env profile to path ${
        writeEnvResult.value
      }: ${JSON.stringify(newEnvConfig)}`
    );

    if (isArmSupportEnabled()) {
      const solutionContext: SolutionContext = {
        projectSettings,
        envInfo: newEnvInfo(targetEnvName, newEnvConfig),
        root: inputs.projectPath || "",
        ...core.tools,
        ...core.tools.tokenProvider,
        answers: inputs,
        cryptoProvider: new LocalCrypto(projectSettings.projectId),
        permissionRequestProvider: inputs.projectPath
          ? new PermissionRequestFileProvider(inputs.projectPath)
          : undefined,
      };

      await getParameterJson(solutionContext);
    }

    return ok(Void);
  }

  @hooks([
    ErrorHandlerMW,
    ProjectSettingsLoaderMW,
    SolutionLoaderMW(defaultSolutionLoader),
    ContextInjecterMW,
    ProjectSettingsWriterMW,
  ])
  async activateEnv(
    env: string,
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<Void, FxError>> {
    if (!isMultiEnvEnabled() || !ctx!.projectSettings) {
      return ok(Void);
    }

    const envConfigs = await environmentManager.listEnvConfigs(inputs.projectPath!);

    if (envConfigs.isErr()) {
      return envConfigs;
    }

    if (envConfigs.isErr() && envConfigs.value.indexOf(env) < 0) {
      return err(NonExistEnvNameError(env));
    }

    ctx!.projectSettings.activeEnvironment = env;
    const core = ctx!.self as FxCore;
    const solutionContext = await loadSolutionContext(
      core.tools,
      inputs,
      ctx!.projectSettings,
      ctx!.projectIdMissing,
      env
    );

    if (!solutionContext.isErr()) {
      ctx!.solutionContext = solutionContext.value;
    }

    this.tools.ui.showMessage("info", `[${env}] is activated.`, false);
    return ok(Void);
  }

  async removeEnv(inputs: Inputs): Promise<Result<Void, FxError>> {
    throw TaskNotSupportError(Stage.removeEnv);
  }
  async switchEnv(inputs: Inputs): Promise<Result<Void, FxError>> {
    throw TaskNotSupportError(Stage.switchEnv);
  }
}

export async function downloadSample(
  fxcore: FxCore | FxCoreV2,
  inputs: Inputs
): Promise<Result<string, FxError>> {
  const folder = inputs[QuestionRootFolder.name] as string;
  const sample = inputs[CoreQuestionNames.Samples] as OptionItem;
  if (sample && sample.data && folder) {
    const url = sample.data as string;
    const sampleId = sample.id;
    const sampleAppPath = path.resolve(folder, sampleId);
    if ((await fs.pathExists(sampleAppPath)) && (await fs.readdir(sampleAppPath)).length > 0) {
      return err(ProjectFolderExistError(sampleAppPath));
    }

    let fetchRes: AxiosResponse<any> | undefined;
    const task1: RunnableTask<Void> = {
      name: `Download code from '${url}'`,
      run: async (...args: any): Promise<Result<Void, FxError>> => {
        try {
          sendTelemetryEvent(Component.core, TelemetryEvent.DownloadSampleStart, {
            [TelemetryProperty.SampleAppName]: sample.id,
            module: "fx-core",
          });
          fetchRes = await fetchCodeZip(url);
          if (fetchRes !== undefined) {
            sendTelemetryEvent(Component.core, TelemetryEvent.DownloadSample, {
              [TelemetryProperty.SampleAppName]: sample.id,
              [TelemetryProperty.Success]: TelemetrySuccess.Yes,
              module: "fx-core",
            });
            return ok(Void);
          } else return err(FetchSampleError());
        } catch (e) {
          sendTelemetryErrorEvent(Component.core, TelemetryEvent.DownloadSample, assembleError(e), {
            [TelemetryProperty.SampleAppName]: sample.id,
            module: "fx-core",
          });
          return err(assembleError(e));
        }
      },
    };

    const task2: RunnableTask<Void> = {
      name: "Save and unzip package",
      run: async (...args: any): Promise<Result<Void, FxError>> => {
        if (fetchRes) {
          await saveFilesRecursively(new AdmZip(fetchRes.data), sampleId, folder);
        }
        return ok(Void);
      },
    };
    const task3: RunnableTask<Void> = {
      name: "post process",
      run: async (...args: any): Promise<Result<Void, FxError>> => {
        await downloadSampleHook(sampleId, sampleAppPath);
        return ok(Void);
      },
    };
    const group = new GroupOfTasks<Void>([task1, task2, task3], {
      sequential: true,
      fastFail: true,
    });
    const runRes = await fxcore.tools.ui.runWithProgress(group, {
      showProgress: true,
      cancellable: false,
    });
    if (runRes.isOk()) {
      return ok(sampleAppPath);
    } else {
      return err(runRes.error);
    }
  }
  return err(InvalidInputError(`invalid answer for '${CoreQuestionNames.Samples}'`, inputs));
}

export * from "./error";
export * from "./tools";
