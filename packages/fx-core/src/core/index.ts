// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { HookContext, hooks } from "@feathersjs/hooks";
import {
  AppPackageFolderName,
  ArchiveFolderName,
  ArchiveLogFileName,
  assembleError,
  BuildFolderName,
  ConfigFolderName,
  CoreCallbackEvent,
  CoreCallbackFunc,
  err,
  Func,
  FunctionRouter,
  FxError,
  InputConfigsFolderName,
  Inputs,
  Json,
  LogProvider,
  ok,
  OptionItem,
  Platform,
  ProjectConfig,
  ProjectSettings,
  QTreeNode,
  Result,
  Solution,
  SolutionConfig,
  SolutionContext,
  Stage,
  StatesFolderName,
  TelemetryReporter,
  Tools,
  UserCancelError,
  v2,
  v3,
  Void,
} from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import * as fs from "fs-extra";
import * as jsonschema from "jsonschema";
import { assign } from "lodash";
import * as path from "path";
import { Container } from "typedi";
import * as uuid from "uuid";
import { environmentManager, sampleProvider } from "..";
import { FeatureFlagName } from "../common/constants";
import { globalStateUpdate } from "../common/globalState";
import { localSettingsFileName } from "../common/localSettingsProvider";
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
  getRootDirectory,
  isMultiEnvEnabled,
  mapToJson,
  saveFilesRecursively,
} from "../common/tools";
import { PluginNames } from "../plugins";
import { MessageExtensionItem } from "../plugins/solution/fx-solution/question";
import { getAllV2ResourcePlugins } from "../plugins/solution/fx-solution/ResourcePluginContainer";
import {
  BuiltInResourcePluginNames,
  BuiltInScaffoldPluginNames,
  BuiltInSolutionNames,
} from "../plugins/solution/fx-solution/v3/constants";
import { CallbackRegistry } from "./callback";
import { LocalCrypto } from "./crypto";
import {
  ArchiveProjectError,
  ArchiveUserFileError,
  CopyFileError,
  FetchSampleError,
  FunctionRouterError,
  InvalidInputError,
  LoadSolutionError,
  MigrateNotImplementError,
  NonExistEnvNameError,
  ObjectIsUndefinedError,
  ProjectFolderExistError,
  ProjectFolderInvalidError,
  ProjectFolderNotExistError,
  TaskNotSupportError,
  WriteFileError,
} from "./error";
import { ConcurrentLockerMW } from "./middleware/concurrentLocker";
import { ContextInjectorMW } from "./middleware/contextInjector";
import {
  askNewEnvironment,
  EnvInfoLoaderMW,
  loadSolutionContext,
  upgradeDefaultFunctionName,
  upgradeProgrammingLanguage,
} from "./middleware/envInfoLoader";
import { EnvInfoWriterMW } from "./middleware/envInfoWriter";
import { EnvInfoWriterMW_V3 } from "./middleware/envInfoWriterV3";
import { ErrorHandlerMW } from "./middleware/errorHandler";
import { LocalSettingsLoaderMW } from "./middleware/localSettingsLoader";
import { LocalSettingsWriterMW } from "./middleware/localSettingsWriter";
import { MigrateConditionHandlerMW } from "./middleware/migrateConditionHandler";
import { ProjectMigratorMW } from "./middleware/projectMigrator";
import {
  loadProjectSettings,
  newSolutionContext,
  ProjectSettingsLoaderMW,
} from "./middleware/projectSettingsLoader";
import { ProjectSettingsWriterMW } from "./middleware/projectSettingsWriter";
import { ProjectUpgraderMW } from "./middleware/projectUpgrader";
import {
  getQuestionsForAddModule,
  getQuestionsForAddResource,
  getQuestionsForCreateProjectV2,
  getQuestionsForCreateProjectV3,
  getQuestionsForDeploy,
  getQuestionsForInit,
  getQuestionsForLocalProvision,
  getQuestionsForMigrateV1Project,
  getQuestionsForProvision,
  getQuestionsForPublish,
  getQuestionsForScaffold,
  getQuestionsForUserTaskV2,
  getQuestionsV2,
  QuestionModelMW,
} from "./middleware/questionModel";
import { SolutionLoaderMW } from "./middleware/solutionLoader";
import {
  BotOptionItem,
  CoreQuestionNames,
  DefaultAppNameFunc,
  ProjectNamePattern,
  QuestionAppName,
  QuestionRootFolder,
  QuestionV1AppName,
  ScratchOptionNo,
  TabOptionItem,
  TabSPFxItem,
} from "./question";
import {
  getAllSolutionPlugins,
  getAllSolutionPluginsV2,
  getSolutionPluginByName,
  getSolutionPluginV2ByName,
} from "./SolutionPluginContainer";
import { newEnvInfo } from "./tools";
import { SupportV1ConditionMW } from "./middleware/supportV1ConditionHandler";
import { ProjectSettingsLoaderMW_V3 } from "./middleware/projectSettingsLoaderV3";
import { SolutionLoaderMW_V3 } from "./middleware/solutionLoaderV3";
import { EnvInfoLoaderMW_V3 } from "./middleware/envInfoLoaderV3";
// TODO: For package.json,
// use require instead of import because of core building/packaging method.
// Using import will cause the build folder structure to change.
const corePackage = require("../../package.json");

export interface CoreHookContext extends HookContext {
  projectSettings?: ProjectSettings;
  solutionContext?: SolutionContext;
  solution?: Solution;
  //for v2 api
  contextV2?: v2.Context;
  solutionV2?: v2.SolutionPlugin;
  envInfoV2?: v2.EnvInfoV2;
  localSettings?: Json;

  //for v3
  envInfoV3?: v3.EnvInfoV3;
  solutionV3?: v3.ISolution;
}

function featureFlagEnabled(flagName: string): boolean {
  const flag = process.env[flagName];
  if (flag !== undefined && flag.toLowerCase() === "true") {
    return true;
  } else {
    return false;
  }
}

export function isV3() {
  return featureFlagEnabled(FeatureFlagName.APIV3);
}

// On VS calling CLI, interactive questions need to be skipped.
export function isVsCallingCli() {
  return featureFlagEnabled(FeatureFlagName.VSCallingCLI);
}

export let Logger: LogProvider;
export let telemetryReporter: TelemetryReporter | undefined;
export let currentStage: Stage;
export let TOOLS: Tools;
export function setTools(tools: Tools) {
  TOOLS = tools;
}
export class FxCore implements v3.ICore {
  tools: Tools;
  isFromSample?: boolean;
  settingsVersion?: string;

  constructor(tools: Tools) {
    this.tools = tools;
    TOOLS = tools;
    Logger = tools.logProvider;
    telemetryReporter = tools.telemetryReporter;
  }

  /**
   * @todo this's a really primitive implement. Maybe could use Subscription Model to
   * refactor later.
   */
  public on(event: CoreCallbackEvent, callback: CoreCallbackFunc): void {
    return CallbackRegistry.set(event, callback);
  }

  async createProject(inputs: Inputs): Promise<Result<string, FxError>> {
    if (isV3()) {
      return this.createProjectV3(inputs);
    } else {
      return this.createProjectV2(inputs);
    }
  }
  @hooks([
    ErrorHandlerMW,
    SupportV1ConditionMW(true),
    QuestionModelMW,
    ContextInjectorMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(true),
  ])
  async createProjectV2(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<string, FxError>> {
    if (!ctx) {
      return err(new ObjectIsUndefinedError("ctx for createProject"));
    }
    currentStage = Stage.create;
    inputs.stage = Stage.create;
    let folder = inputs[QuestionRootFolder.name] as string;
    if (inputs.platform === Platform.VSCode) {
      folder = getRootDirectory();
      try {
        await fs.ensureDir(folder);
      } catch (e) {
        throw ProjectFolderInvalidError(folder);
      }
    }
    const scratch = inputs[CoreQuestionNames.CreateFromScratch] as string;
    let projectPath: string;
    let globalStateDescription = "openReadme";
    const multiEnv = isMultiEnvEnabled();
    if (scratch === ScratchOptionNo.id) {
      // create from sample
      const downloadRes = await downloadSample(inputs, ctx);
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
      inputs.projectPath = projectPath;
      const folderExist = await fs.pathExists(projectPath);
      if (folderExist) {
        return err(ProjectFolderExistError(projectPath));
      }
      await fs.ensureDir(projectPath);
      await fs.ensureDir(path.join(projectPath, `.${ConfigFolderName}`));
      await fs.ensureDir(
        path.join(
          projectPath,
          multiEnv ? path.join("templates", `${AppPackageFolderName}`) : `${AppPackageFolderName}`
        )
      );
      const basicFolderRes = await createBasicFolderStructure(inputs);
      if (basicFolderRes.isErr()) {
        return err(basicFolderRes.error);
      }

      const projectSettings: ProjectSettings = {
        appName: appName,
        projectId: inputs.projectId ? inputs.projectId : uuid.v4(),
        solutionSettings: {
          name: "",
          version: "1.0.0",
        },
        version: getProjectSettingsVersion(),
        isFromSample: false,
      };
      ctx.projectSettings = projectSettings;
      if (multiEnv) {
        const createEnvResult = await this.createEnvWithName(
          environmentManager.getDefaultEnvName(),
          projectSettings,
          inputs
        );
        if (createEnvResult.isErr()) {
          return err(createEnvResult.error);
        }
      }

      const solution = await getSolutionPluginV2ByName(inputs[CoreQuestionNames.Solution]);
      if (!solution) {
        return err(new LoadSolutionError());
      }
      ctx.solutionV2 = solution;
      projectSettings.solutionSettings.name = solution.name;
      const contextV2 = createV2Context(projectSettings);
      ctx.contextV2 = contextV2;
      const scaffoldSourceCodeRes = await solution.scaffoldSourceCode(contextV2, inputs);
      if (scaffoldSourceCodeRes.isErr()) {
        return err(scaffoldSourceCodeRes.error);
      }
      const generateResourceTemplateRes = await solution.generateResourceTemplate(
        contextV2,
        inputs
      );
      if (generateResourceTemplateRes.isErr()) {
        return err(generateResourceTemplateRes.error);
      }
      // ctx.provisionInputConfig = generateResourceTemplateRes.value;
      if (multiEnv) {
        if (solution.createEnv) {
          inputs.copy = false;
          const createEnvRes = await solution.createEnv(contextV2, inputs);
          if (createEnvRes.isErr()) {
            return err(createEnvRes.error);
          }
        }
      } else {
        //TODO lagacy env.default.json
        const state: Json = { solution: {} };
        for (const plugin of getAllV2ResourcePlugins()) {
          state[plugin.name] = {};
        }
        state[PluginNames.LDEBUG]["trustDevCert"] = "true";
        ctx.envInfoV2 = {
          envName: environmentManager.getDefaultEnvName(),
          config: {},
          state: state,
        };
      }
    }

    if (inputs.platform === Platform.VSCode) {
      await globalStateUpdate(globalStateDescription, true);
    }
    return ok(projectPath);
  }
  @hooks([
    ErrorHandlerMW,
    SupportV1ConditionMW(true),
    QuestionModelMW,
    ContextInjectorMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW_V3(true),
  ])
  async createProjectV3(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<string, FxError>> {
    if (!ctx) {
      return err(new ObjectIsUndefinedError("ctx for createProject"));
    }
    currentStage = Stage.create;
    inputs.stage = Stage.create;
    let folder = inputs[QuestionRootFolder.name] as string;
    if (inputs.platform === Platform.VSCode || inputs.platform === Platform.VS) {
      folder = getRootDirectory();
      try {
        await fs.ensureDir(folder);
      } catch (e) {
        throw ProjectFolderInvalidError(folder);
      }
    }
    const scratch = inputs[CoreQuestionNames.CreateFromScratch] as string;
    let projectPath: string;
    let globalStateDescription = "openReadme";
    if (scratch === ScratchOptionNo.id) {
      // create from sample
      const downloadRes = await downloadSample(inputs, ctx);
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
      inputs.projectPath = projectPath;
      const folderExist = await fs.pathExists(projectPath);
      if (folderExist) {
        return err(ProjectFolderExistError(projectPath));
      }
      await fs.ensureDir(projectPath);
      await fs.ensureDir(path.join(projectPath, `.${ConfigFolderName}`));

      let capabilities = inputs[CoreQuestionNames.Capabilities] as string[];

      let projectType = "";
      if (capabilities.includes(TabSPFxItem.id)) projectType = "spfx";
      else if (capabilities.includes(TabOptionItem.id) && capabilities.length === 1)
        projectType = "tab";
      else if (
        (capabilities.includes(BotOptionItem.id) ||
          capabilities.includes(MessageExtensionItem.id)) &&
        !capabilities.includes(TabOptionItem.id)
      )
        projectType = "bot";
      else if (
        (capabilities.includes(BotOptionItem.id) ||
          capabilities.includes(MessageExtensionItem.id)) &&
        capabilities.includes(TabOptionItem.id)
      )
        projectType = "tab+bot";

      const programmingLanguage = inputs[CoreQuestionNames.ProgrammingLanguage] as string;
      // const solution = capabilities.includes(TabSPFxItem.id)
      //   ? BuiltInSolutionNames.spfx
      //   : BuiltInSolutionNames.azure;

      // init
      const initInputs: v2.InputsWithProjectPath & { solution?: string } = {
        ...inputs,
        projectPath: projectPath,
        // solution: solution,
      };
      const initRes = await this._init(initInputs, ctx);
      if (initRes.isErr()) {
        return err(initRes.error);
      }

      // addModule, scaffold and addResource
      if (inputs.platform === Platform.VS) {
        // addModule
        const addModuleInputs: v2.InputsWithProjectPath & { capabilities?: string[] } = {
          ...inputs,
          projectPath: projectPath,
          capabilities: capabilities,
        };
        const addModuleRes = await this._addModule(addModuleInputs, ctx);
        if (addModuleRes.isErr()) {
          return err(addModuleRes.error);
        }
        // addResource
        const addResourceInputs: v2.InputsWithProjectPath & { module?: string; resource?: string } =
          {
            ...inputs,
            projectPath: projectPath,
            module: "0",
            resource: BuiltInResourcePluginNames.webApp, //TODO
          };
        const addResourceRes = await this._addResource(addResourceInputs, ctx);
        if (addResourceRes.isErr()) {
          return err(addResourceRes.error);
        }
        // scaffold
        let templateName = "";
        if (projectType === "tab") templateName = "BlazorTab";
        else if (projectType === "bot") templateName = "BlazorBot";
        else if (projectType === "tabbot") templateName = "BlazorTabBot";
        const scaffoldInputs: v2.InputsWithProjectPath & {
          module?: string;
          template?: OptionItem;
        } = {
          ...inputs,
          projectPath: projectPath,
          module: "0",
          template: {
            id: `${BuiltInScaffoldPluginNames.blazor}/${templateName}`,
            label: `${BuiltInScaffoldPluginNames.blazor}/${templateName}`,
            data: {
              pluginName: BuiltInScaffoldPluginNames.blazor,
              templateName: templateName,
            },
          },
        };
        const scaffoldRes = await this._scaffold(scaffoldInputs, ctx);
        if (scaffoldRes.isErr()) {
          return err(scaffoldRes.error);
        }
      } else {
        if (capabilities.includes(TabOptionItem.id) || capabilities.includes(TabSPFxItem.id)) {
          const addModuleInputs: v2.InputsWithProjectPath & { capabilities?: string[] } = {
            ...inputs,
            projectPath: projectPath,
            capabilities: capabilities.includes(TabOptionItem.id)
              ? [TabOptionItem.id]
              : [TabSPFxItem.id],
          };
          const addModuleRes = await this._addModule(addModuleInputs, ctx);
          if (addModuleRes.isErr()) {
            return err(addModuleRes.error);
          }
          // addResource
          const addResourceInputs: v2.InputsWithProjectPath & {
            module?: string;
            resource?: string;
          } = {
            ...inputs,
            projectPath: projectPath,
            module: "0",
            resource: capabilities.includes(TabOptionItem.id)
              ? BuiltInResourcePluginNames.storage
              : BuiltInResourcePluginNames.spfx, //TODO
          };
          const addResourceRes = await this._addResource(addResourceInputs, ctx);
          if (addResourceRes.isErr()) {
            return err(addResourceRes.error);
          }
          // scaffold
          const pluginName = capabilities.includes(TabOptionItem.id)
            ? BuiltInScaffoldPluginNames.tab
            : BuiltInScaffoldPluginNames.spfx;
          const templateName = capabilities.includes(TabOptionItem.id)
            ? programmingLanguage === "javascript"
              ? "ReactTab_JS"
              : "ReactTab_TS"
            : "SPFxTab";
          const scaffoldInputs: v2.InputsWithProjectPath & {
            module?: string;
            template?: OptionItem;
          } = {
            ...inputs,
            projectPath: projectPath,
            module: "0",
            template: {
              id: `${pluginName}/${templateName}`,
              label: `${pluginName}/${templateName}`,
              data: {
                pluginName: pluginName,
                templateName: templateName, //TODO
              },
            },
          };
          const scaffoldRes = await this._scaffold(scaffoldInputs, ctx);
          if (scaffoldRes.isErr()) {
            return err(scaffoldRes.error);
          }
        }
        capabilities = capabilities.filter((c) => c !== TabOptionItem.id && c !== TabSPFxItem.id);
        if (capabilities.length > 0) {
          const addModuleInputs: v2.InputsWithProjectPath & { capabilities?: string[] } = {
            ...inputs,
            projectPath: projectPath,
            capabilities: capabilities,
          };
          const addModuleRes = await this._addModule(addModuleInputs, ctx);
          if (addModuleRes.isErr()) {
            return err(addModuleRes.error);
          }
          // addResource
          const addResourceInputs: v2.InputsWithProjectPath & {
            module?: string;
            resource?: string;
          } = {
            ...inputs,
            projectPath: projectPath,
            module: "1",
            resource: BuiltInResourcePluginNames.bot, //TODO
          };
          const addResourceRes = await this._addResource(addResourceInputs, ctx);
          if (addResourceRes.isErr()) {
            return err(addResourceRes.error);
          }
          // scaffold
          const templateName =
            programmingLanguage === "javascript" ? "NodejsBot_JS" : "NodejsBot_TS";
          const scaffoldInputs: v2.InputsWithProjectPath & {
            module?: string;
            template?: OptionItem;
          } = {
            ...inputs,
            projectPath: projectPath,
            module: "1",
            resource: BuiltInScaffoldPluginNames.bot, //TODO
            template: {
              id: `${BuiltInScaffoldPluginNames.bot}/${templateName}`,
              label: `${BuiltInScaffoldPluginNames.bot}/${templateName}`,
              data: {
                pluginName: BuiltInScaffoldPluginNames.bot,
                templateName: templateName, //TODO
              },
            }, //TODO
          };
          const scaffoldRes = await this._scaffold(scaffoldInputs, ctx);
          if (scaffoldRes.isErr()) {
            return err(scaffoldRes.error);
          }
        }
      }
    }
    if (inputs.platform === Platform.VSCode) {
      await globalStateUpdate(globalStateDescription, true);
    }
    return ok(projectPath);
  }
  @hooks([
    ErrorHandlerMW,
    SupportV1ConditionMW(true),
    MigrateConditionHandlerMW,
    QuestionModelMW,
    ContextInjectorMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(true),
  ])
  async migrateV1Project(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<string, FxError>> {
    currentStage = Stage.migrateV1;
    inputs.stage = Stage.migrateV1;
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

    const solution = await getAllSolutionPlugins()[0];
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
      cryptoProvider: new LocalCrypto(projectSettings.projectId),
    };

    const archiveResult = await this.archive(projectPath);
    if (archiveResult.isErr()) {
      return err(archiveResult.error);
    }

    await fs.ensureDir(projectPath);
    await fs.ensureDir(path.join(projectPath, `.${ConfigFolderName}`));

    const createResult = await createBasicFolderStructure(inputs);
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
    ctx!.projectSettings = projectSettings;

    if (inputs.platform === Platform.VSCode) {
      await globalStateUpdate(globalStateDescription, true);
    }
    this._setEnvInfoV2(ctx);
    return ok(projectPath);
  }

  async archive(projectPath: string): Promise<Result<Void, FxError>> {
    try {
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
          return err(ArchiveUserFileError(fileName, e.message));
        }

        await archiveLog(
          projectPath,
          `'${fileName}' has been moved to '${ArchiveFolderName}' folder.`
        );
      }
      return ok(Void);
    } catch (e: any) {
      return err(ArchiveProjectError(e.message));
    }
  }

  /**
   * switch to different versions of provisionResources
   */
  async provisionResources(inputs: Inputs): Promise<Result<Void, FxError>> {
    if (isV3()) {
      return this.provisionResourcesV3(inputs);
    } else {
      return this.provisionResourcesV2(inputs);
    }
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(false),
    ProjectMigratorMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(false),
    SolutionLoaderMW,
    QuestionModelMW,
    ContextInjectorMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(),
  ])
  async provisionResourcesV2(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<Void, FxError>> {
    currentStage = Stage.provision;
    inputs.stage = Stage.provision;
    if (!ctx || !ctx.solutionV2 || !ctx.contextV2 || !ctx.envInfoV2) {
      return err(new ObjectIsUndefinedError("Provision input stuff"));
    }
    const envInfo = ctx.envInfoV2;
    const result = await ctx.solutionV2.provisionResources(
      ctx.contextV2,
      inputs,
      envInfo,
      this.tools.tokenProvider
    );
    if (result.kind === "success") {
      ctx.envInfoV2.state = assign(ctx.envInfoV2.state, result.output);
      return ok(Void);
    } else if (result.kind === "partialSuccess") {
      ctx.envInfoV2.state = assign(ctx.envInfoV2.state, result.output);
      return err(result.error);
    } else {
      return err(result.error);
    }
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(false),
    ProjectMigratorMW,
    ProjectSettingsLoaderMW_V3,
    EnvInfoLoaderMW_V3(false),
    SolutionLoaderMW_V3,
    QuestionModelMW,
    ContextInjectorMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW_V3(),
  ])
  async provisionResourcesV3(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<Void, FxError>> {
    currentStage = Stage.provision;
    inputs.stage = Stage.provision;
    if (
      ctx &&
      ctx.solutionV3 &&
      ctx.contextV2 &&
      ctx.envInfoV3 &&
      ctx.solutionV3.provisionResources
    ) {
      const res = await ctx.solutionV3.provisionResources(
        ctx.contextV2,
        inputs as v2.InputsWithProjectPath,
        ctx.envInfoV3,
        TOOLS.tokenProvider
      );
      if (res.isOk()) {
        ctx.envInfoV3 = res.value;
      }
      return res;
    }
    return ok(Void);
  }
  async deployArtifacts(inputs: Inputs): Promise<Result<Void, FxError>> {
    if (isV3()) return this.deployArtifactsV3(inputs);
    else return this.deployArtifactsV2(inputs);
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(false),
    ProjectMigratorMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(false),
    SolutionLoaderMW,
    QuestionModelMW,
    ContextInjectorMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(),
  ])
  async deployArtifactsV2(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    currentStage = Stage.deploy;
    inputs.stage = Stage.deploy;
    if (!ctx || !ctx.solutionV2 || !ctx.contextV2 || !ctx.envInfoV2) {
      const name = undefinedName(
        [ctx, ctx?.solutionV2, ctx?.contextV2, ctx?.envInfoV2],
        ["ctx", "ctx.solutionV2", "ctx.contextV2", "ctx.envInfoV2"]
      );
      return err(new ObjectIsUndefinedError(`Deploy input stuff: ${name}`));
    }

    if (ctx.solutionV2.deploy)
      return await ctx.solutionV2.deploy(
        ctx.contextV2,
        inputs,
        ctx.envInfoV2,
        this.tools.tokenProvider
      );
    else return ok(Void);
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(false),
    ProjectMigratorMW,
    ProjectSettingsLoaderMW_V3,
    EnvInfoLoaderMW_V3(false),
    SolutionLoaderMW_V3,
    QuestionModelMW,
    ContextInjectorMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW_V3(),
  ])
  async deployArtifactsV3(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    currentStage = Stage.deploy;
    inputs.stage = Stage.deploy;
    if (ctx && ctx.solutionV3 && ctx.contextV2 && ctx.envInfoV3 && ctx.solutionV3.deploy) {
      const res = await ctx.solutionV3.deploy(
        ctx.contextV2,
        inputs as v2.InputsWithProjectPath & { modules: string[] },
        ctx.envInfoV3,
        TOOLS.tokenProvider
      );
      return res;
    }
    return ok(Void);
  }
  async localDebug(inputs: Inputs): Promise<Result<Void, FxError>> {
    if (isV3()) return this.localDebugV3(inputs);
    else return this.localDebugV2(inputs);
  }
  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(true),
    ProjectMigratorMW,
    ProjectUpgraderMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(true),
    LocalSettingsLoaderMW,
    SolutionLoaderMW,
    QuestionModelMW,
    ContextInjectorMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(true),
    LocalSettingsWriterMW,
  ])
  async localDebugV2(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    currentStage = Stage.debug;
    inputs.stage = Stage.debug;
    if (!ctx || !ctx.solutionV2 || !ctx.contextV2) {
      const name = undefinedName(
        [ctx, ctx?.solutionV2, ctx?.contextV2],
        ["ctx", "ctx.solutionV2", "ctx.contextV2"]
      );
      return err(new ObjectIsUndefinedError(`localDebug input stuff (${name})`));
    }
    if (!ctx.localSettings) ctx.localSettings = {};
    if (ctx.solutionV2.provisionLocalResource) {
      const res = await ctx.solutionV2.provisionLocalResource(
        ctx.contextV2,
        inputs,
        ctx.localSettings,
        this.tools.tokenProvider
      );
      if (res.kind === "success") {
        ctx.localSettings = res.output;
        return ok(Void);
      } else if (res.kind === "partialSuccess") {
        ctx.localSettings = res.output;
        return err(res.error);
      } else {
        return err(res.error);
      }
    } else {
      return ok(Void);
    }
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(true),
    ProjectMigratorMW,
    ProjectUpgraderMW,
    ProjectSettingsLoaderMW_V3,
    LocalSettingsLoaderMW,
    SolutionLoaderMW_V3,
    QuestionModelMW,
    ContextInjectorMW,
    ProjectSettingsWriterMW,
    LocalSettingsWriterMW,
  ])
  async localDebugV3(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    currentStage = Stage.debug;
    inputs.stage = Stage.debug;
    if (
      ctx &&
      ctx.solutionV3 &&
      ctx.contextV2 &&
      ctx.localSettings &&
      ctx.solutionV3.provisionLocalResources
    ) {
      const res = await ctx.solutionV3.provisionLocalResources(
        ctx.contextV2,
        inputs as v2.InputsWithProjectPath,
        ctx.localSettings,
        TOOLS.tokenProvider
      );
      if (res.isOk()) {
        ctx.localSettings = res.value;
      }
      return res;
    }
    return ok(Void);
  }

  _setEnvInfoV2(ctx?: CoreHookContext) {
    if (ctx && ctx.solutionContext) {
      //workaround, compatible to api v2
      ctx.envInfoV2 = {
        envName: ctx.solutionContext.envInfo.envName,
        config: ctx.solutionContext.envInfo.config,
        state: {},
      };
      ctx.envInfoV2.state = mapToJson(ctx.solutionContext.envInfo.state);
    }
  }
  async publishApplication(inputs: Inputs): Promise<Result<Void, FxError>> {
    if (isV3()) return this.publishApplicationV3(inputs);
    else return this.publishApplicationV2(inputs);
  }
  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(false),
    ProjectMigratorMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(false),
    SolutionLoaderMW,
    QuestionModelMW,
    ContextInjectorMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(),
  ])
  async publishApplicationV2(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<Void, FxError>> {
    currentStage = Stage.publish;
    inputs.stage = Stage.publish;
    if (!ctx || !ctx.solutionV2 || !ctx.contextV2 || !ctx.envInfoV2) {
      const name = undefinedName(
        [ctx, ctx?.solutionV2, ctx?.contextV2, ctx?.envInfoV2],
        ["ctx", "ctx.solutionV2", "ctx.contextV2", "ctx.envInfoV2"]
      );
      return err(new ObjectIsUndefinedError(`publish input stuff: ${name}`));
    }
    return await ctx.solutionV2.publishApplication(
      ctx.contextV2,
      inputs,
      ctx.envInfoV2,
      this.tools.tokenProvider.appStudioToken
    );
  }
  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(false),
    ProjectMigratorMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(false),
    SolutionLoaderMW,
    QuestionModelMW,
    ContextInjectorMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(),
  ])
  async publishApplicationV3(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<Void, FxError>> {
    currentStage = Stage.publish;
    inputs.stage = Stage.publish;
    if (
      ctx &&
      ctx.solutionV3 &&
      ctx.contextV2 &&
      ctx.envInfoV3 &&
      ctx.solutionV3.publishApplication
    ) {
      const res = await ctx.solutionV3.publishApplication(
        ctx.contextV2,
        inputs as v2.InputsWithProjectPath,
        ctx.envInfoV3,
        TOOLS.tokenProvider.appStudioToken
      );
      return res;
    }
    return ok(Void);
  }
  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(false),
    ProjectMigratorMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(false),
    LocalSettingsLoaderMW,
    SolutionLoaderMW,
    QuestionModelMW,
    ContextInjectorMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW(),
  ])
  async executeUserTask(
    func: Func,
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<unknown, FxError>> {
    currentStage = Stage.userTask;
    inputs.stage = Stage.userTask;
    const namespace = func.namespace;
    const array = namespace ? namespace.split("/") : [];
    if ("" !== namespace && array.length > 0) {
      if (!ctx || !ctx.solutionV2 || !ctx.envInfoV2) {
        const name = undefinedName(
          [ctx, ctx?.solutionV2, ctx?.envInfoV2],
          ["ctx", "ctx.solutionV2", "ctx.envInfoV2"]
        );
        return err(new ObjectIsUndefinedError(`executeUserTask input stuff: ${name}`));
      }
      if (!ctx.contextV2) ctx.contextV2 = createV2Context(newProjectSettings());
      if (ctx.solutionV2.executeUserTask) {
        if (!ctx.localSettings) ctx.localSettings = {};
        const res = await ctx.solutionV2.executeUserTask(
          ctx.contextV2,
          inputs,
          func,
          ctx.localSettings,
          ctx.envInfoV2,
          this.tools.tokenProvider
        );
        return res;
      } else return err(FunctionRouterError(func));
    }
    return err(FunctionRouterError(func));
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(true),
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(true),
    SolutionLoaderMW,
    ContextInjectorMW,
    EnvInfoWriterMW(),
  ])
  async getQuestions(
    stage: Stage,
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    if (!ctx) return err(new ObjectIsUndefinedError("getQuestions input stuff"));
    inputs.stage = Stage.getQuestions;
    currentStage = Stage.getQuestions;
    if (stage === Stage.create) {
      delete inputs.projectPath;
      return await this._getQuestionsForCreateProjectV2(inputs);
    } else {
      const contextV2 = ctx.contextV2 ? ctx.contextV2 : createV2Context(newProjectSettings());
      const solutionV2 = ctx.solutionV2 ? ctx.solutionV2 : await getAllSolutionPluginsV2()[0];
      const envInfoV2 = ctx.envInfoV2
        ? ctx.envInfoV2
        : { envName: environmentManager.getDefaultEnvName(), config: {}, state: {} };
      inputs.stage = stage;
      return await this._getQuestions(contextV2, solutionV2, stage, inputs, envInfoV2);
    }
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(true),
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(true),
    SolutionLoaderMW,
    ContextInjectorMW,
    EnvInfoWriterMW(),
  ])
  async getQuestionsForUserTask(
    func: FunctionRouter,
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    if (!ctx) return err(new ObjectIsUndefinedError("getQuestionsForUserTask input stuff"));
    inputs.stage = Stage.getQuestions;
    currentStage = Stage.getQuestions;
    const contextV2 = ctx.contextV2 ? ctx.contextV2 : createV2Context(newProjectSettings());
    const solutionV2 = ctx.solutionV2 ? ctx.solutionV2 : await getAllSolutionPluginsV2()[0];
    const envInfoV2 = ctx.envInfoV2
      ? ctx.envInfoV2
      : { envName: environmentManager.getDefaultEnvName(), config: {}, state: {} };
    return await this._getQuestionsForUserTask(contextV2, solutionV2, func, inputs, envInfoV2);
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(true),
    ProjectMigratorMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(false),
    LocalSettingsLoaderMW,
    ContextInjectorMW,
  ])
  async getProjectConfig(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<ProjectConfig | undefined, FxError>> {
    if (!ctx) return err(new ObjectIsUndefinedError("getProjectConfig input stuff"));
    inputs.stage = Stage.getProjectConfig;
    currentStage = Stage.getProjectConfig;
    return ok({
      settings: ctx!.projectSettings,
      config: ctx!.solutionContext?.envInfo.state,
      localSettings: ctx!.solutionContext?.localSettings,
    });
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(false),
    ProjectMigratorMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(false),
    SolutionLoaderMW,
    QuestionModelMW,
    ContextInjectorMW,
  ])
  async grantPermission(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
    currentStage = Stage.grantPermission;
    inputs.stage = Stage.grantPermission;
    const projectPath = inputs.projectPath;
    if (!projectPath) {
      return err(new ObjectIsUndefinedError("projectPath"));
    }
    return ctx!.solutionV2!.grantPermission!(
      ctx!.contextV2!,
      { ...inputs, projectPath: projectPath },
      ctx!.envInfoV2!,
      this.tools.tokenProvider
    );
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(false),
    ProjectMigratorMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(false),
    SolutionLoaderMW,
    QuestionModelMW,
    ContextInjectorMW,
  ])
  async checkPermission(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
    currentStage = Stage.checkPermission;
    inputs.stage = Stage.checkPermission;
    const projectPath = inputs.projectPath;
    if (!projectPath) {
      return err(new ObjectIsUndefinedError("projectPath"));
    }
    return ctx!.solutionV2!.checkPermission!(
      ctx!.contextV2!,
      { ...inputs, projectPath: projectPath },
      ctx!.envInfoV2!,
      this.tools.tokenProvider
    );
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(true),
    ProjectMigratorMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(false),
    SolutionLoaderMW,
    QuestionModelMW,
    ContextInjectorMW,
  ])
  async listCollaborator(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
    currentStage = Stage.listCollaborator;
    inputs.stage = Stage.listCollaborator;
    const projectPath = inputs.projectPath;
    if (!projectPath) {
      return err(new ObjectIsUndefinedError("projectPath"));
    }
    return ctx!.solutionV2!.listCollaborator!(
      ctx!.contextV2!,
      { ...inputs, projectPath: projectPath },
      ctx!.envInfoV2!,
      this.tools.tokenProvider
    );
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(true),
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(true),
    SolutionLoaderMW,
    QuestionModelMW,
    ContextInjectorMW,
  ])
  async listAllCollaborators(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
    currentStage = Stage.listAllCollaborators;
    inputs.stage = Stage.listAllCollaborators;
    const projectPath = inputs.projectPath;
    if (!projectPath) {
      return err(new ObjectIsUndefinedError("projectPath"));
    }
    return ctx!.solutionV2!.listAllCollaborators!(
      ctx!.contextV2!,
      { ...inputs, projectPath: projectPath },
      ctx!.envInfoV2!,
      this.tools.tokenProvider
    );
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(false),
    ContextInjectorMW,
  ])
  async getSelectedEnv(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<string | undefined, FxError>> {
    if (!isMultiEnvEnabled()) {
      return err(new TaskNotSupportError("getSelectedEnv"));
    }
    return ok(ctx?.envInfoV2?.envName);
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(true),
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(true),
    ContextInjectorMW,
  ])
  async encrypt(
    plaintext: string,
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<string, FxError>> {
    if (!ctx) return err(new ObjectIsUndefinedError("ctx"));
    if (!ctx.contextV2) return err(new ObjectIsUndefinedError("ctx.contextV2"));
    return ctx.contextV2.cryptoProvider.encrypt(plaintext);
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(true),
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW(true),
    ContextInjectorMW,
  ])
  async decrypt(
    ciphertext: string,
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<string, FxError>> {
    if (!ctx) return err(new ObjectIsUndefinedError("ctx"));
    if (!ctx.contextV2) return err(new ObjectIsUndefinedError("ctx.contextV2"));
    return ctx.contextV2.cryptoProvider.decrypt(ciphertext);
  }

  async buildArtifacts(inputs: Inputs): Promise<Result<Void, FxError>> {
    throw new TaskNotSupportError(Stage.build);
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(false),
    ProjectSettingsLoaderMW,
    SolutionLoaderMW,
    EnvInfoLoaderMW(true),
    ContextInjectorMW,
  ])
  async createEnv(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    if (!ctx) return err(new ObjectIsUndefinedError("createEnv input stuff"));
    const projectSettings = ctx.projectSettings;
    if (!isMultiEnvEnabled() || !projectSettings) {
      return ok(Void);
    }

    const core = ctx!.self as FxCore;
    const createEnvCopyInput = await askNewEnvironment(ctx!, inputs);

    if (
      !createEnvCopyInput ||
      !createEnvCopyInput.targetEnvName ||
      !createEnvCopyInput.sourceEnvName
    ) {
      return err(UserCancelError);
    }

    const createEnvResult = await this.createEnvCopy(
      createEnvCopyInput.targetEnvName,
      createEnvCopyInput.sourceEnvName,
      inputs,
      core
    );

    if (createEnvResult.isErr()) {
      return createEnvResult;
    }

    inputs.sourceEnvName = createEnvCopyInput.sourceEnvName;
    inputs.targetEnvName = createEnvCopyInput.targetEnvName;

    if (!ctx.solutionV2 || !ctx.contextV2)
      return err(new ObjectIsUndefinedError("ctx.solutionV2, ctx.contextV2"));
    if (ctx.solutionV2.createEnv) {
      inputs.copy = true;
      return await ctx.solutionV2.createEnv(ctx.contextV2, inputs);
    }
    return ok(Void);
  }

  async createEnvWithName(
    targetEnvName: string,
    projectSettings: ProjectSettings,
    inputs: Inputs
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
    this.tools.logProvider.debug(
      `[core] persist ${targetEnvName} env state to path ${writeEnvResult.value}: ${JSON.stringify(
        newEnvConfig
      )}`
    );
    return ok(Void);
  }

  async createEnvCopy(
    targetEnvName: string,
    sourceEnvName: string,
    inputs: Inputs,
    core: FxCore
  ): Promise<Result<Void, FxError>> {
    // copy env config file
    const targetEnvConfigFilePath = environmentManager.getEnvConfigPath(
      targetEnvName,
      inputs.projectPath!
    );
    const sourceEnvConfigFilePath = environmentManager.getEnvConfigPath(
      sourceEnvName,
      inputs.projectPath!
    );

    try {
      await fs.copy(sourceEnvConfigFilePath, targetEnvConfigFilePath);
    } catch (e) {
      return err(CopyFileError(e as Error));
    }

    TOOLS.logProvider.debug(
      `[core] copy env config file for ${targetEnvName} environment to path ${targetEnvConfigFilePath}`
    );

    return ok(Void);
  }

  // deprecated
  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    SupportV1ConditionMW(true),
    ProjectMigratorMW,
    ProjectSettingsLoaderMW,
    SolutionLoaderMW,
    ContextInjectorMW,
    ProjectSettingsWriterMW,
  ])
  async activateEnv(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    const env = inputs.env;
    if (!env) {
      return err(new ObjectIsUndefinedError("env"));
    }
    if (!isMultiEnvEnabled() || !ctx!.projectSettings) {
      return ok(Void);
    }

    const envConfigs = await environmentManager.listEnvConfigs(inputs.projectPath!);

    if (envConfigs.isErr()) {
      return envConfigs;
    }

    if (envConfigs.isErr() || envConfigs.value.indexOf(env) < 0) {
      return err(NonExistEnvNameError(env));
    }

    const core = ctx!.self as FxCore;
    const solutionContext = await loadSolutionContext(inputs, ctx!.projectSettings, env);

    if (!solutionContext.isErr()) {
      ctx!.provisionInputConfig = solutionContext.value.envInfo.config;
      ctx!.provisionOutputs = solutionContext.value.envInfo.state;
      ctx!.envName = solutionContext.value.envInfo.envName;
    }

    this.tools.ui.showMessage("info", `[${env}] is activated.`, false);
    return ok(Void);
  }

  async _init(
    inputs: v2.InputsWithProjectPath & { solution?: string },
    ctx?: CoreHookContext
  ): Promise<Result<Void, FxError>> {
    if (!ctx) {
      return err(new ObjectIsUndefinedError("ctx for createProject"));
    }
    const appName = inputs[QuestionAppName.name] as string;
    const validateResult = jsonschema.validate(appName, {
      pattern: ProjectNamePattern,
    });
    if (validateResult.errors && validateResult.errors.length > 0) {
      return err(InvalidInputError("invalid app-name", inputs));
    }
    const projectSettings = newProjectSettings();
    projectSettings.appName = appName;
    ctx.projectSettings = projectSettings;
    if (!inputs.solution) {
      return err(InvalidInputError("solution is undefined", inputs));
    }
    const createEnvResult = await this.createEnvWithName(
      environmentManager.getDefaultEnvName(),
      projectSettings,
      inputs
    );
    if (createEnvResult.isErr()) {
      return err(createEnvResult.error);
    }
    await fs.ensureDir(path.join(inputs.projectPath, `.${ConfigFolderName}`));
    await fs.ensureDir(path.join(inputs.projectPath, "templates", `${AppPackageFolderName}`));
    const basicFolderRes = await createBasicFolderStructure(inputs);
    if (basicFolderRes.isErr()) {
      return err(basicFolderRes.error);
    }
    const solution = Container.get<v3.ISolution>(inputs.solution);
    projectSettings.solutionSettings.name = inputs.solution;
    const context = createV2Context(projectSettings);
    ctx.contextV2 = context;
    ctx.solutionV3 = solution;
    return await solution.init(
      context,
      inputs as v2.InputsWithProjectPath & { capabilities: string[] }
    );
  }
  @hooks([ErrorHandlerMW, QuestionModelMW, ContextInjectorMW, ProjectSettingsWriterMW])
  async init(
    inputs: v2.InputsWithProjectPath & { solution?: string },
    ctx?: CoreHookContext
  ): Promise<Result<Void, FxError>> {
    return this._init(inputs, ctx);
  }
  async _addModule(
    inputs: v2.InputsWithProjectPath,
    ctx?: CoreHookContext
  ): Promise<Result<Void, FxError>> {
    if (ctx && ctx.solutionV3 && ctx.contextV2) {
      return await ctx.solutionV3.addModule(
        ctx.contextV2,
        {},
        inputs as v2.InputsWithProjectPath & { capabilities?: string[] }
      );
    }
    return ok(Void);
  }
  @hooks([
    ErrorHandlerMW,
    ProjectSettingsLoaderMW_V3,
    SolutionLoaderMW_V3,
    QuestionModelMW,
    ContextInjectorMW,
    ProjectSettingsWriterMW,
  ])
  async addModule(
    inputs: v2.InputsWithProjectPath,
    ctx?: CoreHookContext
  ): Promise<Result<Void, FxError>> {
    return this._addModule(inputs, ctx);
  }

  @hooks([
    ErrorHandlerMW,
    ProjectSettingsLoaderMW_V3,
    SolutionLoaderMW_V3,
    QuestionModelMW,
    ContextInjectorMW,
    ProjectSettingsWriterMW,
  ])
  async scaffold(
    inputs: v2.InputsWithProjectPath,
    ctx?: CoreHookContext
  ): Promise<Result<Void, FxError>> {
    return this._scaffold(inputs, ctx);
  }
  async _scaffold(
    inputs: v2.InputsWithProjectPath,
    ctx?: CoreHookContext
  ): Promise<Result<Void, FxError>> {
    if (ctx && ctx.solutionV3 && ctx.contextV2) {
      return await ctx.solutionV3.scaffold(ctx.contextV2, inputs);
    }
    return ok(Void);
  }
  @hooks([
    ErrorHandlerMW,
    ProjectSettingsLoaderMW_V3,
    SolutionLoaderMW_V3,
    QuestionModelMW,
    ContextInjectorMW,
    ProjectSettingsWriterMW,
  ])
  async addResource(
    inputs: v2.InputsWithProjectPath,
    ctx?: CoreHookContext
  ): Promise<Result<Void, FxError>> {
    return this._addResource(inputs, ctx);
  }
  async _addResource(
    inputs: v2.InputsWithProjectPath,
    ctx?: CoreHookContext
  ): Promise<Result<Void, FxError>> {
    if (ctx && ctx.solutionV3 && ctx.contextV2) {
      return await ctx.solutionV3.addResource(ctx.contextV2, inputs);
    }
    return ok(Void);
  }

  //V1,V2 questions
  _getQuestionsForCreateProjectV2 = getQuestionsForCreateProjectV2;
  _getQuestionsForCreateProjectV3 = getQuestionsForCreateProjectV3;
  _getQuestionsForUserTask = getQuestionsForUserTaskV2;
  _getQuestions = getQuestionsV2;
  _getQuestionsForMigrateV1Project = getQuestionsForMigrateV1Project;
  //v3 questions
  _getQuestionsForScaffold = getQuestionsForScaffold;
  _getQuestionsForAddModule = getQuestionsForAddModule;
  _getQuestionsForAddResource = getQuestionsForAddResource;
  _getQuestionsForProvision = getQuestionsForProvision;
  _getQuestionsForDeploy = getQuestionsForDeploy;
  _getQuestionsForLocalProvision = getQuestionsForLocalProvision;
  _getQuestionsForPublish = getQuestionsForPublish;
  _getQuestionsForInit = getQuestionsForInit;
}

export async function createBasicFolderStructure(inputs: Inputs): Promise<Result<null, FxError>> {
  if (!inputs.projectPath) {
    return err(new ObjectIsUndefinedError("projectPath"));
  }
  try {
    const appName = inputs[QuestionAppName.name] as string;
    if (inputs.platform !== Platform.VS) {
      await fs.writeFile(
        path.join(inputs.projectPath, `package.json`),
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
    }
    await fs.writeFile(
      path.join(inputs.projectPath!, `.gitignore`),
      isMultiEnvEnabled()
        ? [
            "node_modules",
            `.${ConfigFolderName}/${InputConfigsFolderName}/${localSettingsFileName}`,
            `.${ConfigFolderName}/${StatesFolderName}/*.userdata`,
            ".DS_Store",
            `${ArchiveFolderName}`,
            `${ArchiveLogFileName}`,
            ".env.teamsfx.local",
            "subscriptionInfo.json",
            BuildFolderName,
          ].join("\n")
        : `node_modules\n/.${ConfigFolderName}/*.env\n/.${ConfigFolderName}/*.userdata\n.DS_Store\n${ArchiveFolderName}\n${ArchiveLogFileName}`
    );
  } catch (e) {
    return err(WriteFileError(e));
  }
  return ok(null);
}
export async function downloadSample(
  inputs: Inputs,
  ctx: CoreHookContext
): Promise<Result<string, FxError>> {
  let fxError;
  const progress = TOOLS.ui.createProgressBar("Fetch sample app", 3);
  progress.start();
  const telemetryProperties: any = {
    [TelemetryProperty.Success]: TelemetrySuccess.Yes,
    module: "fx-core",
  };
  try {
    let folder = inputs[QuestionRootFolder.name] as string;
    if (inputs.platform === Platform.VSCode) {
      folder = getRootDirectory();
      await fs.ensureDir(folder);
    }
    const sampleId = inputs[CoreQuestionNames.Samples] as string;
    if (!(sampleId && folder)) {
      throw InvalidInputError(`invalid answer for '${CoreQuestionNames.Samples}'`, inputs);
    }
    telemetryProperties[TelemetryProperty.SampleAppName] = sampleId;
    const samples = sampleProvider.SampleCollection.samples.filter(
      (sample) => sample.id.toLowerCase() === sampleId.toLowerCase()
    );
    if (samples.length === 0) {
      throw InvalidInputError(`invalid sample id: '${sampleId}'`, inputs);
    }
    const sample = samples[0];
    const url = sample.link as string;
    let sampleAppPath = path.resolve(folder, sampleId);
    if ((await fs.pathExists(sampleAppPath)) && (await fs.readdir(sampleAppPath)).length > 0) {
      let suffix = 1;
      while (await fs.pathExists(sampleAppPath)) {
        sampleAppPath = `${folder}/${sampleId}_${suffix++}`;
      }
    }
    progress.next(`Downloading from ${url}`);
    const fetchRes = await fetchCodeZip(url, sample.id);
    if (fetchRes.isErr()) {
      throw fetchRes.error;
    } else if (!fetchRes.value) {
      throw FetchSampleError(sample.id);
    }
    progress.next("Unzipping the sample package");
    await saveFilesRecursively(new AdmZip(fetchRes.value.data), sampleId, sampleAppPath);
    await downloadSampleHook(sampleId, sampleAppPath);
    progress.next("Update project settings");
    const loadInputs: Inputs = {
      ...inputs,
      projectPath: sampleAppPath,
    };
    const projectSettingsRes = await loadProjectSettings(loadInputs, true);
    if (projectSettingsRes.isOk()) {
      const projectSettings = projectSettingsRes.value;
      projectSettings.projectId = inputs.projectId ? inputs.projectId : uuid.v4();
      projectSettings.isFromSample = true;
      inputs.projectId = projectSettings.projectId;
      telemetryProperties[TelemetryProperty.ProjectId] = projectSettings.projectId;
      ctx.projectSettings = projectSettings;
      inputs.projectPath = sampleAppPath;
    } else {
      telemetryProperties[TelemetryProperty.ProjectId] =
        "unknown, failed to set projectId in projectSettings.json";
    }
    progress.end(true);
    sendTelemetryEvent(Component.core, TelemetryEvent.DownloadSample, telemetryProperties);
    return ok(sampleAppPath);
  } catch (e) {
    fxError = assembleError(e);
    progress.end(false);
    telemetryProperties[TelemetryProperty.Success] = TelemetrySuccess.No;
    sendTelemetryErrorEvent(
      Component.core,
      TelemetryEvent.DownloadSample,
      fxError,
      telemetryProperties
    );
    return err(fxError);
  }
}

export function newProjectSettings(): ProjectSettings {
  const projectSettings: ProjectSettings = {
    appName: "",
    projectId: uuid.v4(),
    version: getProjectSettingsVersion(),
    solutionSettings: {
      name: "",
    },
  };
  return projectSettings;
}

export function createV2Context(projectSettings: ProjectSettings): v2.Context {
  const context: v2.Context = {
    userInteraction: TOOLS.ui,
    logProvider: TOOLS.logProvider,
    telemetryReporter: TOOLS.telemetryReporter!,
    cryptoProvider: new LocalCrypto(projectSettings.projectId),
    permissionRequestProvider: TOOLS.permissionRequest,
    projectSetting: projectSettings,
  };
  return context;
}

export function undefinedName(objs: any[], names: string[]) {
  for (let i = 0; i < objs.length; ++i) {
    if (objs[i] === undefined) {
      return names[i];
    }
  }
  return undefined;
}

export function getProjectSettingsVersion() {
  if (isMultiEnvEnabled()) return "2.0.0";
  else return "1.0.0";
}

export * from "./error";
export * from "./tools";
