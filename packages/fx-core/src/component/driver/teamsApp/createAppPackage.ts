// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import { Colors, FxError, PluginManifest, Result, err, ok } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import fs from "fs-extra";
import * as path from "path";
import { Service } from "typedi";
import { getLocalizedString } from "../../../common/localizeUtils";
import { ErrorContextMW } from "../../../core/globalVars";
import {
  FileNotFoundError,
  InvalidActionInputError,
  JSONSyntaxError,
  MissingEnvironmentVariablesError,
} from "../../../error/common";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { WrapDriverContext } from "../util/wrapUtil";
import { Constants } from "./constants";
import { CreateAppPackageArgs } from "./interfaces/CreateAppPackageArgs";
import { manifestUtils } from "./utils/ManifestUtils";
import { expandEnvironmentVariable, getEnvironmentVariables } from "../../utils/common";
import { TelemetryPropertyKey } from "./utils/telemetry";
import { InvalidFileOutsideOfTheDirectotryError } from "../../../error/teamsApp";

export const actionName = "teamsApp/zipAppPackage";

@Service(actionName)
export class CreateAppPackageDriver implements StepDriver {
  description = getLocalizedString("driver.teamsApp.description.createAppPackageDriver");
  readonly progressTitle = getLocalizedString(
    "plugins.appstudio.createPackage.progressBar.message"
  );

  public async execute(
    args: CreateAppPackageArgs,
    context: DriverContext
  ): Promise<ExecutionResult> {
    const wrapContext = new WrapDriverContext(context, actionName, actionName);
    const res = await this.build(args, wrapContext);
    return {
      result: res,
      summaries: wrapContext.summaries,
    };
  }

  @hooks([
    ErrorContextMW({ source: "Teams", component: "CreateAppPackageDriver" }),
    addStartAndEndTelemetry(actionName, actionName),
  ])
  public async build(
    args: CreateAppPackageArgs,
    context: WrapDriverContext
  ): Promise<Result<Map<string, string>, FxError>> {
    const result = this.validateArgs(args);
    if (result.isErr()) {
      return err(result.error);
    }

    let manifestPath = args.manifestPath;
    if (!path.isAbsolute(manifestPath)) {
      manifestPath = path.join(context.projectPath, manifestPath);
    }

    const manifestRes = await manifestUtils.getManifestV3(manifestPath, context);
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }
    const manifest = manifestRes.value;
    // Deal with relative path
    // Environment variables should have been replaced by value
    // ./build/appPackage/appPackage.dev.zip instead of ./build/appPackage/appPackage.${{TEAMSFX_ENV}}.zip
    let zipFileName = args.outputZipPath;
    if (!path.isAbsolute(zipFileName)) {
      zipFileName = path.join(context.projectPath, zipFileName);
    }
    const zipFileDir = path.dirname(zipFileName);
    await fs.mkdir(zipFileDir, { recursive: true });

    let jsonFileName = args.outputJsonPath;
    if (!path.isAbsolute(jsonFileName)) {
      jsonFileName = path.join(context.projectPath, jsonFileName);
    }
    const jsonFileDir = path.dirname(jsonFileName);
    await fs.mkdir(jsonFileDir, { recursive: true });

    const appDirectory = path.dirname(manifestPath);

    const colorFile = path.resolve(appDirectory, manifest.icons.color);
    if (!(await fs.pathExists(colorFile))) {
      const error = new FileNotFoundError(
        actionName,
        colorFile,
        "https://aka.ms/teamsfx-actions/teamsapp-zipAppPackage"
      );
      return err(error);
    }
    const colorFileRelativePath = path.relative(appDirectory, colorFile);
    if (colorFileRelativePath.startsWith("..")) {
      return err(new InvalidFileOutsideOfTheDirectotryError(colorFile));
    }

    const outlineFile = path.resolve(appDirectory, manifest.icons.outline);
    if (!(await fs.pathExists(outlineFile))) {
      const error = new FileNotFoundError(
        actionName,
        outlineFile,
        "https://aka.ms/teamsfx-actions/teamsapp-zipAppPackage"
      );
      return err(error);
    }
    const outlineFileRelativePath = path.relative(appDirectory, outlineFile);
    if (outlineFileRelativePath.startsWith("..")) {
      return err(new InvalidFileOutsideOfTheDirectotryError(outlineFile));
    }

    // pre-check existence
    if (
      manifest.localizationInfo &&
      manifest.localizationInfo.additionalLanguages &&
      manifest.localizationInfo.additionalLanguages.length > 0
    ) {
      for (const language of manifest.localizationInfo.additionalLanguages) {
        const file = language.file;
        const fileName = `${appDirectory}/${file}`;
        if (!(await fs.pathExists(fileName))) {
          return err(
            new FileNotFoundError(
              actionName,
              fileName,
              "https://aka.ms/teamsfx-actions/teamsapp-zipAppPackage"
            )
          );
        }
      }
    }

    const zip = new AdmZip();
    zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest, null, 4)));

    // outline.png & color.png, relative path
    let dir = path.dirname(manifest.icons.color);
    zip.addLocalFile(colorFile, dir === "." ? "" : dir);
    dir = path.dirname(manifest.icons.outline);
    zip.addLocalFile(outlineFile, dir === "." ? "" : dir);

    // localization file
    if (
      manifest.localizationInfo &&
      manifest.localizationInfo.additionalLanguages &&
      manifest.localizationInfo.additionalLanguages.length > 0
    ) {
      for (const language of manifest.localizationInfo.additionalLanguages) {
        const file = language.file;
        const fileName = path.resolve(appDirectory, file);
        const relativePath = path.relative(appDirectory, fileName);
        if (relativePath.startsWith("..")) {
          return err(new InvalidFileOutsideOfTheDirectotryError(fileName));
        }
        const dir = path.dirname(file);
        zip.addLocalFile(fileName, dir === "." ? "" : dir);
      }
    }

    // API ME, API specification and Adaptive card templates
    if (
      manifest.composeExtensions &&
      manifest.composeExtensions.length > 0 &&
      manifest.composeExtensions[0].composeExtensionType == "apiBased" &&
      manifest.composeExtensions[0].apiSpecificationFile
    ) {
      const apiSpecificationFile = path.resolve(
        appDirectory,
        manifest.composeExtensions[0].apiSpecificationFile
      );
      const checkExistenceRes = await this.validateReferencedFile(
        apiSpecificationFile,
        appDirectory
      );
      if (checkExistenceRes.isErr()) {
        return err(checkExistenceRes.error);
      }
      const expandedEnvVarResult = await CreateAppPackageDriver.expandOpenAPIEnvVars(
        apiSpecificationFile,
        context
      );
      if (expandedEnvVarResult.isErr()) {
        return err(expandedEnvVarResult.error);
      }
      const openAPIContent = expandedEnvVarResult.value;
      const attr = await fs.stat(apiSpecificationFile);
      zip.addFile(
        manifest.composeExtensions[0].apiSpecificationFile,
        Buffer.from(openAPIContent),
        "",
        attr.mode
      );

      if (manifest.composeExtensions[0].commands.length > 0) {
        for (const command of manifest.composeExtensions[0].commands) {
          if (command.apiResponseRenderingTemplateFile) {
            const adaptiveCardFile = path.resolve(
              appDirectory,
              command.apiResponseRenderingTemplateFile
            );
            const checkExistenceRes = await this.validateReferencedFile(
              adaptiveCardFile,
              appDirectory
            );
            if (checkExistenceRes.isErr()) {
              return err(checkExistenceRes.error);
            }
            const dir = path.dirname(command.apiResponseRenderingTemplateFile);
            this.addFileInZip(zip, dir, adaptiveCardFile);
          }
        }
      }
    }

    // API plugin
    if (
      manifest.apiPlugins &&
      manifest.apiPlugins.length > 0 &&
      manifest.apiPlugins[0].pluginFile
    ) {
      const pluginFile = path.resolve(appDirectory, manifest.apiPlugins[0].pluginFile);
      const checkExistenceRes = await this.validateReferencedFile(pluginFile, appDirectory);
      if (checkExistenceRes.isErr()) {
        return err(checkExistenceRes.error);
      }
      const dir = path.dirname(manifest.apiPlugins[0].pluginFile);
      this.addFileInZip(zip, dir, pluginFile);

      // Add API spec and templates
      const addResponseTemplateRes = await this.addPluginRelatedFiles(
        zip,
        pluginFile,
        appDirectory
      );
      if (addResponseTemplateRes.isErr()) {
        return err(addResponseTemplateRes.error);
      }
    }

    zip.writeZip(zipFileName);

    if (await fs.pathExists(jsonFileName)) {
      await fs.chmod(jsonFileName, 0o777);
    }
    await fs.writeFile(jsonFileName, JSON.stringify(manifest, null, 4));
    await fs.chmod(jsonFileName, 0o444);

    const builtSuccess = [
      { content: "(√)Done: ", color: Colors.BRIGHT_GREEN },
      { content: "Teams Package ", color: Colors.BRIGHT_WHITE },
      { content: zipFileName, color: Colors.BRIGHT_MAGENTA },
      { content: " built successfully!", color: Colors.BRIGHT_WHITE },
    ];
    context.logProvider.info(builtSuccess);
    return ok(new Map());
  }

  private static async expandOpenAPIEnvVars(
    openAPISpecPath: string,
    ctx: WrapDriverContext
  ): Promise<Result<string, FxError>> {
    const content = await fs.readFile(openAPISpecPath, "utf8");
    const vars = getEnvironmentVariables(content);
    ctx.addTelemetryProperties({
      [TelemetryPropertyKey.customizedOpenAPIKeys]: vars.join(";"),
    });
    const result = expandEnvironmentVariable(content);
    const notExpandedVars = getEnvironmentVariables(result);
    if (notExpandedVars.length > 0) {
      return err(
        new MissingEnvironmentVariablesError("teamsApp", notExpandedVars.join(","), openAPISpecPath)
      );
    }
    return ok(result);
  }

  private validateArgs(args: CreateAppPackageArgs): Result<any, FxError> {
    const invalidParams: string[] = [];
    if (!args || !args.manifestPath) {
      invalidParams.push("manifestPath");
    }
    if (!args || !args.outputJsonPath) {
      invalidParams.push("outputJsonPath");
    }
    if (!args || !args.outputZipPath) {
      invalidParams.push("outputZipPath");
    }
    if (invalidParams.length > 0) {
      return err(
        new InvalidActionInputError(
          actionName,
          invalidParams,
          "https://aka.ms/teamsfx-actions/teamsapp-zipAppPackage"
        )
      );
    } else {
      return ok(undefined);
    }
  }

  private async validateReferencedFile(
    file: string,
    directory: string
  ): Promise<Result<undefined, FxError>> {
    if (!(await fs.pathExists(file))) {
      return err(
        new FileNotFoundError(
          actionName,
          file,
          "https://aka.ms/teamsfx-actions/teamsapp-zipAppPackage"
        )
      );
    }

    const relativePath = path.relative(directory, file);
    if (relativePath.startsWith("..")) {
      return err(new InvalidFileOutsideOfTheDirectotryError(file));
    }

    return ok(undefined);
  }

  private async addPluginRelatedFiles(
    zip: AdmZip,
    pluginFile: string,
    appDirectory: string
  ): Promise<Result<undefined, FxError>> {
    let pluginContent;
    try {
      pluginContent = (await fs.readJSON(pluginFile)) as PluginManifest;
    } catch (e) {
      return err(new JSONSyntaxError(pluginFile, e, actionName));
    }
    const runtimes = pluginContent.runtimes;
    if (runtimes && runtimes.length > 0) {
      for (const runtime of runtimes) {
        if (runtime.type === "openapi" && runtime.spec?.url) {
          const specFile = path.resolve(path.dirname(pluginFile), runtime.spec.url);
          // add openapi spec
          const checkExistenceRes = await this.validateReferencedFile(specFile, appDirectory);
          if (checkExistenceRes.isErr()) {
            return err(checkExistenceRes.error);
          }

          const dir = path.relative(appDirectory, path.dirname(specFile));
          this.addFileInZip(zip, dir, specFile);
        }
      }
    }

    const functions = pluginContent.functions;
    if (functions && functions.length > 0) {
      for (const func of functions) {
        const templates = func.capabilities?.rendering_templates;
        if (templates) {
          for (const key of Object.keys(templates)) {
            const template = templates[key];
            const templateFile = path.resolve(path.dirname(pluginFile), template.template_url);
            const checkExistenceRes = await this.validateReferencedFile(templateFile, appDirectory);
            if (checkExistenceRes.isErr()) {
              return err(checkExistenceRes.error);
            }
            const dir = path.relative(appDirectory, path.dirname(templateFile));
            this.addFileInZip(zip, dir, templateFile);
          }
        }
      }
    }

    return ok(undefined);
  }

  private addFileInZip(zip: AdmZip, dir: string, filePath: string) {
    zip.addLocalFile(filePath, dir === "." ? "" : dir);
  }
}
