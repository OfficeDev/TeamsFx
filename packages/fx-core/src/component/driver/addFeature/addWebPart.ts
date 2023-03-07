// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Result,
  FxError,
  ok,
  err,
  ManifestUtil,
  Platform,
  IStaticTab,
  v3,
  Inputs,
  Stage,
} from "@microsoft/teamsfx-api";
import { hooks } from "@feathersjs/hooks/lib";
import { Service } from "typedi";
import { StepDriver, ExecutionResult } from "../interface/stepDriver";
import { DriverContext } from "../interface/commonArgs";
import { WrapDriverContext } from "../util/wrapUtil";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { manifestUtils } from "../../resource/appManifest/utils/ManifestUtils";
import { AppStudioResultFactory } from "../../resource/appManifest/results";
import { AppStudioError } from "../../resource/appManifest/errors";
import { getLocalizedString } from "../../../common/localizeUtils";
import { HelpLinks } from "../../../common/constants";
import { getAbsolutePath, wrapRun } from "../../utils/common";
import { AddWebPartArgs } from "./interface/AddWebPartArgs";
import { Utils } from "../../resource/spfx/utils/utils";
import { camelCase } from "lodash";
import path from "path";
import { getTemplatesFolder } from "../../../folder";
import { YoChecker } from "../../resource/spfx/depsChecker/yoChecker";
import { GeneratorChecker } from "../../resource/spfx/depsChecker/generatorChecker";
import { isGeneratorCheckerEnabled, isYoCheckerEnabled } from "../../../common/tools";
import { DependencyInstallError } from "../../resource/spfx/error";
import { cpUtils } from "../../../common/deps-checker";
import { DefaultManifestProvider } from "../../resource/appManifest/manifestProvider";
import * as fs from "fs-extra";
import * as util from "util";
import { ManifestTemplate } from "../../resource/spfx/utils/constants";
import { SPFxGenerator } from "../../generator/spfxGenerator";
import { createContextV3 } from "../../utils";
import { SPFXQuestionNames } from "../../resource/spfx/utils/questions";
import { Constants } from "./utility/constants";
import { NoConfigurationError } from "./error/noConfigurationError";

@Service(Constants.ActionName)
export class AddWebPartDriver implements StepDriver {
  description = getLocalizedString("driver.spfx.add.description");

  @hooks([addStartAndEndTelemetry(Constants.ActionName, Constants.ActionName)])
  public async run(
    args: AddWebPartArgs,
    context: DriverContext
  ): Promise<Result<Map<string, string>, FxError>> {
    const wrapContext = new WrapDriverContext(context, Constants.ActionName, Constants.ActionName);
    return wrapRun(() => this.add(args, wrapContext));
  }

  public async execute(args: AddWebPartArgs, context: DriverContext): Promise<ExecutionResult> {
    const wrapContext = new WrapDriverContext(context, Constants.ActionName, Constants.ActionName);
    const res = await this.run(args, wrapContext);
    return {
      result: res,
      summaries: wrapContext.summaries,
    };
  }

  public async add(args: AddWebPartArgs, context: WrapDriverContext): Promise<Map<string, string>> {
    const webpartName = args.webpartName;
    const spfxFolder = args.spfxFolder;
    const manifestPath = args.manifestPath;
    const localManifestPath = args.localManifestPath;

    const yorcPath = path.join(spfxFolder, Constants.YO_RC_FILE);
    if (!(await fs.pathExists(yorcPath))) {
      throw new NoConfigurationError();
    }

    const inputs: Inputs = { platform: context.platform, stage: Stage.addWebpart };
    inputs[SPFXQuestionNames.webpart_name] = webpartName;
    inputs["spfxFolder"] = spfxFolder;
    inputs["manifestPath"] = manifestPath;
    inputs["localManifestPath"] = localManifestPath;
    const yeomanRes = await SPFxGenerator.doYeomanScaffold(
      createContextV3(),
      inputs,
      context.projectPath
    );
    if (yeomanRes.isErr()) throw yeomanRes.error;

    const componentId = yeomanRes.value;
    const remoteStaticSnippet: IStaticTab = {
      entityId: componentId,
      name: webpartName,
      contentUrl: util.format(Constants.REMOTE_CONTENT_URL, componentId, componentId),
      websiteUrl: ManifestTemplate.WEBSITE_URL,
      scopes: ["personal"],
    };
    const localStaticSnippet: IStaticTab = {
      entityId: componentId,
      name: webpartName,
      contentUrl: util.format(Constants.LOCAL_CONTENT_URL, componentId, componentId),
      websiteUrl: ManifestTemplate.WEBSITE_URL,
      scopes: ["personal"],
    };

    inputs["addManifestPath"] = localManifestPath;
    const localRes = await manifestUtils.addCapabilities(
      { ...inputs, projectPath: context.projectPath },
      [{ name: "staticTab", snippet: localStaticSnippet }]
    );
    if (localRes.isErr()) throw localRes.error;

    inputs["addManifestPath"] = manifestPath;
    const remoteRes = await manifestUtils.addCapabilities(
      { ...inputs, projectPath: context.projectPath },
      [{ name: "staticTab", snippet: remoteStaticSnippet }]
    );
    if (remoteRes.isErr()) throw remoteRes.error;

    return new Map();
  }
}
