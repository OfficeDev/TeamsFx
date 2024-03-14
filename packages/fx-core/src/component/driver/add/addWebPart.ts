// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, FxError, IStaticTab, Inputs, Stage } from "@microsoft/teamsfx-api";
import { hooks } from "@feathersjs/hooks/lib";
import { Service } from "typedi";
import { StepDriver, ExecutionResult } from "../interface/stepDriver";
import { DriverContext } from "../interface/commonArgs";
import { WrapDriverContext } from "../util/wrapUtil";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { manifestUtils } from "../teamsApp/utils/ManifestUtils";
import { getLocalizedString } from "../../../common/localizeUtils";
import { wrapRun } from "../../utils/common";
import { AddWebPartArgs } from "./interface/AddWebPartArgs";
import path from "path";
import * as fs from "fs-extra";
import * as util from "util";
import { ManifestTemplate } from "../../generator/spfx/utils/constants";
import { SPFxGenerator } from "../../generator/spfx/spfxGenerator";
import { createContextV3 } from "../../utils";
import { Constants } from "./utility/constants";
import { NoConfigurationError } from "./error/noConfigurationError";
import { QuestionNames } from "../../../question/questionNames";

@Service(Constants.ActionName)
export class AddWebPartDriver implements StepDriver {
  description = getLocalizedString("driver.spfx.add.description");

  @hooks([addStartAndEndTelemetry(Constants.ActionName, Constants.ActionName)])
  public async run(
    args: AddWebPartArgs,
    context: DriverContext
  ): Promise<Result<Map<string, string>, FxError>> {
    const wrapContext = new WrapDriverContext(context, Constants.ActionName, Constants.ActionName);
    return wrapRun(() => this.add(args, wrapContext), Constants.ActionName);
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
    const framework = args.framework;
    const spfxFolder = args.spfxFolder;
    const manifestPath = args.manifestPath;
    const localManifestPath = args.localManifestPath;
    const spfxPackage = args.spfxPackage;

    const yorcPath = path.join(spfxFolder, Constants.YO_RC_FILE);
    context.logProvider.verbose(`Checking configuration file under ${yorcPath}...`);
    if (!(await fs.pathExists(yorcPath))) {
      throw new NoConfigurationError();
    }
    context.logProvider.verbose(`Configuration file exists.`);

    const inputs: Inputs = { platform: context.platform, stage: Stage.addWebpart };
    inputs[QuestionNames.SPFxWebpartName] = webpartName;
    inputs[QuestionNames.SPFxFramework] = framework;
    inputs[QuestionNames.SPFxFolder] = spfxFolder;
    inputs[QuestionNames.ManifestPath] = manifestPath;
    inputs[QuestionNames.LocalTeamsAppManifestFilePath] = localManifestPath;
    inputs[QuestionNames.SPFxInstallPackage] = spfxPackage;
    context.logProvider.verbose(`Adding web part with Yeoman generator...`);
    context.logProvider.debug(
      `SPFx web part name: ${webpartName}, SPFx folder: ${spfxFolder}, manifest path: ${manifestPath}, local manifest path: ${localManifestPath}`
    );
    const yeomanRes = await SPFxGenerator.doYeomanScaffold(
      createContextV3(),
      inputs,
      context.projectPath
    );
    if (yeomanRes.isErr()) throw yeomanRes.error;
    context.logProvider.verbose(`Succeeded to add web part '${webpartName}'.`);

    const componentId = yeomanRes.value;
    const remoteStaticSnippet: IStaticTab = {
      entityId: componentId,
      name: webpartName,
      contentUrl: util.format(Constants.REMOTE_CONTENT_URL, componentId),
      websiteUrl: ManifestTemplate.WEBSITE_URL,
      scopes: ["personal", "team", "groupChat"],
    };
    const localStaticSnippet: IStaticTab = {
      entityId: componentId,
      name: webpartName,
      contentUrl: util.format(Constants.LOCAL_CONTENT_URL, componentId),
      websiteUrl: ManifestTemplate.WEBSITE_URL,
      scopes: ["personal", "team", "groupChat"],
    };

    inputs["addManifestPath"] = localManifestPath;
    context.logProvider.verbose(
      `Exposing web part with component id ${componentId} to local manifest file under ${localManifestPath}...`
    );
    const localRes = await manifestUtils.addCapabilities(
      { ...inputs, projectPath: context.projectPath },
      [{ name: "staticTab", snippet: localStaticSnippet }]
    );
    if (localRes.isErr()) throw localRes.error;
    context.logProvider.verbose(`Succeeded to update local manifest file.`);

    inputs["addManifestPath"] = manifestPath;
    context.logProvider.verbose(
      `Exposing web part with component id ${componentId} to remote manifest file under ${manifestPath}...`
    );
    const remoteRes = await manifestUtils.addCapabilities(
      { ...inputs, projectPath: context.projectPath },
      [{ name: "staticTab", snippet: remoteStaticSnippet }]
    );
    if (remoteRes.isErr()) throw remoteRes.error;
    context.logProvider.verbose(`Succeeded to update remote manifest file.`);

    context.ui?.showMessage(
      "info",
      getLocalizedString("driver.spfx.add.successNotice", webpartName),
      false
    );
    return new Map();
  }
}
