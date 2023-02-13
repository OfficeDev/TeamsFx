// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, FxError, ok, err, ManifestUtil, Platform } from "@microsoft/teamsfx-api";
import { hooks } from "@feathersjs/hooks/lib";
import { Service } from "typedi";
import fs from "fs-extra";
import * as path from "path";
import { EOL } from "os";
import { StepDriver, ExecutionResult } from "../interface/stepDriver";
import { DriverContext } from "../interface/commonArgs";
import { WrapDriverContext } from "../util/wrapUtil";
import { ValidateTeamsAppArgs } from "./interfaces/ValidateTeamsAppArgs";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { manifestUtils } from "../../resource/appManifest/utils/ManifestUtils";
import { AppStudioResultFactory } from "../../resource/appManifest/results";
import { AppStudioError } from "../../resource/appManifest/errors";
import { AppStudioClient } from "../../resource/appManifest/appStudioClient";
import { getDefaultString, getLocalizedString } from "../../../common/localizeUtils";
import { HelpLinks } from "../../../common/constants";
import { AppStudioScopes } from "../../../common/tools";
import { getAbsolutePath } from "../../utils/common";

const actionName = "teamsApp/validate";

@Service(actionName)
export class ValidateTeamsAppDriver implements StepDriver {
  description = getLocalizedString("driver.teamsApp.description.validateDriver");

  public async run(
    args: ValidateTeamsAppArgs,
    context: DriverContext
  ): Promise<Result<Map<string, string>, FxError>> {
    const wrapContext = new WrapDriverContext(context, actionName, actionName);
    const res = await this.validate(args, wrapContext);
    return res;
  }

  public async execute(
    args: ValidateTeamsAppArgs,
    context: DriverContext
  ): Promise<ExecutionResult> {
    const wrapContext = new WrapDriverContext(context, actionName, actionName);
    const res = await this.validate(args, wrapContext);
    return {
      result: res,
      summaries: wrapContext.summaries,
    };
  }

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  public async validate(
    args: ValidateTeamsAppArgs,
    context: WrapDriverContext,
    withEmptyCapabilities?: boolean
  ): Promise<Result<Map<string, string>, FxError>> {
    /*const result = this.validateArgs(args);
    if (result.isErr()) {
      return err(result.error);
    }

    if (args.appPackagePath) {
      let appPackagePath = args.appPackagePath;
      if (!path.isAbsolute(appPackagePath)) {
        appPackagePath = path.join(context.projectPath, appPackagePath);
      }
      if (!(await fs.pathExists(appPackagePath))) {
        return err(
          AppStudioResultFactory.UserError(
            AppStudioError.FileNotFoundError.name,
            AppStudioError.FileNotFoundError.message(appPackagePath)
          )
        );
      }
      const archivedFile = await fs.readFile(appPackagePath);

      const appStudioTokenRes = await context.m365TokenProvider.getAccessToken({
        scopes: AppStudioScopes,
      });
      if (appStudioTokenRes.isErr()) {
        return err(appStudioTokenRes.error);
      }
      const appStudioToken = appStudioTokenRes.value;

      const validationResult = await AppStudioClient.partnerCenterAppPackageValidation(
        archivedFile,
        appStudioToken
      );

      // logs in output window
      const errors = validationResult.errors
        .map((error) => {
          return `(x) Error: ${error.content} \n${getLocalizedString("core.option.learnMore")}: ${
            error.helpUrl
          }`;
        })
        .join(EOL);
      const warnings = validationResult.warnings
        .map((warning) => {
          return `(!) Warning: ${warning.content} \n${getLocalizedString(
            "core.option.learnMore"
          )}: ${warning.helpUrl}`;
        })
        .join(EOL);
      const outputMessage =
        EOL +
        getLocalizedString(
          "driver.teamsApp.summary.validate",
          validationResult.errors.length + validationResult.warnings.length,
          validationResult.notes.length,
          errors,
          warnings,
          undefined
        );
      context.logProvider?.info(outputMessage);

      const message = getLocalizedString(
        "driver.teamsApp.validate.result",
        validationResult.errors.length + validationResult.warnings.length,
        validationResult.notes.length,
        "command:fx-extension.showOutputChannel"
      );
      context.ui?.showMessage("info", message, false);
      return ok(new Map());
    }

    const state = this.loadCurrentState();
    const manifestRes = await manifestUtils.getManifestV3(
      getAbsolutePath(args.manifestPath!, context.projectPath),
      state,
      withEmptyCapabilities
    );
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }
    const manifest = manifestRes.value;

    let validationResult;
    if (manifest.$schema) {
      try {
        validationResult = await ManifestUtil.validateManifest(manifest);
      } catch (e: any) {
        return err(
          AppStudioResultFactory.UserError(
            AppStudioError.ValidationFailedError.name,
            AppStudioError.ValidationFailedError.message([
              getLocalizedString(
                "error.appstudio.validateFetchSchemaFailed",
                manifest.$schema,
                e.message
              ),
            ]),
            HelpLinks.WhyNeedProvision
          )
        );
      }
    } else {
      return err(
        AppStudioResultFactory.UserError(
          AppStudioError.ValidationFailedError.name,
          AppStudioError.ValidationFailedError.message([
            getLocalizedString("error.appstudio.validateSchemaNotDefined"),
          ]),
          HelpLinks.WhyNeedProvision
        )
      );
    }

    if (validationResult.length > 0) {
      const errMessage = AppStudioError.ValidationFailedError.message(validationResult);
      context.logProvider?.error(getLocalizedString("plugins.appstudio.validationFailedNotice"));
      const validationFailed = AppStudioResultFactory.UserError(
        AppStudioError.ValidationFailedError.name,
        errMessage,
        "https://aka.ms/teamsfx-actions/teamsapp-validate"
      );
      return err(validationFailed);
    }
    const validationSuccess = getLocalizedString("plugins.appstudio.validationSucceedNotice");
    if (context.platform === Platform.VS) {
      context.logProvider.info(validationSuccess);
    } else {
      context.ui?.showMessage("info", validationSuccess, false);
    }*/
    return ok(new Map());
  }

  private loadCurrentState() {
    return {
      TAB_ENDPOINT: process.env.TAB_ENDPOINT,
      TAB_DOMAIN: process.env.TAB_DOMAIN,
      BOT_ID: process.env.BOT_ID,
      BOT_DOMAIN: process.env.BOT_DOMAIN,
      ENV_NAME: process.env.TEAMSFX_ENV,
    };
  }

  private validateArgs(args: ValidateTeamsAppArgs): Result<any, FxError> {
    if (!args || (!args.manifestPath && !args.appPackagePath)) {
      return err(
        AppStudioResultFactory.UserError(
          AppStudioError.InvalidParameterError.name,
          [
            getDefaultString("driver.teamsApp.validate.invalidParameter", actionName),
            getLocalizedString("driver.teamsApp.validate.invalidParameter", actionName),
          ],
          "https://aka.ms/teamsfx-actions/teamsapp-validate"
        )
      );
    }
    return ok(undefined);
  }
}
