// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Xiaofu Huang <xiaofhua@microsoft.com>
 */
import * as path from "path";
import semver from "semver";
import { Service } from "typedi";
import { FxError, Result } from "@microsoft/teamsfx-api";
import { DependencyStatus } from "../../../common/deps-checker";
import {
  LocalCertificate,
  LocalCertificateManager,
} from "../../../common/local/localCertificateManager";
import { wrapRun } from "../../utils/common";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { WrapDriverContext } from "../util/wrapUtil";
import {
  Summaries,
  TelemetryDepsCheckStatus,
  TelemetryDevCertStatus,
  TelemetryProperties,
  toolsInstallDescription,
} from "./constant";
import { DotnetInstallationUserError } from "./error/dotnetInstallationUserError";
import { FuncInstallationUserError } from "./error/funcInstallationUserError";
import { InstallToolArgs } from "./interfaces/InstallToolArgs";
import { InvalidActionInputError } from "../../../error/common";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { updateProgress } from "../middleware/updateProgress";
import { hooks } from "@feathersjs/hooks/lib";
import { getLocalizedString } from "../../../common/localizeUtils";
import { FuncToolChecker } from "../../../common/deps-checker/internal/funcToolChecker";
import { DotnetChecker } from "../../../common/deps-checker/internal/dotnetChecker";

const ACTION_NAME = "devTool/install";
const helpLink = "https://aka.ms/teamsfx-actions/devtool-install";

const outputKeys = {
  sslCertFile: "sslCertFile",
  sslKeyFile: "sslKeyFile",
  funcPath: "funcPath",
  dotnetPath: "dotnetPath",
};

@Service(ACTION_NAME)
export class ToolsInstallDriver implements StepDriver {
  description = toolsInstallDescription();

  async run(
    args: InstallToolArgs,
    context: DriverContext
  ): Promise<Result<Map<string, string>, FxError>> {
    const wrapContext = new WrapDriverContext(context, ACTION_NAME, ACTION_NAME);
    return await this._run(args, wrapContext);
  }

  async execute(
    args: InstallToolArgs,
    context: DriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    const wrapContext = new WrapDriverContext(context, ACTION_NAME, ACTION_NAME);
    const result = await this._run(args, wrapContext, outputEnvVarNames);
    return {
      result: result,
      summaries: wrapContext.summaries,
    };
  }

  @hooks([
    addStartAndEndTelemetry(ACTION_NAME, ACTION_NAME),
    updateProgress(getLocalizedString("driver.prerequisite.progressBar")),
  ])
  async _run(
    args: InstallToolArgs,
    wrapContext: WrapDriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<Result<Map<string, string>, FxError>> {
    return wrapRun(async () => {
      const impl = new ToolsInstallDriverImpl(wrapContext);
      return await impl.run(args, outputEnvVarNames);
    });
  }
}

export class ToolsInstallDriverImpl {
  constructor(private context: WrapDriverContext) {}

  async run(
    args: InstallToolArgs,
    outputEnvVarNames?: Map<string, string>
  ): Promise<Map<string, string>> {
    const res = new Map<string, string>();
    this.validateArgs(args);

    this.setArgTelemetry(args);

    if (args.devCert) {
      const localCertRes = await this.resolveLocalCertificate(
        args.devCert.trust,
        outputEnvVarNames
      );
      localCertRes.forEach((v, k) => res.set(k, v));
    }

    if (args.func) {
      const funcRes = await this.resolveFuncCoreTools(
        `${args.func.version}`,
        args.func.symlinkDir,
        outputEnvVarNames
      );
      funcRes.forEach((v, k) => res.set(k, v));
    }

    if (args.dotnet) {
      const dotnetRes = await this.resolveDotnet(outputEnvVarNames);
      dotnetRes.forEach((v, k) => res.set(k, v));
    }

    return res;
  }

  async resolveLocalCertificate(
    trustDevCert: boolean,
    outputEnvVarNames?: Map<string, string>
  ): Promise<Map<string, string>> {
    const res = new Map<string, string>();
    // Do not print any log in LocalCertificateManager, use the error message returned instead.
    const certManager = new LocalCertificateManager(this.context.ui);
    const localCertResult = await certManager.setupCertificate(trustDevCert);
    if (trustDevCert) {
      let name = outputEnvVarNames?.get(outputKeys.sslCertFile);
      if (name) {
        res.set(name, localCertResult.certPath);
      }
      name = outputEnvVarNames?.get(outputKeys.sslKeyFile);
      if (name) {
        res.set(name, localCertResult.keyPath);
      }
    }
    this.setDevCertTelemetry(trustDevCert, localCertResult);

    if (typeof localCertResult.isTrusted === "undefined") {
      this.context.logProvider.warning(Summaries.devCertSkipped());
      this.context.addSummary(Summaries.devCertSkipped());
    } else if (localCertResult.isTrusted === false) {
      throw localCertResult.error;
    } else {
      this.context.addSummary(Summaries.devCertSuccess(trustDevCert));
    }

    return res;
  }

  async resolveFuncCoreTools(
    version: string,
    symlinkDir?: string,
    outputEnvVarNames?: Map<string, string>
  ): Promise<Map<string, string>> {
    const res = new Map<string, string>();
    const funcToolChecker = new FuncToolChecker();
    const funcStatus = await funcToolChecker.resolve({
      version: version,
      symlinkDir: symlinkDir,
      projectPath: this.context.projectPath,
    });

    this.setDepsCheckTelemetry(TelemetryProperties.funcStatus, funcStatus);

    if (!funcStatus.isInstalled && funcStatus.error) {
      throw new FuncInstallationUserError(ACTION_NAME, funcStatus.error, funcStatus.error.helpLink);
    } else if (funcStatus.error) {
      this.context.logProvider.warning(funcStatus.error.message);
      this.context.addSummary(
        Summaries.funcSuccess(funcStatus.details.binFolders) + funcStatus.error.message
      );
    } else {
      this.context.addSummary(Summaries.funcSuccess(funcStatus.details.binFolders));
    }

    if (funcStatus?.details?.binFolders !== undefined) {
      const funcBinFolder = funcStatus.details.binFolders.join(path.delimiter);
      const name = outputEnvVarNames?.get(outputKeys.funcPath);
      if (name) {
        res.set(name, funcBinFolder);
      }
    }
    return res;
  }

  async resolveDotnet(outputEnvVarNames?: Map<string, string>): Promise<Map<string, string>> {
    const res = new Map<string, string>();
    const dotnetChecker = new DotnetChecker();
    const dotnetStatus = await dotnetChecker.resolve();

    this.setDepsCheckTelemetry(TelemetryProperties.dotnetStatus, dotnetStatus);

    if (!dotnetStatus.isInstalled && dotnetStatus.error) {
      throw new DotnetInstallationUserError(
        ACTION_NAME,
        dotnetStatus.error,
        dotnetStatus.error.helpLink
      );
    } else if (dotnetStatus.error) {
      this.context.logProvider.warning(dotnetStatus.error?.message);
      this.context.addSummary(dotnetStatus.error?.message);
    } else {
      this.context.addSummary(Summaries.dotnetSuccess(dotnetStatus?.details?.binFolders));
    }
    if (dotnetStatus?.details?.binFolders !== undefined) {
      const dotnetBinFolder = `${dotnetStatus.details.binFolders
        .map((f: string) => path.dirname(f))
        .join(path.delimiter)}`;
      const name = outputEnvVarNames?.get(outputKeys.dotnetPath);
      if (name) {
        res.set(name, dotnetBinFolder);
      }
    }
    return res;
  }

  private validateArgs(args: InstallToolArgs): void {
    if (!!args.devCert && typeof args.devCert?.trust !== "boolean") {
      throw new InvalidActionInputError(ACTION_NAME, ["devCert.trust"], helpLink);
    }
    if (typeof args.func !== "undefined") {
      if (typeof args.func !== "object") {
        throw new InvalidActionInputError(ACTION_NAME, ["func"], helpLink);
      }
      if (
        (typeof args.func.version !== "string" && typeof args.func.version !== "number") ||
        !semver.validRange(`${args.func?.version}`)
      ) {
        throw new InvalidActionInputError(ACTION_NAME, ["func.version"], helpLink);
      }
      if (typeof args.func.symlinkDir !== "string" && typeof args.func.symlinkDir !== "undefined") {
        throw new InvalidActionInputError(ACTION_NAME, ["func.symlinkDir"], helpLink);
      }
    }
    if (!!args.dotnet && typeof args.dotnet !== "boolean") {
      throw new InvalidActionInputError(ACTION_NAME, ["dotnet"], helpLink);
    }
  }

  private setArgTelemetry(args: InstallToolArgs): void {
    this.context.addTelemetryProperties({
      [TelemetryProperties.driverArgs]: JSON.stringify({
        devCert: args.devCert,
        func: {
          version: args.func?.version,
          symlinkDir: args.func?.symlinkDir
            ? path.resolve(args.func.symlinkDir) === path.resolve("./devTools/func")
              ? "<default>"
              : "<unknown>"
            : "<undefined>",
        },
        dotnet: args.dotnet,
      }),
    });
  }

  private setDevCertTelemetry(trustDevCert: boolean, localCertResult: LocalCertificate) {
    this.context.addTelemetryProperties({
      [TelemetryProperties.devCertStatus]: !trustDevCert
        ? TelemetryDevCertStatus.Disabled
        : localCertResult.alreadyTrusted
        ? TelemetryDevCertStatus.AlreadyTrusted
        : localCertResult.isTrusted
        ? TelemetryDevCertStatus.Trusted
        : TelemetryDevCertStatus.NotTrusted,
    });
  }

  private setDepsCheckTelemetry(
    propertyName: typeof TelemetryProperties[keyof typeof TelemetryProperties],
    depStatus: DependencyStatus
  ): void {
    this.context.addTelemetryProperties({
      [propertyName]: depStatus.isInstalled
        ? depStatus.error
          ? TelemetryDepsCheckStatus.warn
          : TelemetryDepsCheckStatus.success
        : TelemetryDepsCheckStatus.failed,
    });
    if (depStatus.telemetryProperties) {
      this.context.addTelemetryProperties(depStatus.telemetryProperties);
    }
  }
}
