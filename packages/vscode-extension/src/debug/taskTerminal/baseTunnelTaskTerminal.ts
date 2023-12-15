// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Xiaofu Huang <xiaofhua@microsoft.com>
 */
import * as util from "util";
import * as vscode from "vscode";
import { err, FxError, ok, Result, UserError, Void } from "@microsoft/teamsfx-api";
import { assembleError, envUtil, TunnelType } from "@microsoft/teamsfx-core";
import { Correlator } from "@microsoft/teamsfx-core";
import { LocalTelemetryReporter } from "@microsoft/teamsfx-core";
import { DotenvOutput } from "@microsoft/teamsfx-core";
import { pathUtils } from "@microsoft/teamsfx-core";
import VsCodeLogInstance from "../../commonlib/log";
import { ExtensionErrors, ExtensionSource } from "../../error";
import * as globalVariables from "../../globalVariables";
import { ProgressHandler } from "../../progressHandler";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../../telemetry/extTelemetryEvents";
import { getDefaultString, localize } from "../../utils/localizeUtils";
import { getLocalDebugSession, Step } from "../commonUtils";
import {
  baseTunnelDisplayMessages,
  RecommendedOperations,
  TunnelDisplayMessages,
} from "../constants";
import { doctorConstant } from "../depsChecker/doctorConstant";
import { localTelemetryReporter } from "../localTelemetryReporter";
import { BaseTaskTerminal } from "./baseTaskTerminal";

export interface IBaseTunnelArgs {
  type?: string;
  env?: string;
}

export type OutputInfo = {
  file: string | undefined;
};

type EndpointInfo = {
  src: string;
  dest: string;
  keys: string[];
};

export abstract class BaseTunnelTaskTerminal extends BaseTaskTerminal {
  protected static tunnelTaskTerminals: Map<string, BaseTunnelTaskTerminal> = new Map<
    string,
    BaseTunnelTaskTerminal
  >();

  protected outputMessageList: string[];

  protected readonly progressHandler: ProgressHandler;
  protected readonly step: Step;

  constructor(taskDefinition: vscode.TaskDefinition, stepNumber: number) {
    super(taskDefinition);

    for (const terminal of BaseTunnelTaskTerminal.tunnelTaskTerminals.values()) {
      terminal.close();
    }

    BaseTunnelTaskTerminal.tunnelTaskTerminals.set(this.taskTerminalId, this);

    this.progressHandler = new ProgressHandler(
      baseTunnelDisplayMessages.taskName,
      stepNumber,
      "terminal"
    );
    this.step = new Step(stepNumber);
    this.outputMessageList = [];
  }

  public do(): Promise<Result<Void, FxError>> {
    return Correlator.runWithId(getLocalDebugSession().id, () =>
      localTelemetryReporter.runWithTelemetryProperties(
        TelemetryEvent.DebugStartLocalTunnelTask,
        this.generateTelemetries(),
        () => this._do()
      )
    );
  }

  public static stopAll(): void {
    for (const task of BaseTunnelTaskTerminal.tunnelTaskTerminals.values()) {
      task.close();
    }
  }

  protected abstract generateTelemetries(): { [key: string]: string };
  protected abstract _do(): Promise<Result<Void, FxError>>;

  protected resolveArgs(args: IBaseTunnelArgs): void {
    if (!args) {
      throw BaseTaskTerminal.taskDefinitionError("args");
    }

    if (args.type) {
      if (typeof args.type !== "string" || !Object.values(TunnelType).includes(args.type)) {
        throw BaseTaskTerminal.taskDefinitionError("args.type");
      }
    }

    if (typeof args.env !== "undefined" && typeof args.env !== "string") {
      throw BaseTaskTerminal.taskDefinitionError("args.env");
    }
  }

  protected async outputStartMessage(tunnelDisplayMessages: TunnelDisplayMessages): Promise<void> {
    VsCodeLogInstance.info(tunnelDisplayMessages.title());
    VsCodeLogInstance.outputChannel.appendLine("");
    VsCodeLogInstance.outputChannel.appendLine(
      tunnelDisplayMessages.checkNumber(this.step.totalSteps)
    );
    VsCodeLogInstance.outputChannel.appendLine("");

    this.writeEmitter.fire(`${tunnelDisplayMessages.startTerminalMessage}\r\n\r\n`);

    await this.progressHandler.start();
  }

  protected async outputSuccessSummary(
    tunnelDisplayMessages: TunnelDisplayMessages,
    tunnelInfoArr: EndpointInfo[],
    envs: OutputInfo
  ): Promise<void> {
    const duration = this.getDurationInSeconds();
    VsCodeLogInstance.outputChannel.appendLine(tunnelDisplayMessages.summary());
    VsCodeLogInstance.outputChannel.appendLine("");

    for (const outputMessage of this.outputMessageList) {
      VsCodeLogInstance.outputChannel.appendLine(`${doctorConstant.Tick} ${outputMessage}`);
    }

    let isFirstTunnel = true;
    for (const tunnelInfo of tunnelInfoArr) {
      VsCodeLogInstance.outputChannel.appendLine(
        `${
          isFirstTunnel ? doctorConstant.Tick : doctorConstant.TickWhiteSpace
        } ${tunnelDisplayMessages.successSummary(
          tunnelInfo.src,
          tunnelInfo.dest,
          envs.file,
          tunnelInfo.keys
        )}`
      );
      isFirstTunnel = false;
    }

    VsCodeLogInstance.outputChannel.appendLine("");
    VsCodeLogInstance.outputChannel.appendLine(
      tunnelDisplayMessages.learnMore(tunnelDisplayMessages.learnMoreHelpLink)
    );
    VsCodeLogInstance.outputChannel.appendLine("");
    if (duration) {
      VsCodeLogInstance.info(tunnelDisplayMessages.durationMessage(duration));
    }

    for (const tunnelInfo of tunnelInfoArr) {
      this.writeEmitter.fire(
        `\r\n${tunnelDisplayMessages.terminalSuccessSummary(
          tunnelInfo.src,
          tunnelInfo.dest,
          envs.file,
          tunnelInfo.keys
        )}\r\n`
      );
    }
    this.writeEmitter.fire(`\r\n${tunnelDisplayMessages.successTerminalMessage}\r\n\r\n`);

    await this.progressHandler.end(true);

    localTelemetryReporter.sendTelemetryEvent(
      TelemetryEvent.DebugStartLocalTunnelTaskStarted,
      Object.assign(
        { [TelemetryProperty.Success]: TelemetrySuccess.Yes },
        this.generateTelemetries()
      ),
      {
        [LocalTelemetryReporter.PropertyDuration]: duration ?? -1,
      }
    );
  }

  protected async outputFailureSummary(
    tunnelDisplayMessages: TunnelDisplayMessages,
    error?: any
  ): Promise<void> {
    const fxError = error ? assembleError(error) : TunnelError.StartTunnelError();
    VsCodeLogInstance.outputChannel.appendLine(tunnelDisplayMessages.summary());

    VsCodeLogInstance.outputChannel.appendLine("");
    for (const outputMessage of this.outputMessageList) {
      VsCodeLogInstance.outputChannel.appendLine(`${doctorConstant.Tick} ${outputMessage}`);
    }

    VsCodeLogInstance.outputChannel.appendLine(`${doctorConstant.Cross} ${fxError.message}`);

    VsCodeLogInstance.outputChannel.appendLine("");
    VsCodeLogInstance.outputChannel.appendLine(
      tunnelDisplayMessages.learnMore(tunnelDisplayMessages.learnMoreHelpLink)
    );
    VsCodeLogInstance.outputChannel.appendLine("");

    this.writeEmitter.fire(`\r\n${tunnelDisplayMessages.errorTerminalMessage}\r\n`);

    await this.progressHandler.end(false);

    localTelemetryReporter.sendTelemetryErrorEvent(
      TelemetryEvent.DebugStartLocalTunnelTaskStarted,
      fxError,
      Object.assign(
        {
          [TelemetryProperty.Success]: TelemetrySuccess.No,
        },
        this.generateTelemetries()
      ),
      {
        [LocalTelemetryReporter.PropertyDuration]: this.getDurationInSeconds() ?? -1,
      }
    );
  }

  protected async savePropertiesToEnv(
    env: string | undefined,
    envVars: {
      [key: string]: string;
    }
  ): Promise<Result<OutputInfo, FxError>> {
    try {
      const result: OutputInfo = {
        file: undefined,
      };
      if (!globalVariables.workspaceUri?.fsPath || !env) {
        return ok(result);
      }

      if (Object.entries(envVars).length === 0) {
        return ok(result);
      }

      const res = await envUtil.writeEnv(globalVariables.workspaceUri.fsPath, env, envVars);
      const envFilePathResult = await pathUtils.getEnvFilePath(
        globalVariables.workspaceUri.fsPath,
        env
      );
      if (envFilePathResult.isOk()) {
        result.file = envFilePathResult.value;
      }
      return res.isOk() ? ok(result) : err(res.error);
    } catch (error: any) {
      return err(TunnelError.TunnelEnvError(error));
    }
  }

  protected async readPropertiesFromEnv(
    env: string | undefined
  ): Promise<Result<DotenvOutput, FxError>> {
    if (!globalVariables.workspaceUri?.fsPath || !env) {
      return ok({});
    }
    return await envUtil.readEnv(globalVariables.workspaceUri.fsPath, env, false, false);
  }
}

export const TunnelError = Object.freeze({
  TunnelEnvError: (error: Error) => {
    const userError = new UserError({
      source: ExtensionSource,
      name: ExtensionErrors.TunnelEnvError,
      message: `${getDefaultString("teamstoolkit.localDebug.tunnelEnvError")} ${
        error?.message ?? ""
      }`,
      displayMessage: `${localize("teamstoolkit.localDebug.tunnelEnvError")} ${
        error?.message ?? ""
      }`,
      helpLink: baseTunnelDisplayMessages.learnMoreHelpLink,
    });
    userError.recommendedOperation = RecommendedOperations.DebugInTestTool;
    return userError;
  },
  StartTunnelError: (error?: Error) => {
    const userError = new UserError({
      source: ExtensionSource,
      name: ExtensionErrors.StartTunnelError,
      message: `${getDefaultString("teamstoolkit.localDebug.startTunnelError")} ${
        error?.message ?? ""
      }`,
      displayMessage: `${localize("teamstoolkit.localDebug.startTunnelError")} ${
        error?.message ?? ""
      }`,
      helpLink: baseTunnelDisplayMessages.learnMoreHelpLink,
    });
    userError.recommendedOperation = RecommendedOperations.DebugInTestTool;
    return userError;
  },
  DevTunnelOperationError: (operationName: string, error?: Error) => {
    const userError = new UserError({
      source: ExtensionSource,
      name: ExtensionErrors.DevTunnelOperationError,
      message: `${util.format(
        getDefaultString("teamstoolkit.localDebug.devTunnelOperationError"),
        operationName
      )} ${error?.message ?? ""}`,
      displayMessage: `${util.format(
        localize("teamstoolkit.localDebug.devTunnelOperationError"),
        operationName
      )} ${error?.message ?? ""}`,
      helpLink: baseTunnelDisplayMessages.learnMoreHelpLink,
    });
    userError.recommendedOperation = RecommendedOperations.DebugInTestTool;
    return userError;
  },
  TunnelResourceLimitExceededError: (error: Error) => {
    const userError = new UserError({
      source: ExtensionSource,
      name: ExtensionErrors.TunnelResourceLimitExceededError,
      message: error?.message,
      displayMessage: error?.message,
      helpLink: baseTunnelDisplayMessages.learnMoreHelpLink,
    });
    userError.recommendedOperation = RecommendedOperations.DebugInTestTool;
    return userError;
  },
});
