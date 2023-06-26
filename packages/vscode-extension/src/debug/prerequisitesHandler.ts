// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Qianhao Dong <qidon@microsoft.com>
 */
import {
  FxError,
  M365TokenProvider,
  Result,
  SystemError,
  UserError,
  UserErrorOptions,
  Void,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import {
  AppStudioScopes,
  DependencyStatus,
  DepsCheckerError,
  DepsType,
  LocalEnvManager,
  NodeNotFoundError,
  NodeNotLtsError,
  Prerequisite,
  TelemetryContext,
  V3NodeNotSupportedError,
  assembleError,
  getSideloadingStatus,
  CheckerFactory,
} from "@microsoft/teamsfx-core";
import * as os from "os";
import * as util from "util";
import * as vscode from "vscode";
import { signedOut } from "../commonlib/common/constant";
import VsCodeLogInstance from "../commonlib/log";
import M365TokenInstance from "../commonlib/m365Login";
import { ExtensionErrors, ExtensionSource } from "../error";
import { VS_CODE_UI } from "../extension";
import * as globalVariables from "../globalVariables";
import { openAccountHelpHandler, tools } from "../handlers";
import { ProgressHandler } from "../progressHandler";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import { TelemetryEvent, TelemetryProperty } from "../telemetry/extTelemetryEvents";
import { getDefaultString, localize } from "../utils/localizeUtils";
import * as commonUtils from "./commonUtils";
import { Step } from "./commonUtils";
import {
  DisplayMessages,
  prerequisiteCheckForGetStartedDisplayMessages,
  v3PrerequisiteCheckTaskDisplayMessages,
} from "./constants";
import { doctorConstant } from "./depsChecker/doctorConstant";
import { localTelemetryReporter } from "./localTelemetryReporter";
import { ProgressHelper } from "./progressHelper";
import { allRunningTeamsfxTasks, terminateAllRunningTeamsfxTasks } from "./teamsfxTaskHandler";

enum Checker {
  M365Account = "Microsoft 365 Account",
  Ports = "ports occupancy",
}

const DepsDisplayName = {
  [DepsType.LtsNode]: "Node.js",
  [DepsType.ProjectNode]: "Node.js",
};

interface CheckResult {
  checker: string;
  result: ResultStatus;
  error?: FxError;
  successMsg?: string;
  warnMsg?: string;
  failureMsg?: string;
}

enum ResultStatus {
  success = "success",
  warn = "warn",
  failed = "failed",
}

const ProgressMessage = Object.freeze({
  [Checker.M365Account]: `Checking ${Checker.M365Account}`,
  [Checker.Ports]: `Checking ${Checker.Ports}`,
  [DepsType.LtsNode]: `Checking ${DepsDisplayName[DepsType.LtsNode]}`,
  [DepsType.ProjectNode]: `Checking ${DepsDisplayName[DepsType.ProjectNode]}`,
});

type PortCheckerInfo = { checker: Checker.Ports; ports: number[] };
type PrerequisiteCheckerInfo = {
  checker: Checker | Checker.M365Account | Checker.Ports | DepsType.LtsNode | DepsType.ProjectNode;
  [key: string]: any;
};

type PrerequisiteOrderedChecker = {
  info: PrerequisiteCheckerInfo;
  fastFail: boolean;
};

async function runWithCheckResultTelemetryProperties(
  eventName: string,
  initialProperties: { [key: string]: string },
  action: (ctx: TelemetryContext) => Promise<CheckResult>
): Promise<CheckResult> {
  return await localTelemetryReporter.runWithTelemetryGeneric(
    eventName,
    action,
    (result: CheckResult) => {
      return result.result === ResultStatus.success ? undefined : result.error;
    },
    initialProperties
  );
}

// Mainly addresses two issues:
// 1. Some error messages contain special characters which will cause the whole debug-check-results to be redacted.
// 2. CheckResult[] is hard to parse in kusto query (an array of objects).
//
// `debug-check-results` contains only known content and we know it will not be redacted.
// `debug-check-results-raw` might contain arbitrary string and be redacted.
function convertCheckResultsForTelemetry(checkResults: CheckResult[]): [string, string] {
  const resultRaw: { [checker: string]: unknown } = {};
  const resultSafe: { [checker: string]: { [key: string]: string | undefined } } = {};
  for (const checkResult of checkResults) {
    resultRaw[checkResult.checker] = checkResult;
    resultSafe[checkResult.checker] = {
      result: checkResult.result,
      source: checkResult.error?.source,
      errorCode: checkResult.error?.name,
      errorType:
        checkResult.error === undefined
          ? undefined
          : checkResult.error instanceof UserError
          ? "user"
          : checkResult.error instanceof SystemError
          ? "system"
          : "unknown",
    };
  }

  return [JSON.stringify(resultRaw), JSON.stringify(resultSafe)];
}

function addCheckResultsForTelemetry(
  checkResults: CheckResult[],
  properties: { [key: string]: string },
  errorProps: string[]
): void {
  const [resultRaw, resultSafe] = convertCheckResultsForTelemetry(checkResults);
  properties[TelemetryProperty.DebugCheckResultsSafe] = resultSafe;
  properties[TelemetryProperty.DebugCheckResults] = resultRaw;
  // only the raw event contains error message
  errorProps.push(TelemetryProperty.DebugCheckResults);
}

async function checkPort(
  localEnvManager: LocalEnvManager,
  ports: number[],
  displayMessage: string,
  additionalTelemetryProperties: { [key: string]: string }
): Promise<CheckResult> {
  return await runWithCheckResultTelemetryProperties(
    TelemetryEvent.DebugPrereqsCheckPorts,
    additionalTelemetryProperties,
    async (ctx: TelemetryContext) => {
      VsCodeLogInstance.outputChannel.appendLine(displayMessage);
      const portsInUse = await localEnvManager.getPortsInUse(ports);
      const formatPortStr = (ports: number[]) =>
        ports.length > 1 ? ports.join(", ") : `${ports[0]}`;
      if (portsInUse.length > 0) {
        ctx.properties[TelemetryProperty.DebugPortsInUse] = JSON.stringify(portsInUse);
        const message = util.format(
          getDefaultString("teamstoolkit.localDebug.portsAlreadyInUse"),
          formatPortStr(portsInUse)
        );
        const displayMessage = util.format(
          localize("teamstoolkit.localDebug.portsAlreadyInUse"),
          formatPortStr(portsInUse)
        );

        return {
          checker: Checker.Ports,
          result: ResultStatus.failed,
          failureMsg: doctorConstant.Port,
          error: new UserError(
            ExtensionSource,
            ExtensionErrors.PortAlreadyInUse,
            message,
            displayMessage
          ),
        };
      }
      return {
        checker: Checker.Ports,
        result: ResultStatus.success,
        successMsg: doctorConstant.PortSuccess.replace("@port", formatPortStr(ports)),
      };
    }
  );
}

export async function checkPrerequisitesForGetStarted(): Promise<Result<void, FxError>> {
  const nodeChecker = await getOrderedCheckersForGetStarted();
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.GetStartedPrerequisitesStart);
  const res = await _checkAndInstall(prerequisiteCheckForGetStartedDisplayMessages, nodeChecker, {
    [TelemetryProperty.DebugIsTransparentTask]: "false",
  });
  if (res.error) {
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.GetStartedPrerequisites, res.error);
    return err(res.error);
  }
  return ok(undefined);
}

export async function checkAndInstallForTask(
  prerequisites: string[],
  ports: number[] | undefined,
  telemetryProperties: { [key: string]: string }
): Promise<Result<Void, FxError>> {
  const orderedCheckers = await getOrderedCheckersForTask(prerequisites, ports);

  const additionalTelemetryProperties = Object.assign(
    {
      [TelemetryProperty.DebugIsTransparentTask]: "true",
    },
    telemetryProperties
  );
  return await localTelemetryReporter.runWithTelemetryProperties(
    TelemetryEvent.DebugPrerequisites,
    additionalTelemetryProperties,
    async (ctx: TelemetryContext) => {
      // terminate all running teamsfx tasks
      if (allRunningTeamsfxTasks.size > 0) {
        VsCodeLogInstance.info("Terminate all running teamsfx tasks.");
        terminateAllRunningTeamsfxTasks();
      }

      const res = await _checkAndInstall(
        v3PrerequisiteCheckTaskDisplayMessages,
        orderedCheckers,
        additionalTelemetryProperties
      );
      if (res.error) {
        const debugSession = commonUtils.getLocalDebugSession();
        addCheckResultsForTelemetry(
          res.checkResults,
          debugSession.properties,
          debugSession.errorProps
        );
        addCheckResultsForTelemetry(res.checkResults, ctx.properties, ctx.errorProps);
        return err(res.error);
      }
      return ok(Void);
    }
  );
}

async function _checkAndInstall(
  displayMessages: DisplayMessages,
  orderedCheckers: PrerequisiteOrderedChecker[],
  additionalTelemetryProperties: { [key: string]: string }
): Promise<{ checkResults: CheckResult[]; error?: FxError }> {
  let progressHelper: ProgressHelper | undefined;
  const checkResults: CheckResult[] = [];
  try {
    const localEnvManager = new LocalEnvManager(
      VsCodeLogInstance,
      ExtTelemetry.reporter,
      VS_CODE_UI
    );

    VsCodeLogInstance.outputChannel.show();
    VsCodeLogInstance.info(displayMessages.title);
    VsCodeLogInstance.outputChannel.appendLine("");

    const step = new Step(orderedCheckers.length);

    VsCodeLogInstance.outputChannel.appendLine(displayMessages.checkNumber(step.totalSteps));
    progressHelper = new ProgressHelper(
      new ProgressHandler(displayMessages.taskName, step.totalSteps)
    );

    await progressHelper.start(
      orderedCheckers.map((v) => {
        return {
          key: v.info.checker,
          detail: ProgressMessage[v.info.checker],
        };
      })
    );
    VsCodeLogInstance.outputChannel.appendLine("");

    for (const orderedChecker of orderedCheckers) {
      const orderedCheckerInfo = orderedChecker.info as PrerequisiteCheckerInfo;
      const checkResult = await getCheckPromise(
        orderedCheckerInfo,
        localEnvManager,
        step,
        additionalTelemetryProperties
      ).finally(async () => await progressHelper?.end(orderedCheckerInfo.checker));
      checkResults.push(checkResult);
      if (orderedChecker.fastFail) {
        await checkFailure(checkResults, displayMessages, progressHelper);
      }
    }
    await handleCheckResults(checkResults, displayMessages, progressHelper);
  } catch (error: unknown) {
    const fxError = assembleError(error);
    await progressHelper?.stop(false);
    return { checkResults: checkResults, error: fxError };
  }
  return { checkResults: checkResults };
}

function getCheckPromise(
  checkerInfo: PrerequisiteCheckerInfo,
  localEnvManager: LocalEnvManager,
  step: Step,
  additionalTelemetryProperties: { [key: string]: string }
): Promise<CheckResult> {
  switch (checkerInfo.checker) {
    case DepsType.LtsNode:
    case DepsType.ProjectNode:
      return checkNode(checkerInfo.checker, step.getPrefix(), additionalTelemetryProperties);
    case Checker.M365Account:
      return checkM365Account(step.getPrefix(), true, additionalTelemetryProperties);
    case Checker.Ports:
      return checkPort(
        localEnvManager,
        (checkerInfo as PortCheckerInfo)?.ports ?? [],
        `${step.getPrefix()} ${ProgressMessage[Checker.Ports]} ...`,
        additionalTelemetryProperties
      );
  }
}

async function ensureM365Account(
  showLoginPage: boolean
): Promise<Result<{ token: string; tenantId?: string; loginHint?: string }, FxError>> {
  // Check M365 account token
  const m365Result = await localTelemetryReporter.runWithTelemetry(
    TelemetryEvent.DebugPrereqsCheckM365AccountSignIn,
    async (
      ctx: TelemetryContext
    ): Promise<Result<{ token: string; tenantId?: string; loginHint?: string }, FxError>> => {
      const m365Login: M365TokenProvider = M365TokenInstance;
      let loginStatusRes = await m365Login.getStatus({ scopes: AppStudioScopes });
      if (loginStatusRes.isErr()) {
        ctx.properties[TelemetryProperty.DebugM365AccountStatus] = "error";
        return err(loginStatusRes.error);
      }
      ctx.properties[TelemetryProperty.DebugM365AccountStatus] = loginStatusRes.value.status;

      let token = loginStatusRes.value.token;
      let upn = loginStatusRes.value.accountInfo?.upn;
      let tid = loginStatusRes.value.accountInfo?.tid;
      if (loginStatusRes.value.status === signedOut && showLoginPage) {
        const tokenRes = await tools.tokenProvider.m365TokenProvider.getAccessToken({
          scopes: AppStudioScopes,
          showDialog: true,
        });
        if (tokenRes.isErr()) {
          return err(tokenRes.error);
        }
        loginStatusRes = await m365Login.getStatus({ scopes: AppStudioScopes });
        if (loginStatusRes.isErr()) {
          return err(loginStatusRes.error);
        }
        token = loginStatusRes.value.token;
        upn = loginStatusRes.value.accountInfo?.upn;
        tid = loginStatusRes.value.accountInfo?.tid;
      }
      if (token === undefined) {
        // corner case but need to handle
        return err(
          new SystemError(
            ExtensionSource,
            ExtensionErrors.PrerequisitesNoM365AccountError,
            "No Microsoft 365 account login"
          )
        );
      }
      const loginHint = typeof upn === "string" ? upn : undefined;
      const tenantId = typeof tid === "string" ? tid : undefined;
      return ok({ token, tenantId, loginHint });
    }
  );
  if (m365Result.isErr()) {
    return err(m365Result.error);
  }

  // Check sideloading permission
  const sideloadingResult = await localTelemetryReporter.runWithTelemetry(
    TelemetryEvent.DebugPrereqsCheckM365Sideloading,
    async (ctx: TelemetryContext) => {
      const isSideloadingEnabled = await getSideloadingStatus(m365Result.value.token);
      // true, false or undefined for error
      ctx.properties[TelemetryProperty.DebugIsSideloadingAllowed] = `${isSideloadingEnabled}`;
      if (isSideloadingEnabled === false) {
        // sideloading disabled
        return err(
          new UserError(
            ExtensionSource,
            ExtensionErrors.PrerequisitesSideloadingDisabledError,
            getDefaultString("teamstoolkit.accountTree.sideloadingWarningTooltip"),
            localize("teamstoolkit.accountTree.sideloadingWarningTooltip")
          )
        );
      }

      return ok(undefined);
    }
  );
  if (sideloadingResult.isErr()) {
    return err(sideloadingResult.error);
  }

  return m365Result;
}

function checkM365Account(
  prefix: string,
  showLoginPage: boolean,
  additionalTelemetryProperties: { [key: string]: string }
): Promise<CheckResult> {
  return runWithCheckResultTelemetryProperties(
    TelemetryEvent.DebugPrereqsCheckM365Account,
    additionalTelemetryProperties,
    async (): Promise<CheckResult> => {
      let result = ResultStatus.success;
      let error = undefined;
      let loginHint = undefined;
      const failureMsg = Checker.M365Account;
      try {
        VsCodeLogInstance.outputChannel.appendLine(
          `${prefix} ${ProgressMessage[Checker.M365Account]} ...`
        );

        const accountResult = await ensureM365Account(showLoginPage);
        if (accountResult.isErr()) {
          result = ResultStatus.failed;
          error = accountResult.error;
          openAccountHelpHandler();
        } else {
          loginHint = accountResult.value.loginHint;
        }
      } catch (err: unknown) {
        result = ResultStatus.failed;
        if (!error) {
          error = assembleError(err);
        }
      }

      return {
        checker: Checker.M365Account,
        result: result,
        successMsg:
          result && loginHint
            ? doctorConstant.SignInSuccess.split("@account").join(`${loginHint}`)
            : Checker.M365Account,
        failureMsg: failureMsg,
        error: error,
      };
    }
  );
}

async function checkNode(
  nodeDep: DepsType.LtsNode | DepsType.ProjectNode,
  prefix: string,
  additionalTelemetryProperties: { [key: string]: string }
): Promise<CheckResult> {
  return await runWithCheckResultTelemetryProperties(
    TelemetryEvent.DebugPrereqsCheckNode,
    additionalTelemetryProperties,
    async (ctx: TelemetryContext) => {
      try {
        VsCodeLogInstance.outputChannel.appendLine(`${prefix} ${ProgressMessage[nodeDep]} ...`);
        const nodeChecker = CheckerFactory.createChecker(nodeDep);
        const nodeStatus = await nodeChecker.resolve({
          projectPath: globalVariables.workspaceUri?.fsPath,
        });
        ctx.properties[TelemetryProperty.DebugNodeVersion] = `${nodeStatus.details.installVersion}`;
        return {
          checker: nodeStatus.name,
          result: nodeStatus.isInstalled
            ? nodeStatus.error
              ? ResultStatus.warn
              : ResultStatus.success
            : ResultStatus.failed,
          successMsg: nodeStatus.isInstalled
            ? doctorConstant.NodeSuccess.split("@Version").join(nodeStatus.details.installVersion)
            : nodeStatus.name,
          failureMsg: nodeStatus.name,
          error: nodeStatus.error
            ? handleDepsCheckerError(nodeStatus.error, nodeStatus)
            : undefined,
        };
      } catch (error: unknown) {
        return {
          checker: DepsDisplayName[nodeDep],
          result: ResultStatus.failed,
          successMsg: DepsDisplayName[nodeDep],
          failureMsg: DepsDisplayName[nodeDep],
          error: handleDepsCheckerError(error),
        };
      }
    }
  );
}

function handleDepsCheckerError(error: any, dep?: DependencyStatus): FxError {
  if (dep) {
    if (error instanceof NodeNotFoundError) {
      handleNodeNotFoundError(error);
    }
    if (error instanceof V3NodeNotSupportedError) {
      handleNodeNotLtsError(error);
    }
    if (error instanceof NodeNotLtsError) {
      handleV3NodeNotSupportedError(error);
    }
  }
  return error instanceof DepsCheckerError
    ? new UserError({
        error,
        source: ExtensionSource,
        name: ExtensionErrors.PrerequisitesValidationError,
        helpLink: error.helpLink,
      })
    : assembleError(error);
}

function handleNodeNotFoundError(error: NodeNotFoundError) {
  error.message = `${doctorConstant.NodeNotFound}${os.EOL}${doctorConstant.WhiteSpace}${doctorConstant.RestartVSCode}`;
}

function handleV3NodeNotSupportedError(error: V3NodeNotSupportedError) {
  error.message = `${error.message}${os.EOL}${doctorConstant.WhiteSpace}${doctorConstant.RestartVSCode}`;
}

function handleNodeNotLtsError(error: V3NodeNotSupportedError) {
  error.message = `${error.message}${os.EOL}${doctorConstant.WhiteSpace}${doctorConstant.RestartVSCode}`;
}

async function handleCheckResults(
  results: CheckResult[],
  displayMessages: DisplayMessages,
  progressHelper?: ProgressHelper,
  fromLocalDebug = true
): Promise<void> {
  if (results.length <= 0) {
    return;
  }

  let shouldStop = false;
  const output = VsCodeLogInstance.outputChannel;
  const successes = results.filter((a) => a.result === ResultStatus.success);
  const failures = results.filter((a) => a.result === ResultStatus.failed);
  const warnings = results.filter((a) => a.result === ResultStatus.warn);
  output.show();
  output.appendLine("");
  output.appendLine(displayMessages.summary);

  if (failures.length > 0) {
    shouldStop = true;
  }
  if (successes.length > 0) {
    output.appendLine("");
  }

  for (const result of successes) {
    output.appendLine(`${doctorConstant.Tick} ${result.successMsg ?? result.checker} `);
  }

  for (const result of warnings) {
    output.appendLine("");
    output.appendLine(`${doctorConstant.Exclamation} ${result.warnMsg ?? result.checker} `);
    outputCheckResultError(result, output);
  }

  for (const result of failures) {
    output.appendLine("");
    output.appendLine(`${doctorConstant.Cross} ${result.failureMsg ?? result.checker}`);
    outputCheckResultError(result, output);
  }
  output.appendLine("");
  output.appendLine(displayMessages.learnMore(displayMessages.learnMoreHelpLink));
  output.appendLine("");

  if (fromLocalDebug) {
    if (!shouldStop) {
      if (displayMessages.launchServices) {
        output.appendLine(displayMessages.launchServices);
        output.appendLine("");
      }
      await progressHelper?.stop(true);
    }

    if (shouldStop) {
      await progressHelper?.stop(false);
      const message =
        getDefaultString(displayMessages.errorMessageKey) +
        " " +
        displayMessages.showDetailMessage();

      // show failure summary in display message
      const displayMessage =
        util.format(
          localize("teamstoolkit.localDebug.failedCheckers"),
          failures.map((f) => f.failureMsg ?? f.checker).join(", ")
        ) +
        localize(displayMessages.errorDisplayMessageKey) +
        " " +
        displayMessages.showDetailDisplayMessage();

      const errorOptions: UserErrorOptions = {
        source: ExtensionSource,
        name: displayMessages.errorName,
        message: message,
        displayMessage: displayMessage,
        helpLink: displayMessages.errorHelpLink,
      };
      throw new UserError(errorOptions);
    }
  }
}

function outputCheckResultError(result: CheckResult, output: vscode.OutputChannel) {
  if (result.error) {
    output.appendLine(`${doctorConstant.WhiteSpace}${result.error.message}`);
  }
}

async function checkFailure(
  checkResults: CheckResult[],
  displayMessages: DisplayMessages,
  progressHelper?: ProgressHelper
) {
  if (checkResults.some((r) => r.result === ResultStatus.failed)) {
    await handleCheckResults(checkResults, displayMessages, progressHelper);
  }
}

async function getOrderedCheckersForGetStarted(): Promise<PrerequisiteOrderedChecker[]> {
  const workspacePath = globalVariables.workspaceUri?.fsPath;
  return [
    {
      info: { checker: workspacePath ? DepsType.ProjectNode : DepsType.LtsNode },
      fastFail: false,
    },
  ];
}

async function getOrderedCheckersForTask(
  prerequisites: string[],
  ports?: number[]
): Promise<PrerequisiteOrderedChecker[]> {
  const checkers: PrerequisiteOrderedChecker[] = [];
  if (prerequisites.includes(Prerequisite.nodejs)) {
    checkers.push({ info: { checker: DepsType.ProjectNode }, fastFail: true });
  }
  if (prerequisites.includes(Prerequisite.m365Account)) {
    checkers.push({ info: { checker: Checker.M365Account }, fastFail: false });
  }

  if (prerequisites.includes(Prerequisite.portOccupancy)) {
    checkers.push({ info: { checker: Checker.Ports, ports: ports }, fastFail: false });
  }
  return checkers;
}
