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
  err,
  ok,
} from "@microsoft/teamsfx-api";
import {
  AppStudioScopes,
  DependencyStatus,
  DepsCheckerError,
  DepsManager,
  DepsType,
  LocalEnvManager,
  NodeNotFoundError,
  NodeNotLtsError,
  TelemetryContext,
  V3NodeNotSupportedError,
  assembleError,
  getSideloadingStatus,
  ErrorCategory,
  PackageService,
} from "@microsoft/teamsfx-core";
import * as os from "os";
import * as util from "util";
import * as vscode from "vscode";
import { signedOut } from "../../commonlib/common/constant";
import VsCodeLogInstance from "../../commonlib/log";
import M365TokenInstance from "../../commonlib/m365Login";
import { ExtensionErrors, ExtensionSource } from "../../error/error";
import { VS_CODE_UI } from "../../qm/vsc_ui";
import { tools, workspaceUri } from "../../globalVariables";
import { checkCopilotCallback } from "../../handlers/accounts/checkAccessCallback";
import { ProgressHandler } from "../progressHandler";
import { ExtTelemetry } from "../../telemetry/extTelemetry";
import { TelemetryEvent, TelemetryProperty } from "../../telemetry/extTelemetryEvents";
import { getDefaultString, localize } from "../../utils/localizeUtils";
import { Step } from "../common/step";
import { DisplayMessages, RecommendedOperations } from "../common/debugConstants";
import { doctorConstant } from "./doctorConstant";
import { vscodeLogger } from "./vscodeLogger";
import { vscodeTelemetry } from "./vscodeTelemetry";
import { localTelemetryReporter } from "../localTelemetryReporter";
import { ProgressHelper } from "../progressHelper";
import { WebviewPanel } from "../../controls/webviewPanel";
import { PanelType } from "../../controls/PanelType";
import {
  ResultStatus,
  Checker,
  ProgressMessage,
  copilotCheckServiceScope,
  DepsDisplayName,
} from "./prerequisitesCheckerConstants";
import {
  CheckResult,
  PrerequisiteOrderedChecker,
  PrerequisiteCheckerInfo,
  PortCheckerInfo,
} from "../common/types";

export async function _checkAndInstall(
  displayMessages: DisplayMessages,
  orderedCheckers: PrerequisiteOrderedChecker[],
  additionalTelemetryProperties: { [key: string]: string }
): Promise<{ checkResults: CheckResult[]; error?: FxError }> {
  let progressHelper: ProgressHelper | undefined;
  const checkResults: CheckResult[] = [];
  try {
    const enabledCheckers = parseCheckers(orderedCheckers);

    const localEnvManager = new LocalEnvManager(
      VsCodeLogInstance,
      ExtTelemetry.reporter,
      VS_CODE_UI
    );

    VsCodeLogInstance.outputChannel.show();
    VsCodeLogInstance.info(displayMessages.title);
    VsCodeLogInstance.outputChannel.appendLine("");

    // Get deps
    const depsManager = new DepsManager(vscodeLogger, vscodeTelemetry);

    const step = new Step(enabledCheckers.length);

    VsCodeLogInstance.outputChannel.appendLine(displayMessages.checkNumber(step.totalSteps));
    progressHelper = new ProgressHelper(
      new ProgressHandler(displayMessages.taskName, step.totalSteps)
    );

    await progressHelper.start(
      enabledCheckers.map((v) => {
        return {
          key: v.checker,
          detail: ProgressMessage[v.checker],
        };
      })
    );
    VsCodeLogInstance.outputChannel.appendLine("");

    for (const orderedChecker of orderedCheckers) {
      const orderedCheckerInfo = orderedChecker.info;
      const checkResult = await getCheckPromise(
        orderedCheckerInfo,
        depsManager,
        localEnvManager,
        step,
        additionalTelemetryProperties
        // eslint-disable-next-line @typescript-eslint/no-misused-promises
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
          // eslint-disable-next-line no-secrets/no-secrets
          getDefaultString("teamstoolkit.localDebug.portsAlreadyInUse"),
          formatPortStr(portsInUse)
        );
        const displayMessage = util.format(
          // eslint-disable-next-line no-secrets/no-secrets
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

function getCheckPromise(
  checkerInfo: PrerequisiteCheckerInfo,
  depsManager: DepsManager,
  localEnvManager: LocalEnvManager,
  step: Step,
  additionalTelemetryProperties: { [key: string]: string }
): Promise<CheckResult> {
  switch (checkerInfo.checker) {
    case DepsType.LtsNode:
    case DepsType.ProjectNode:
      return checkNode(
        checkerInfo.checker,
        depsManager,
        step.getPrefix(),
        additionalTelemetryProperties
      );
    case Checker.M365Account:
      return checkM365Account(step.getPrefix(), true, additionalTelemetryProperties);
    case Checker.CopilotAccess:
      return checkM365AccountCopilot(step.getPrefix(), true, additionalTelemetryProperties);
    case Checker.Ports:
      return checkPort(
        localEnvManager,
        (checkerInfo as PortCheckerInfo)?.ports ?? [],
        `${step.getPrefix()} ${ProgressMessage[Checker.Ports]} ...`,
        additionalTelemetryProperties
      );
  }
}

function parseCheckers(orderedCheckers: PrerequisiteOrderedChecker[]): PrerequisiteCheckerInfo[] {
  const parsedCheckers: PrerequisiteCheckerInfo[] = [];
  for (const orderedChecker of orderedCheckers) {
    parsedCheckers.push(orderedChecker.info);
  }
  return parsedCheckers;
}

function ensureM365Account(
  showLoginPage: boolean
): Promise<Result<{ token: string; tenantId?: string; loginHint?: string }, FxError>> {
  // Check M365 account token
  return localTelemetryReporter.runWithTelemetry(
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
        const e = new SystemError(
          ExtensionSource,
          ExtensionErrors.PrerequisitesNoM365AccountError,
          "No Microsoft 365 account login"
        );
        e.categories = [ErrorCategory.Internal];
        return err(e);
      }
      const loginHint = typeof upn === "string" ? upn : undefined;
      const tenantId = typeof tid === "string" ? tid : undefined;
      return ok({ token, tenantId, loginHint });
    }
  );
}

async function ensureCopilotAccess(
  showLoginPage: boolean
): Promise<Result<{ token: string; tenantId?: string; loginHint?: string }, FxError>> {
  const m365Result = await ensureM365Account(showLoginPage);
  if (m365Result.isErr()) {
    return err(m365Result.error);
  }

  // Check copilot access
  const copilotResult = await localTelemetryReporter.runWithTelemetry(
    TelemetryEvent.DebugPrereqsCheckM365Copilot,
    async (ctx: TelemetryContext) => {
      const m365Login: M365TokenProvider = M365TokenInstance;
      const copilotTokenRes = await m365Login.getAccessToken({
        scopes: [copilotCheckServiceScope],
        showDialog: false,
      });
      let hasCopilotAccess: boolean | undefined = undefined;
      if (copilotTokenRes.isOk()) {
        hasCopilotAccess = await PackageService.GetSharedInstance().getCopilotStatus(
          copilotTokenRes.value,
          false
        );
      }

      // true, false or undefined for error
      ctx.properties[TelemetryProperty.DebugHasCopilotAccess] = String(!!hasCopilotAccess);
      if (hasCopilotAccess === false) {
        // copilot disabled
        return err(
          new UserError(
            ExtensionSource,
            ExtensionErrors.PrerequisitesNoCopilotAccessError,
            getDefaultString("teamstoolkit.accountTree.copilotWarningTooltip"),
            localize("teamstoolkit.accountTree.copilotWarningTooltip")
          )
        );
      }

      return ok(undefined);
    }
  );
  if (copilotResult.isErr()) {
    return err(copilotResult.error);
  }

  return m365Result;
}

async function ensureSideloding(
  showLoginPage: boolean
): Promise<Result<{ token: string; tenantId?: string; loginHint?: string }, FxError>> {
  const m365Result = await ensureM365Account(showLoginPage);
  if (m365Result.isErr()) {
    return err(m365Result.error);
  }

  // Check sideloading permission
  const sideloadingResult = await localTelemetryReporter.runWithTelemetry(
    TelemetryEvent.DebugPrereqsCheckM365Sideloading,
    async (ctx: TelemetryContext) => {
      const isSideloadingEnabled = await getSideloadingStatus(m365Result.value.token);
      // true, false or undefined for error
      ctx.properties[TelemetryProperty.DebugIsSideloadingAllowed] = String(!!isSideloadingEnabled);
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

        const accountResult = await ensureSideloding(showLoginPage);
        if (accountResult.isErr()) {
          result = ResultStatus.failed;
          error = accountResult.error;
          WebviewPanel.createOrShow(PanelType.AccountHelp);
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

function checkM365AccountCopilot(
  prefix: string,
  showLoginPage: boolean,
  additionalTelemetryProperties: { [key: string]: string }
): Promise<CheckResult> {
  return runWithCheckResultTelemetryProperties(
    TelemetryEvent.DebugPrereqsCheckM365Copilot,
    additionalTelemetryProperties,
    async (): Promise<CheckResult> => {
      let result = ResultStatus.success;
      let error = undefined;
      let loginHint = undefined;
      const warnMsg = Checker.CopilotAccess;
      try {
        VsCodeLogInstance.outputChannel.appendLine(
          `${prefix} ${ProgressMessage[Checker.CopilotAccess]} ...`
        );

        const accountResult = await ensureCopilotAccess(showLoginPage);
        if (accountResult.isErr()) {
          result = ResultStatus.warn;
          error = accountResult.error;
          await checkCopilotCallback();
        } else {
          loginHint = accountResult.value.loginHint;
        }
      } catch (err: unknown) {
        result = ResultStatus.warn;
        if (!error) {
          error = assembleError(err);
        }
      }

      return {
        checker: Checker.CopilotAccess,
        result: result,
        successMsg:
          result && loginHint
            ? doctorConstant.SignInCopilotSuccess.split("@account").join(`${loginHint}`)
            : Checker.CopilotAccess,
        warnMsg: warnMsg,
        error: error,
      };
    }
  );
}

async function checkNode(
  nodeDep: DepsType.LtsNode | DepsType.ProjectNode,
  depsManager: DepsManager,
  prefix: string,
  additionalTelemetryProperties: { [key: string]: string }
): Promise<CheckResult> {
  return await runWithCheckResultTelemetryProperties(
    TelemetryEvent.DebugPrereqsCheckNode,
    additionalTelemetryProperties,
    async () => {
      try {
        VsCodeLogInstance.outputChannel.appendLine(`${prefix} ${ProgressMessage[nodeDep]} ...`);
        const nodeStatus = await depsManager.ensureDependency(nodeDep, true, {
          projectPath: workspaceUri?.fsPath,
        });
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
      const userError = new UserError(errorOptions);
      // Recommend to open test tool if M365 account check failed
      if (failures.find((f) => f.checker === Checker.M365Account)) {
        userError.recommendedOperation = RecommendedOperations.DebugInTestTool;
      }
      throw userError;
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
