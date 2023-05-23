/**
 * @author huajiezhang <huajiezhang@microsoft.com>
 */
import { err, FxError, ok, Result, LogProvider } from "@microsoft/teamsfx-api";
import { Service } from "typedi";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { hooks } from "@feathersjs/hooks";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { TelemetryConstant } from "../../constant/commonConstant";
import { ProgressMessages } from "../../messages";
import { DotenvOutput } from "../../utils/envUtil";
import { ScriptExecutionError, ScriptTimeoutError } from "../../../error/script";
import { getSystemEncoding } from "../../utils/charsetUtils";
import * as path from "path";
import os from "os";
import fs from "fs-extra";
import iconv from "iconv-lite";
import child_process from "child_process";

const ACTION_NAME = "script";
const SET_ENV_CMD1 = "::set-output ";
const SET_ENV_CMD2 = "::set-teamsfx-env ";

interface ScriptDriverArgs {
  run: string;
  workingDirectory?: string;
  shell?: string;
  timeout?: number;
  redirectTo?: string;
}

@Service(ACTION_NAME)
export class ScriptDriver implements StepDriver {
  @hooks([addStartAndEndTelemetry(ACTION_NAME, TelemetryConstant.SCRIPT_COMPONENT)])
  async run(args: unknown, context: DriverContext): Promise<Result<Map<string, string>, FxError>> {
    const typedArgs = args as ScriptDriverArgs;
    await context.progressBar?.next(
      ProgressMessages.runCommand(typedArgs.run, typedArgs.workingDirectory ?? "./")
    );
    const res = await executeCommand(
      typedArgs.run,
      context.projectPath,
      context.logProvider,
      context.ui,
      typedArgs.workingDirectory,
      undefined,
      typedArgs.shell,
      typedArgs.timeout,
      typedArgs.redirectTo
    );
    if (res.isErr()) return err(res.error);
    const outputs = res.value[1];
    const kvArray: [string, string][] = Object.keys(outputs).map((k) => [k, outputs[k]]);
    return ok(new Map(kvArray));
  }

  @hooks([addStartAndEndTelemetry(ACTION_NAME, TelemetryConstant.SCRIPT_COMPONENT)])
  async execute(args: unknown, ctx: DriverContext): Promise<ExecutionResult> {
    const res = await this.run(args, ctx);
    const summaries: string[] = res.isOk()
      ? [`Successfully executed command ${maskSecretValues((args as any).run)}`]
      : [];
    return { result: res, summaries: summaries };
  }
}

export const scriptDriver = new ScriptDriver();

export async function executeCommand(
  command: string,
  projectPath: string,
  logProvider: LogProvider,
  ui: DriverContext["ui"],
  workingDirectory?: string,
  env?: NodeJS.ProcessEnv,
  shell?: string,
  timeout?: number,
  redirectTo?: string
): Promise<Result<[string, DotenvOutput], FxError>> {
  return new Promise(async (resolve, reject) => {
    const platform = os.platform();
    let workingDir = workingDirectory || ".";
    workingDir = path.isAbsolute(workingDir) ? workingDir : path.join(projectPath, workingDir);
    if (platform === "win32") {
      workingDir = capitalizeFirstLetter(path.resolve(workingDir ?? ""));
    }
    const defaultOsToShellMap: any = {
      win32: "powershell",
      darwin: "bash",
      linux: "bash",
    };
    let run = command;
    shell = shell || defaultOsToShellMap[platform] || "pwsh";
    let appendFile: string | undefined = undefined;
    if (redirectTo) {
      appendFile = path.isAbsolute(redirectTo) ? redirectTo : path.join(projectPath, redirectTo);
    }
    if (shell === "cmd") {
      run = `%ComSpec% /D /E:ON /V:OFF /S /C "CALL ${command}"`;
    }
    await logProvider.info(`Start to run command: "${command}" on path: "${workingDir}".`);
    const allOutputStrings: string[] = [];
    const systemEncoding = await getSystemEncoding();
    const stderrStrings: string[] = [];
    const cp = child_process.exec(
      run,
      {
        shell: shell,
        cwd: workingDir,
        encoding: "buffer",
        env: { ...process.env, ...env },
        timeout: timeout,
      },
      async (error) => {
        if (error) {
          error.message = stderrStrings.join("").trim() || error.message;
          resolve(err(convertScriptErrorToFxError(error, run)));
        } else {
          // handle '::set-output' or '::set-teamsfx-env' pattern
          const outputString = allOutputStrings.join("");
          const outputObject = parseSetOutputCommand(outputString);
          resolve(ok([outputString, outputObject]));
        }
      }
    );
    const dataHandler = (data: string) => {
      if (appendFile) {
        fs.appendFileSync(appendFile, data);
      }
      allOutputStrings.push(data);
    };
    cp.stdout?.on("data", (data: Buffer) => {
      const str = bufferToString(data, systemEncoding);
      logProvider.info(` [script action stdout] ${maskSecretValues(str)}`);
      dataHandler(str);
    });
    cp.stderr?.on("data", (data: Buffer) => {
      const str = bufferToString(data, systemEncoding);
      logProvider.warning(` [script action stderr] ${maskSecretValues(str)}`);
      dataHandler(str);
      stderrStrings.push(str);
    });
  });
}

export function bufferToString(data: Buffer, systemEncoding: string): string {
  const str =
    systemEncoding === "utf8" || systemEncoding === "utf-8"
      ? data.toString()
      : iconv.decode(data, systemEncoding);
  return str;
}

export function convertScriptErrorToFxError(
  error: child_process.ExecException,
  run: string
): ScriptTimeoutError | ScriptExecutionError {
  if (error.killed) {
    return new ScriptTimeoutError(run);
  } else {
    return new ScriptExecutionError(run, error.message);
  }
}

export function parseSetOutputCommand(stdout: string): DotenvOutput {
  const lines = stdout.toString().replace(/\r\n?/gm, "\n").split(/\r?\n/);
  const output: DotenvOutput = {};
  for (const line of lines) {
    if (line.startsWith(SET_ENV_CMD1) || line.startsWith(SET_ENV_CMD2)) {
      const str = line.startsWith(SET_ENV_CMD1)
        ? line.substring(SET_ENV_CMD1.length).trim()
        : line.substring(SET_ENV_CMD2.length).trim();
      const arr = str.split("=");
      if (arr.length === 2) {
        const key = arr[0].trim();
        const value = arr[1].trim();
        output[key] = value;
      }
    }
  }
  return output;
}

export function maskSecretValues(stdout: string): string {
  for (const key of Object.keys(process.env)) {
    if (key.startsWith("SECRET_")) {
      const value = process.env[key];
      if (value) {
        stdout = stdout.replace(value, "***");
      }
    }
  }
  return stdout;
}

export function capitalizeFirstLetter(raw: string): string {
  return raw.charAt(0).toUpperCase() + raw.slice(1);
}
