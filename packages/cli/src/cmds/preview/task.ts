// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import cp from "child_process";
import { err, FxError, LogLevel, ok, Result } from "@microsoft/teamsfx-api";
import treeKill from "tree-kill";
import { ServiceLogWriter } from "./serviceLogWriter";
import { CLILogProvider } from "./../../commonlib/log";

interface TaskOptions {
  shell: boolean | string;
  cwd?: string;
  env?: NodeJS.ProcessEnv;
}

export interface TaskResult {
  command: string;
  args?: string[];
  options?: TaskOptions;
  success: boolean;
  stdout: string[];
  stderr: string[];
  exitCode: number | null;
}

export class Task {
  private taskTitle: string;
  private background: boolean;
  private command: string;
  private args?: string[];
  private options?: TaskOptions;

  private resolved = false;
  private task: cp.ChildProcess | undefined;

  constructor(
    taskTitle: string,
    background: boolean,
    command: string,
    args?: string[],
    options?: TaskOptions
  ) {
    this.taskTitle = taskTitle;
    this.background = background;
    this.command = command;
    this.args = args;
    this.options = options;
  }

  /**
   * wait for the task to end
   */
  public async wait(
    startCallback: (taskTitle: string, background: boolean) => Promise<void>,
    stopCallback: (
      taskTitle: string,
      background: boolean,
      result: TaskResult
    ) => Promise<FxError | null>
  ): Promise<Result<TaskResult, FxError>> {
    await startCallback(this.taskTitle, this.background);
    this.task = cp.spawn(this.command, this.args, this.options);
    const stdout: string[] = [];
    const stderr: string[] = [];
    return new Promise((resolve) => {
      this.task?.stdout?.on("data", (data) => {
        stdout.push(data.toString());
      });
      this.task?.stderr?.on("data", (data) => {
        stderr.push(data.toString());
      });
      // eslint-disable-next-line @typescript-eslint/no-misused-promises
      this.task?.on("exit", async (code) => {
        const result: TaskResult = {
          command: this.command,
          options: this.options,
          success: code === 0,
          stdout: stdout,
          stderr: stderr,
          exitCode: code,
        };
        const error = await stopCallback(this.taskTitle, this.background, result);
        if (error) {
          resolve(err(error));
        } else {
          resolve(ok(result));
        }
      });
    });
  }

  /**
   * wait until stdout/stderr of the task matches the pattern or the task ends
   */
  public async waitFor(
    pattern: RegExp,
    startCallback: (
      taskTitle: string,
      background: boolean,
      serviceLogWriter?: ServiceLogWriter
    ) => Promise<void>,
    stopCallback: (
      taskTitle: string,
      background: boolean,
      result: TaskResult
    ) => Promise<FxError | null>,
    timeout?: number,
    serviceLogWriter?: ServiceLogWriter,
    logProvider?: CLILogProvider
  ): Promise<Result<TaskResult, FxError>> {
    await serviceLogWriter?.write(
      this.taskTitle,
      `${this.command} ${this.args ? this.args?.join(" ") : ""}\n`
    );
    await startCallback(this.taskTitle, this.background, serviceLogWriter);
    this.task = cp.spawn(this.command, this.args, this.options);
    const stdout: string[] = [];
    const stderr: string[] = [];
    return new Promise((resolve) => {
      if (timeout !== undefined) {
        // eslint-disable-next-line @typescript-eslint/no-misused-promises
        setTimeout(async () => {
          if (!this.resolved) {
            this.resolved = true;
            const result: TaskResult = {
              command: this.command,
              options: this.options,
              success: false,
              stdout: stdout,
              stderr: stderr,
              exitCode: null,
            };
            const error = await stopCallback(this.taskTitle, this.background, result);
            if (error) {
              resolve(err(error));
            } else {
              resolve(ok(result));
            }
          }
        }, timeout);
      }

      // eslint-disable-next-line @typescript-eslint/no-misused-promises
      this.task?.stdout?.on("data", async (data) => {
        const dataStr = data.toString();
        await serviceLogWriter?.write(this.taskTitle, dataStr);
        if (logProvider) {
          logProvider.necessaryLog(LogLevel.Info, dataStr.trim(), true);
        }
        stdout.push(dataStr);
        if (!this.resolved) {
          const match = pattern.test(dataStr);
          if (match) {
            this.resolved = true;
            const result: TaskResult = {
              command: this.command,
              options: this.options,
              success: true,
              stdout: stdout,
              stderr: stderr,
              exitCode: null,
            };
            const error = await stopCallback(this.taskTitle, this.background, result);
            if (error) {
              resolve(err(error));
            } else {
              resolve(ok(result));
            }
          }
        }
      });
      // eslint-disable-next-line @typescript-eslint/no-misused-promises
      this.task?.stderr?.on("data", async (data) => {
        const dataStr = data.toString();
        await serviceLogWriter?.write(this.taskTitle, dataStr);
        if (logProvider) {
          logProvider.necessaryLog(LogLevel.Info, dataStr.trim(), true);
        }
        stderr.push(dataStr);
        if (!this.resolved) {
          const match = pattern.test(dataStr);
          if (match) {
            this.resolved = true;
            const result: TaskResult = {
              command: this.command,
              options: this.options,
              success: false,
              stdout: stdout,
              stderr: stderr,
              exitCode: null,
            };
            const error = await stopCallback(this.taskTitle, this.background, result);
            if (error) {
              resolve(err(error));
            } else {
              resolve(ok(result));
            }
          }
        }
      });

      // eslint-disable-next-line @typescript-eslint/no-misused-promises
      this.task?.on("exit", async (code) => {
        if (!this.resolved) {
          this.resolved = true;
          const result: TaskResult = {
            command: this.command,
            options: this.options,
            success: false,
            stdout: stdout,
            stderr: stderr,
            exitCode: code,
          };
          const error = await stopCallback(this.taskTitle, this.background, result);
          if (error) {
            resolve(err(error));
          } else {
            resolve(ok(result));
          }
        }
      });
    });
  }

  public async terminate(): Promise<void> {
    return new Promise((resolve) => {
      if (this.task?.exitCode) {
        resolve();
      }
      const pid = this.task?.pid;
      if (pid === undefined) {
        resolve();
      } else {
        treeKill(pid, (error) => {
          if (error) {
            // ignore any error
            resolve();
          } else {
            resolve();
          }
        });
      }
    });
  }
}
