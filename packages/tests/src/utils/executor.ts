// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ProgrammingLanguage } from "@microsoft/teamsfx-core";
import { execAsync, editDotEnvFile } from "./commonUtils";
import {
  TemplateProjectFolder,
  Capability,
  LocalDebugError,
} from "./constants";
import path from "path";
import fs from "fs-extra";
import * as os from "os";
import { spawn, ChildProcessWithoutNullStreams } from "child_process";
import { expect } from "chai";

export class Executor {
  static async execute(
    command: string,
    cwd: string,
    processEnv?: NodeJS.ProcessEnv,
    timeout?: number
  ) {
    try {
      const result = await execAsync(command, {
        cwd,
        env: processEnv ?? process.env,
        timeout: timeout ?? 0,
      });
      if (result.stderr) {
        /// the command exit with 0
        console.log(
          `[Pending] "${command}" in ${cwd} with some stderr: ${result.stderr}`
        );
        return { ...result, success: false };
      } else {
        console.log(`[Success] "${command}" in ${cwd}.`);
        return { ...result, success: true };
      }
    } catch (e: any) {
      if (e.killed && e.signal == "SIGTERM") {
        console.error(`[Failed] "${command}" in ${cwd}. Timeout and killed.`);
      } else {
        console.error(
          `[Failed] "${command}" in ${cwd} with error: ${e.message}`
        );
      }
      return { stdout: "", stderr: e.message as string, success: false };
    }
  }

  static login() {
    const command = `az login --service-principal -u ${process.env.AZURE_CLIENT_ID} -p ${process.env.AZURE_CLIENT_SECRET} -t ${process.env.AZURE_TENANT_ID}`;
    return this.execute(command, process.cwd());
  }

  static concatProcessEnv(
    processEnv: NodeJS.ProcessEnv,
    env: Record<string, string>
  ) {
    return Object.assign({}, processEnv, env);
  }

  static async createProject(
    workspace: string,
    appName: string,
    capability: Capability,
    language: ProgrammingLanguage,
    customized: Record<string, string> = {}
  ) {
    const command =
      `teamsapp new --interactive false --app-name ${appName} --capability ${capability} --programming-language ${language} ` +
      Object.entries(customized)
        .map(([key, value]) => "--" + key + " " + value)
        .join(" ");
    return this.execute(command, workspace);
  }

  static async addEnv(workspace: string, newEnv: string, env = "dev") {
    const command = `teamsapp env add ${newEnv} --env ${env}`;
    return this.execute(command, workspace);
  }

  static async addSPFxWebPart(
    workspace: string,
    spfxFolder: string,
    webpartName: string,
    manifestPath: string,
    localManifestPath: string
  ) {
    const command =
      `teamsapp add spfx-web-part --spfx-webpart-name ${webpartName}` +
      ` --spfx-folder ${spfxFolder} --teams-manifest-file ${manifestPath}` +
      ` --local-teams-manifest-file ${localManifestPath} --interactive false `;
    return this.execute(command, workspace);
  }

  static async upgrade(workspace: string, isV3 = true) {
    const prefix = isV3 ? "teamsapp" : "teamsfx";
    const command = `${prefix} upgrade --force`;
    return this.execute(command, workspace);
  }

  static async executeCmd(
    workspace: string,
    cmd: string,
    env = "dev",
    processEnv?: NodeJS.ProcessEnv,
    npx = false,
    isV3 = true
  ) {
    const npxCommand = npx ? "npx " : "";
    const cliPrefix = isV3 ? "teamsapp" : "teamsfx";
    const command = `${npxCommand} ${cliPrefix} ${cmd} --env ${env}`;
    return this.execute(command, workspace, processEnv);
  }

  static async provision(workspace: string, env = "dev", isV3 = true) {
    return this.executeCmd(workspace, "provision", env, undefined, false, isV3);
  }

  static async provisionWithCustomizedProcessEnv(
    workspace: string,
    processEnv: NodeJS.ProcessEnv,
    env = "dev",
    npx = false,
    isV3 = true
  ) {
    return this.executeCmd(workspace, "provision", env, processEnv, npx, isV3);
  }

  static async validate(
    workspace: string,
    env = "dev",
    manifestFolderName = "appPackage"
  ) {
    return this.executeCmd(
      workspace,
      `validate --manifest-file ./${manifestFolderName}/manifest.json`,
      env
    );
  }

  static async validateWithCustomizedProcessEnv(
    workspace: string,
    processEnv: NodeJS.ProcessEnv,
    env = "dev",
    npx = false,
    isV3 = true
  ) {
    return this.executeCmd(workspace, "deploy", env, processEnv, npx, isV3);
  }

  static async deploy(workspace: string, env = "dev") {
    return this.executeCmd(workspace, "deploy", env);
  }

  static async deployWithCustomizedProcessEnv(
    workspace: string,
    processEnv: NodeJS.ProcessEnv,
    env = "dev",
    npx = false,
    isV3 = true
  ) {
    return this.executeCmd(workspace, "deploy", env, processEnv, npx, isV3);
  }

  static async publish(workspace: string, env = "dev") {
    return this.executeCmd(workspace, "publish", env);
  }

  static async publishWithCustomizedProcessEnv(
    workspace: string,
    processEnv: NodeJS.ProcessEnv,
    env = "dev",
    npx = false,
    isV3 = true
  ) {
    return this.executeCmd(workspace, "publish", env, processEnv, npx, isV3);
  }

  static async preview(workspace: string, env = "dev") {
    return this.executeCmd(workspace, "preview", env);
  }

  static debugProject(
    projectPath: string,
    env: "local" | "dev" = "local",
    v3 = true,
    processEnv: NodeJS.ProcessEnv = process.env,
    onData?: (data: string) => void,
    onError?: (data: string) => void
  ) {
    console.log(`[start] ${env} debug ... `);
    const childProcess = spawn(
      os.type() === "Windows_NT"
        ? v3
          ? "teamsapp.cmd"
          : "teamsfx.cmd"
        : v3
        ? "teamsapp"
        : "teamsfx",
      ["preview", v3 ? "--env" : "", v3 ? `${env}` : `--${env}`],
      {
        cwd: projectPath,
        env: processEnv ? processEnv : process.env,
      }
    );
    childProcess.stdout.on("data", (data) => {
      const dataString = data.toString();
      if (onData) {
        onData(dataString);
      }
    });
    childProcess.stderr.on("data", (data) => {
      const dataString = data.toString();
      if (onError) {
        onError(dataString);
      }
    });
    return childProcess;
  }

  static async previewWithCustomizedProcessEnv(
    workspace: string,
    processEnv: NodeJS.ProcessEnv,
    env = "dev",
    npx = false,
    isV3 = true
  ) {
    return this.executeCmd(workspace, "preview", env, processEnv, npx, isV3);
  }

  static async installCLI(workspace: string, version: string, global: boolean) {
    const packageName = version.startsWith("3.")
      ? "@microsoft/teamsapp-cli"
      : "@microsoft/teamsfx-cli";
    if (global) {
      const command = `npm install -g ${packageName}@${version}`;
      return this.execute(command, workspace);
    } else {
      const command = `npm install ${packageName}@${version}`;
      return this.execute(command, workspace);
    }
  }

  static async createTemplateProject(
    appName: string,
    testFolder: string,
    template: TemplateProjectFolder,
    processEnv?: NodeJS.ProcessEnv
  ) {
    const command = `teamsapp new sample ${template} --interactive false `;
    const timeout = 100000;
    try {
      await this.execute(command, testFolder, processEnv, timeout);

      //  change original template name to appName
      await this.execute(
        `mv ./${template} ./${appName}`,
        testFolder,
        processEnv ? processEnv : process.env,
        timeout
      );

      const localEnvPath = path.resolve(
        testFolder,
        appName,
        "env",
        ".env.local"
      );
      const remoteEnvPath = path.resolve(
        testFolder,
        appName,
        "env",
        ".env.dev"
      );
      editDotEnvFile(localEnvPath, "TEAMS_APP_NAME", appName);
      editDotEnvFile(remoteEnvPath, "TEAMS_APP_NAME", appName);

      const message = `scaffold project to ${path.resolve(
        testFolder,
        appName
      )} with template ${template}`;
      console.log(message);
    } catch (e: any) {
      console.log(
        `Run \`${command}\` failed with error msg: ${JSON.stringify(e)}.`
      );
      if (e.killed && e.signal == "SIGTERM") {
        console.log(`Command ${command} killed due to timeout ${timeout}`);
      }
    }
  }

  static async openTemplateProject(
    appName: string,
    testFolder: string,
    template: TemplateProjectFolder,
    processEnv?: NodeJS.ProcessEnv,
    subFolder?: string
  ) {
    const timeout = 100000;
    let oldPath = "";
    if (subFolder) {
      oldPath = path.resolve(
        __dirname,
        "..",
        "e2e/resource",
        subFolder,
        template
      );
    } else {
      oldPath = path.resolve(__dirname, "..", "e2e/resource", template);
    }
    const newPath = path.resolve(testFolder, appName);
    try {
      await this.execute(
        `mv ${oldPath} ${newPath}`,
        testFolder,
        processEnv,
        timeout
      );
    } catch (error) {
      console.log(error);
      throw new Error(`Failed to open project: ${newPath}`);
    }
    const localEnvPath = path.resolve(testFolder, appName, "env", ".env.local");
    const remoteEnvPath = path.resolve(testFolder, appName, "env", ".env.dev");
    editDotEnvFile(localEnvPath, "TEAMS_APP_NAME", appName);
    editDotEnvFile(remoteEnvPath, "TEAMS_APP_NAME", appName);
    console.log(`successfully open project: ${newPath}`);
  }

  static async package(
    workspace: string,
    env = "dev",
    manifestFolderName = "appPackage"
  ) {
    return this.executeCmd(
      workspace,
      `package --manifest-file ./${manifestFolderName}/manifest.json`,
      env
    );
  }

  static startDevtunnel(
    onData?: (data: string) => void,
    onError?: (data: string) => void
  ) {
    const child = spawn(
      os.type() === "Windows_NT"
        ? "devtunnel"
        : `${os.homedir()}/bin/devtunnel`,
      ["host", "-p", "3978", "--allow-anonymous"],
      {
        env: process.env,
      }
    );
    child.stdout.on("data", (data) => {
      const dataString = data.toString();
      if (onData) {
        onData(dataString);
      }
    });
    child.stderr.on("data", (data) => {
      const dataString = data.toString();
      if (onError) {
        onError(dataString);
      }
    });
    return child;
  }

  static deleteTunnel(
    tunnelName: string,
    onData?: (data: string) => void,
    onError?: (data: string) => void
  ) {
    const child = spawn(
      os.type() === "Windows_NT"
        ? "devtunnel"
        : `${os.homedir()}/bin/devtunnel`,
      ["delete", tunnelName, "-f"],
      {
        env: process.env,
      }
    );
    console.log("delete tunnel: ", tunnelName);
    child.stdout.on("data", (data) => {
      const dataString = data.toString();
      if (onData) {
        onData(dataString);
      }
    });
    child.stderr.on("data", (data) => {
      const dataString = data.toString();
      if (onError) {
        onError(dataString);
      }
    });
    return child;
  }

  static deleteAllTunnel(
    onData?: (data: string) => void,
    onError?: (data: string) => void
  ) {
    const child = spawn(
      os.type() === "Windows_NT"
        ? "devtunnel"
        : `${os.homedir()}/bin/devtunnel`,
      ["delete-all", "-f"],
      {
        env: process.env,
      }
    );
    child.stdout.on("data", (data) => {
      const dataString = data.toString();
      if (onData) {
        onData(dataString);
      }
    });
    child.stderr.on("data", (data) => {
      const dataString = data.toString();
      if (onError) {
        onError(dataString);
      }
    });
  }

  static spawnCommand(
    projectPath: string,
    command: string,
    args: string[],
    onData?: (data: string) => void,
    onError?: (data: string) => void
  ) {
    const childProcess = spawn(
      os.type() === "Windows_NT" ? command + ".cmd" : command,
      args,
      {
        cwd: projectPath,
        env: process.env,
      }
    );
    childProcess.stdout.on("data", (data) => {
      const dataString = data.toString();
      if (onData) {
        onData(dataString);
      }
    });
    childProcess.stderr.on("data", (data) => {
      const dataString = data.toString();
      if (onError) {
        onError(dataString);
      }
    });
    return childProcess;
  }

  static debugBotFunctionPreparation(projectPath: string) {
    let envFile = "";
    let tunnelName = "";
    let envContent = "";
    try {
      envFile = path.resolve(projectPath, "env", ".env.local");
      envContent = fs.readFileSync(envFile, "utf-8");
    } catch (error) {
      console.log("read file error", error);
      console.log("create .env.local file");
      fs.writeFileSync(envFile, ''); 
    }
    const domainRegex = /Connect via browser: https:\/\/(\S+)/;
    const endpointRegex = /Connect via browser: (\S+)/;
    const tunnelNameRegex = /Ready to accept connections for tunnel: (\S+)/;
    const devtunnelProcess = Executor.startDevtunnel(
      (data) => {
        if (data) {
          // start devtunnel
          const domainFound = data.match(domainRegex);
          const endpointFound = data.match(endpointRegex);
          const tunnelNameFound = data.match(tunnelNameRegex);
          if (domainFound && endpointFound) {
            if (domainFound[1] && endpointFound[1]) {
              const domain = domainFound[1];
              const endpoint = endpointFound[1];
              try {
                console.log(endpoint);
                envContent += `\nBOT_ENDPOINT=${endpoint}`;
                envContent += `\nBOT_DOMAIN=${domain}`;
                envContent += `\nBOT_FUNCTION_ENDPOINT=${endpoint}`;
                envContent += `\nPROVISIONOUTPUT__BOTOUTPUT__SITEENDPOINT=${endpoint}`;
                envContent += `\nPROVISIONOUTPUT__BOTOUTPUT__VALIDDOMAIN=${domain}`;
                fs.writeFileSync(envFile, envContent);
                console.log(envContent);
              } catch (error) {
                console.log(error);
              }
            }
          }
          if (tunnelNameFound) {
            if (tunnelNameFound[1]) {
              tunnelName = tunnelNameFound[1];
              console.log(tunnelName);
            }
          }
        }
      },
      (error) => {
        console.log(error);
      }
    );
    return { devtunnelProcess, tunnelName };
  }

  static async cliPreview(projectPath: string, includeBot: boolean) {
    console.log("======= debug with cli ========");
    console.log("botFlag: ", includeBot);
    let tunnelName = "";
    let devtunnelProcess = null;
    if (includeBot) {
      const tunnel = Executor.debugBotFunctionPreparation(projectPath);
      tunnelName = tunnel.tunnelName;
      devtunnelProcess = tunnel.devtunnelProcess;
      await new Promise((resolve) => setTimeout(resolve, 60 * 1000));
    }
    const debugProcess = Executor.debugProject(
      projectPath,
      "local",
      true,
      process.env,
      (data) => {
        if (data) {
          console.log(data);
        }
      },
      (error) => {
        const errorMsg = error.toString();
        if (
          // skip warning messages
          errorMsg.includes(LocalDebugError.WarningError)
        ) {
          console.log("[skip error] ", error);
        } else {
          expect.fail(errorMsg);
        }
      }
    );
    await new Promise((resolve) => setTimeout(resolve, 2 * 60 * 1000));
    return {
      tunnelName,
      devtunnelProcess,
      debugProcess,
    };
  }

  static async killPort(
    port: number
  ): Promise<{ stdout: string; stderr: string }> {
    // windows
    if (process.platform === "win32") {
      const command = `for /f "tokens=5" %a in ('netstat -ano ^| find ":${port}"') do taskkill /PID %a /F`;
      console.log("run command: ", command);
      const result = await execAsync(command);
      return result;
    } else {
      const command = `kill -9 $(lsof -t -i:${port})`;
      console.log("run command: ", command);
      const result = await execAsync(command);
      return result;
    }
  }

  static async closeProcess(
    childProcess: ChildProcessWithoutNullStreams | null
  ) {
    if (childProcess) {
      try {
        if (os.type() === "Windows_NT") {
          console.log(`taskkill /F /PID "${childProcess.pid}"`);
          await execAsync(`taskkill /F /PID "${childProcess.pid}"`);
          childProcess.kill("SIGKILL");
        } else {
          console.log("kill process", childProcess.spawnargs.join(" "));
          childProcess.kill("SIGKILL");
        }
      } catch (error) {
        console.log(error);
      }
    }
  }
}
