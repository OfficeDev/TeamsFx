import { err, FxError, ok, Result, UserError } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import { cloneDeep, merge } from "lodash";
import { settingsUtil } from "./settingsUtil";
import { LocalCrypto } from "../../core/crypto";
import { getDefaultString, getLocalizedString } from "../../common/localizeUtils";
import { pathUtils } from "./pathUtils";
import { TOOLS } from "../../core/globalVars";
import * as path from "path";
import { EOL } from "os";

export type DotenvOutput = {
  [k: string]: string;
};

export class EnvUtil {
  /**
   * read .env file and set to process.env (if loadToProcessEnv = true)
   * if silent = true, no error will return if .env file is not available, this function returns ok({ TEAMSFX_ENV: env })
   * if silent = false, this function will return error if .env file is not available.
   * @param projectPath
   * @param env
   * @param loadToProcessEnv
   * @param silent
   * @returns
   */
  async readEnv(
    projectPath: string,
    env: string,
    loadToProcessEnv = true,
    silent = true
  ): Promise<Result<DotenvOutput, FxError>> {
    // read
    const dotEnvFilePathRes = await pathUtils.getEnvFilePath(projectPath, env);
    if (dotEnvFilePathRes.isErr()) return err(dotEnvFilePathRes.error);
    const dotEnvFilePath = dotEnvFilePathRes.value;
    if (!dotEnvFilePath || !(await fs.pathExists(dotEnvFilePath))) {
      if (silent) {
        // .env file does not exist, just ignore
        process.env.TEAMSFX_ENV = env;
        return ok({ TEAMSFX_ENV: env });
      } else {
        return err(
          new UserError({
            source: "core",
            name: "DotEnvFileNotExistError",
            displayMessage: getLocalizedString("error.DotEnvFileNotExistError", env, env),
            message: getDefaultString("error.DotEnvFileNotExistError", env, env),
          })
        );
      }
    }
    // deserialize
    const parseResult = dotenvUtil.deserialize(
      await fs.readFile(dotEnvFilePath, { encoding: "utf8" })
    );

    // decrypt
    const settingsRes = await settingsUtil.readSettings(projectPath);
    if (settingsRes.isErr()) {
      return err(settingsRes.error);
    }
    const projectId = settingsRes.value.trackingId;
    const cryptoProvider = new LocalCrypto(projectId);
    for (const key of Object.keys(parseResult.obj)) {
      if (key.startsWith("SECRET_")) {
        const raw = parseResult.obj[key];
        if (raw.startsWith("crypto_")) {
          const decryptRes = await cryptoProvider.decrypt(raw);
          if (decryptRes.isErr()) return err(decryptRes.error);
          parseResult.obj[key] = decryptRes.value;
        }
      }
    }
    parseResult.obj.TEAMSFX_ENV = env;
    if (loadToProcessEnv) {
      merge(process.env, parseResult.obj);
    }
    return ok(parseResult.obj);
  }

  /**
   * write env variables into .env file,
   * if .env file does not exist, this function will create a default one
   * if .env fila path is not available, the default path is `./env/.env.{env}`
   * @param projectPath
   * @param env
   * @param envs
   * @returns
   */
  async writeEnv(
    projectPath: string,
    env: string,
    envs: DotenvOutput
  ): Promise<Result<undefined, FxError>> {
    envs.TEAMSFX_ENV = env;
    //encrypt
    const settingsRes = await settingsUtil.readSettings(projectPath);
    if (settingsRes.isErr()) {
      return err(settingsRes.error);
    }
    const projectId = settingsRes.value.trackingId;
    const cryptoProvider = new LocalCrypto(projectId);
    for (const key of Object.keys(envs)) {
      let value = envs[key];
      if (value && key.startsWith("SECRET_")) {
        const res = await cryptoProvider.encrypt(value);
        if (res.isErr()) return err(res.error);
        value = res.value;
        envs[key] = value;
      }
    }

    //replace existing, if env file not exist, create a default one
    const dotEnvFilePathRes = await pathUtils.getEnvFilePath(projectPath, env);
    if (dotEnvFilePathRes.isErr()) return err(dotEnvFilePathRes.error);
    const dotEnvFilePath =
      dotEnvFilePathRes.value || path.resolve(projectPath, "env", `.env.${env ? env : "dev"}`);
    const envFileExists = await fs.pathExists(dotEnvFilePath);
    const parsedDotenv = envFileExists
      ? dotenvUtil.deserialize(await fs.readFile(dotEnvFilePath))
      : { obj: {} };
    merge(parsedDotenv.obj, envs);

    //serialize
    const content = dotenvUtil.serialize(parsedDotenv);

    //persist
    if (!envFileExists) await fs.ensureFile(dotEnvFilePath);
    await fs.writeFile(dotEnvFilePath, content, { encoding: "utf8" });
    if (!envFileExists) {
      TOOLS.logProvider.info("  Created environment file at " + dotEnvFilePath + EOL + EOL);
    }
    return ok(undefined);
  }
  async listEnv(projectPath: string): Promise<Result<string[], FxError>> {
    const folderRes = await pathUtils.getEnvFolderPath(projectPath);
    if (folderRes.isErr()) return err(folderRes.error);
    const envFolderPath = folderRes.value;
    if (!envFolderPath) return ok([]);
    const list = await fs.readdir(envFolderPath);
    const envs = list
      .filter((fileName) => fileName.startsWith(".env."))
      .map((fileName) => fileName.substring(5));
    return ok(envs);
  }
  object2map(obj: DotenvOutput): Map<string, string> {
    const map = new Map<string, string>();
    for (const key of Object.keys(obj)) {
      map.set(key, obj[key]);
    }
    return map;
  }
  map2object(map: Map<string, string>): DotenvOutput {
    const obj: DotenvOutput = {};
    for (const key of map.keys()) {
      obj[key] = map.get(key) || "";
    }
    return obj;
  }
}

export const envUtil = new EnvUtil();

const NEW_LINE_SPLITTER = /\r?\n/;
const LINE_RE =
  /(?:^|^)\s*(?:export\s+)?([\w.-]+)(?:\s*=\s*?|:\s+?)(\s*'(?:\\'|[^'])*'|\s*"(?:\\"|[^"])*"|\s*`(?:\\`|[^`])*`|[^#\r\n]+)?\s*(?:#.*)?(?:$|$)/gm;
type DotenvParsedLine =
  | string
  | { key: string; value: string; comment?: string; quote?: '"' | "'" };
export interface DotenvParseResult {
  lines?: DotenvParsedLine[];
  obj: DotenvOutput;
}

export class DotenvUtil {
  deserialize(src: string | Buffer): DotenvParseResult {
    const lines: DotenvParsedLine[] = [];
    const obj: DotenvOutput = {};
    const stringLines = src.toString().replace(/\r\n?/gm, "\n").split(NEW_LINE_SPLITTER);
    for (const line of stringLines) {
      const match = LINE_RE.exec(line);
      if (match) {
        let inlineComment;
        // extract key
        const key = match[1];
        // extract value
        let value = match[2] || "";
        // try to find comments
        const valueIndex = match[0].indexOf(value);
        if (valueIndex >= 0) {
          const remaining = match[0].substring(valueIndex + value.length).trim();
          if (remaining.startsWith("#")) {
            inlineComment = remaining;
          }
        }
        // trim
        value = value.trim();
        //quote
        const firstChar = value[0];
        value = value.replace(/^(['"`])([\s\S]*)\1$/gm, "$2");
        //de-escape
        if (firstChar === '"') {
          value = value.replace(/\\n/g, "\n");
          value = value.replace(/\\r/g, "\r");
        }
        //output
        obj[key] = value;
        const parsedLine: DotenvParsedLine = { key: key, value: value };
        if (inlineComment) parsedLine.comment = inlineComment;
        if (firstChar === '"' || firstChar === "'") parsedLine.quote = firstChar as '"' | "'";
        lines.push(parsedLine);
      } else {
        lines.push(line);
      }
    }
    return { lines: lines, obj: obj };
  }
  serialize(parsed: DotenvParseResult): string {
    const array: string[] = [];
    const obj = cloneDeep(parsed.obj);
    //append lines
    if (parsed.lines) {
      parsed.lines.forEach((line) => {
        if (typeof line === "string") {
          // keep comment line or empty line
          array.push(line);
        } else {
          if (obj[line.key] !== undefined) {
            // use kv in obj
            line.value = obj[line.key];
            delete obj[line.key];
          }
          if (line.value.includes("#") && !line.quote) {
            // if value contains '#', need add quote
            line.quote = '"';
          }
          let value = line.value;
          if (line.quote) {
            value = line.quote === "'" ? value.replace(/'/g, "\\'") : value.replace(/"/g, '\\"');
            value = `${line.quote}${value}${line.quote}`;
          }
          array.push(`${line.key}=${value}${line.comment ? line.comment : ""}`);
        }
      });
    }
    //append additional kvs in object
    for (const key of Object.keys(obj)) {
      let value = parsed.obj[key];
      if (value.includes("#")) value = `"${value}"`; // if value contains '#', need add quote
      array.push(`${key}=${value}`);
    }
    return array.join("\n").trim();
  }
}

export const dotenvUtil = new DotenvUtil();
