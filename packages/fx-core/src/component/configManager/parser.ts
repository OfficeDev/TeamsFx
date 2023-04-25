/**
 * @author yefuwang@microsoft.com
 */

import { FxError, Result, ok, err } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import { load } from "js-yaml";
import Ajv from "ajv";
import { globalVars } from "../../core/globalVars";
import { InvalidYamlSchemaError, YamlFieldMissingError, YamlFieldTypeError } from "../../error/yml";
import { IYamlParser, ProjectModel, RawProjectModel, LifecycleNames } from "./interface";
import { Lifecycle } from "./lifecycle";
import path from "path";
import { getResourceFolder } from "../../folder";
import { YAMLDiagnostics } from "./diagnostic";

const ajv = new Ajv();
ajv.addKeyword("deprecationMessage");
const schemaPath = path.join(getResourceFolder(), "yaml.schema.json");
const schema = fs.readJSONSync(schemaPath);
const validator = ajv.compile(schema);
const schemaString = fs.readFileSync(path.join(getResourceFolder(), "yaml.schema.json"), "utf8");
const yamlDiagnostic = new YAMLDiagnostics(schemaPath, schemaString);

const environmentFolderPath = "environmentFolderPath";
const writeToEnvironmentFile = "writeToEnvironmentFile";

function parseRawProjectModel(obj: Record<string, unknown>): Result<RawProjectModel, FxError> {
  const result: RawProjectModel = {};
  if (environmentFolderPath in obj) {
    if (typeof obj["environmentFolderPath"] !== "string") {
      return err(new YamlFieldTypeError("environmentFolderPath", "string"));
    }
    result.environmentFolderPath = obj[environmentFolderPath] as string;
  }
  for (const name of LifecycleNames) {
    if (name in obj) {
      const value = obj[name];
      if (!Array.isArray(value)) {
        return err(new YamlFieldTypeError(name, "array"));
      }
      for (const elem of value) {
        if (!("uses" in elem)) {
          return err(new YamlFieldMissingError(`${name}.uses`));
        }
        if (!(typeof elem["uses"] === "string")) {
          return err(new YamlFieldTypeError(`${name}.uses`, "string"));
        }
        if (!("with" in elem)) {
          return err(new YamlFieldMissingError(`${name}.with`));
        }
        if (!(typeof elem["with"] === "object")) {
          return err(new YamlFieldTypeError(`${name}.with`, "object"));
        }
        if (elem["env"]) {
          if (typeof elem["env"] !== "object" || Array.isArray(elem["env"])) {
            return err(new YamlFieldTypeError(`${name}.env`, "object"));
          }
          for (const envVar in elem["env"]) {
            if (typeof elem["env"][envVar] !== "string") {
              return err(new YamlFieldTypeError(`${name}.env.${envVar}`, "string"));
            }
          }
        }
        if (elem[writeToEnvironmentFile]) {
          if (
            typeof elem[writeToEnvironmentFile] !== "object" ||
            Array.isArray(elem[writeToEnvironmentFile])
          ) {
            return err(new YamlFieldTypeError(`${name}.writeToEnvironmentFile`, "object"));
          }
          for (const envVar in elem[writeToEnvironmentFile]) {
            if (typeof elem[writeToEnvironmentFile][envVar] !== "string") {
              return err(
                new YamlFieldTypeError(`${name}.writeToEnvironmentFile.${envVar}`, "string")
              );
            }
          }
        }
      }
      result[name] = value;
    }
  }

  return ok(result);
}

export class YamlParser implements IYamlParser {
  async parse(path: string, validateSchema?: boolean): Promise<Result<ProjectModel, FxError>> {
    const raw = await this.parseRaw(path, validateSchema);
    if (raw.isErr()) {
      return err(raw.error);
    }
    const result: ProjectModel = {};
    for (const name of LifecycleNames) {
      if (name in raw.value) {
        const definitions = raw.value[name];
        if (definitions) {
          result[name] = new Lifecycle(name, definitions);
        }
      }
    }

    if (raw.value.environmentFolderPath) {
      result.environmentFolderPath = raw.value.environmentFolderPath;
    }

    return ok(result);
  }

  private async parseRaw(
    path: string,
    validateSchema?: boolean
  ): Promise<Result<RawProjectModel, FxError>> {
    let diagnostic: string | undefined = undefined;
    try {
      globalVars.ymlFilePath = path;
      const str = await fs.readFile(path, "utf8");
      diagnostic = await yamlDiagnostic.doValidation(path, str);
      const content = load(str);
      // note: typeof null === "object" typeof undefined === "undefined" in js
      if (typeof content !== "object" || Array.isArray(content) || content === null) {
        return err(new InvalidYamlSchemaError(path, diagnostic));
      }
      const value = content as unknown as Record<string, unknown>;
      if (validateSchema) {
        const valid = validator(value);
        if (!valid) {
          return err(new InvalidYamlSchemaError(path, diagnostic));
        }
      }
      return parseRawProjectModel(value);
    } catch (error) {
      return err(new InvalidYamlSchemaError(path, diagnostic));
    }
  }
}

export const yamlParser = new YamlParser();
