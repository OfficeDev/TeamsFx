// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import SwaggerParser from "@apidevtools/swagger-parser";
import { OpenAPIV3 } from "openapi-types";
import {
  APIInfo,
  ErrorType,
  GenerateResult,
  ParseOptions,
  ValidateResult,
  ValidationStatus,
  ListAPIResult,
  ProjectType,
} from "./interfaces";
import { SpecParserError } from "./specParserError";
import { Utils } from "./utils";
import { ConstantString } from "./constants";

/**
 * A class that parses an OpenAPI specification file and provides methods to validate, list, and generate artifacts.
 */
export class SpecParser {
  public readonly pathOrSpec: string | OpenAPIV3.Document;
  public readonly parser: SwaggerParser;
  public readonly options: Required<ParseOptions>;

  private apiMap: { [key: string]: OpenAPIV3.PathItemObject } | undefined;
  private spec: OpenAPIV3.Document | undefined;
  private unResolveSpec: OpenAPIV3.Document | undefined;
  private isSwaggerFile: boolean | undefined;

  private defaultOptions: ParseOptions = {
    allowMissingId: false,
    allowSwagger: false,
    allowAPIKeyAuth: false,
    allowMultipleParameters: false,
    allowBearerTokenAuth: false,
    allowOauth2: false,
    allowMethods: ["get", "post"],
    projectType: ProjectType.SME,
  };

  /**
   * Creates a new instance of the SpecParser class.
   * @param pathOrDoc The path to the OpenAPI specification file or the OpenAPI specification object.
   * @param options The options for parsing the OpenAPI specification file.
   */
  constructor(pathOrDoc: string | OpenAPIV3.Document, options?: ParseOptions) {
    this.pathOrSpec = pathOrDoc;
    this.parser = new SwaggerParser();
    this.options = {
      ...this.defaultOptions,
      ...(options ?? {}),
    } as Required<ParseOptions>;
  }

  /**
   * Validates the OpenAPI specification file and returns a validation result.
   *
   * @returns A validation result object that contains information about any errors or warnings in the specification file.
   */
  async validate(): Promise<ValidateResult> {
    try {
      try {
        await this.loadSpec();
        await this.parser.validate(this.spec!, {
          validate: {
            schema: false,
          },
        });
      } catch (e) {
        return {
          status: ValidationStatus.Error,
          warnings: [],
          errors: [{ type: ErrorType.SpecNotValid, content: (e as Error).toString() }],
        };
      }

      if (!this.options.allowSwagger && this.isSwaggerFile) {
        return {
          status: ValidationStatus.Error,
          warnings: [],
          errors: [
            { type: ErrorType.SwaggerNotSupported, content: ConstantString.SwaggerNotSupported },
          ],
        };
      }

      return Utils.validateSpec(this.spec!, this.parser, !!this.isSwaggerFile, this.options);
    } catch (err) {
      throw new SpecParserError((err as Error).toString(), ErrorType.ValidateFailed);
    }
  }

  async listSupportedAPIInfo(): Promise<APIInfo[]> {
    try {
      await this.loadSpec();
      const apiMap = this.getAllSupportedAPIs(this.spec!);
      const apiInfos: APIInfo[] = [];
      for (const key in apiMap) {
        const pathObjectItem = apiMap[key];
        const [method, path] = key.split(" ");
        const operationId = pathObjectItem.operationId;

        // In Browser environment, this api is by default not support api without operationId
        if (!operationId) {
          continue;
        }

        const command = Utils.parseApiInfo(pathObjectItem, this.options);

        const apiInfo: APIInfo = {
          method: method,
          path: path,
          title: command.title,
          id: operationId,
          parameters: command.parameters!,
          description: command.description!,
        };

        apiInfos.push(apiInfo);
      }

      return apiInfos;
    } catch (err) {
      throw new SpecParserError((err as Error).toString(), ErrorType.listSupportedAPIInfoFailed);
    }
  }

  /**
   * Lists all the OpenAPI operations in the specification file.
   * @returns A string array that represents the HTTP method and path of each operation, such as ['GET /pets/{petId}', 'GET /user/{userId}']
   * according to copilot plugin spec, only list get and post method without auth
   */
  // eslint-disable-next-line @typescript-eslint/require-await
  async list(): Promise<ListAPIResult[]> {
    throw new Error("Method not implemented.");
  }

  /**
   * Generate specs according to the filters.
   * @param filter An array of strings that represent the filters to apply when generating the artifacts. If filter is empty, it would process nothing.
   */
  // eslint-disable-next-line @typescript-eslint/require-await
  async getFilteredSpecs(
    filter: string[],
    signal?: AbortSignal
  ): Promise<[OpenAPIV3.Document, OpenAPIV3.Document]> {
    throw new Error("Method not implemented.");
  }

  /**
   * Generates and update artifacts from the OpenAPI specification file. Generate Adaptive Cards, update Teams app manifest, and generate a new OpenAPI specification file.
   * @param manifestPath A file path of the Teams app manifest file to update.
   * @param filter An array of strings that represent the filters to apply when generating the artifacts. If filter is empty, it would process nothing.
   * @param outputSpecPath File path of the new OpenAPI specification file to generate. If not specified or empty, no spec file will be generated.
   * @param pluginFilePath File path of the api plugin file to generate.
   */
  // eslint-disable-next-line @typescript-eslint/require-await
  async generateForCopilot(
    manifestPath: string,
    filter: string[],
    outputSpecPath: string,
    pluginFilePath: string,
    signal?: AbortSignal
  ): Promise<GenerateResult> {
    throw new Error("Method not implemented.");
  }
  /**
   * Generates and update artifacts from the OpenAPI specification file. Generate Adaptive Cards, update Teams app manifest, and generate a new OpenAPI specification file.
   * @param manifestPath A file path of the Teams app manifest file to update.
   * @param filter An array of strings that represent the filters to apply when generating the artifacts. If filter is empty, it would process nothing.
   * @param outputSpecPath File path of the new OpenAPI specification file to generate. If not specified or empty, no spec file will be generated.
   * @param adaptiveCardFolder Folder path where the Adaptive Card files will be generated. If not specified or empty, Adaptive Card files will not be generated.
   * @param isMe Boolean that indicates whether the project is an Messaging Extension. For Messaging Extension, composeExtensions will be added in Teams app manifest.
   */
  // eslint-disable-next-line @typescript-eslint/require-await
  async generate(
    manifestPath: string,
    filter: string[],
    outputSpecPath: string,
    adaptiveCardFolder?: string,
    signal?: AbortSignal
  ): Promise<GenerateResult> {
    throw new Error("Method not implemented.");
  }

  private async loadSpec(): Promise<void> {
    if (!this.spec) {
      this.unResolveSpec = (await this.parser.parse(this.pathOrSpec)) as OpenAPIV3.Document;
      if (!this.unResolveSpec.openapi && (this.unResolveSpec as any).swagger === "2.0") {
        this.isSwaggerFile = true;
      }

      const clonedUnResolveSpec = JSON.parse(JSON.stringify(this.unResolveSpec));
      this.spec = (await this.parser.dereference(clonedUnResolveSpec)) as OpenAPIV3.Document;
    }
  }

  private getAllSupportedAPIs(spec: OpenAPIV3.Document): {
    [key: string]: OpenAPIV3.OperationObject;
  } {
    if (this.apiMap !== undefined) {
      return this.apiMap;
    }
    const result = this.listSupportedAPIs(spec, this.options);
    this.apiMap = result;
    return result;
  }

  private listSupportedAPIs(
    spec: OpenAPIV3.Document,
    options: ParseOptions
  ): {
    [key: string]: OpenAPIV3.OperationObject;
  } {
    const paths = spec.paths;
    const result: { [key: string]: OpenAPIV3.OperationObject } = {};
    for (const path in paths) {
      const methods = paths[path];
      for (const method in methods) {
        const operationObject = (methods as any)[method] as OpenAPIV3.OperationObject;
        if (options.allowMethods?.includes(method) && operationObject) {
          const validateResult = Utils.isSupportedApi(method, path, spec, options);
          if (validateResult.isValid) {
            result[`${method.toUpperCase()} ${path}`] = operationObject;
          }
        }
      }
    }
    return result;
  }
}
