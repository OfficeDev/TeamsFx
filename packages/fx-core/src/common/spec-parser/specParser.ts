// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import * as util from "util";
import SwaggerParser from "@apidevtools/swagger-parser";
import { OpenAPIV3 } from "openapi-types";
import { SpecParserError } from "./specParserError";
import {
  ErrorResult,
  ErrorType,
  ValidateResult,
  ValidationStatus,
  WarningResult,
  WarningType,
} from "./interfaces";
import { ConstantString } from "./constants";

/**
 * A class that parses an OpenAPI specification file and provides methods to validate, list, and generate artifacts.
 */
export class SpecParser {
  private specPath: string;
  private apiMap: { [key: string]: OpenAPIV3.PathItemObject } | undefined;
  private spec: OpenAPIV3.Document | undefined;

  /**
   * Creates a new instance of the SpecParser class.
   * @param path The URL or file path of the OpenAPI specification file. The OpenAPI specification file must have a version of 3.0 or higher.
   */
  constructor(path: string) {
    this.specPath = path;
  }

  /**
   * Validates the OpenAPI specification file and returns a validation result.
   *
   * @returns A validation result object that contains information about any errors or warnings in the specification file.
   */
  async validate(): Promise<ValidateResult> {
    const errors: ErrorResult[] = [];
    const warnings: WarningResult[] = [];
    const parser = new SwaggerParser();
    try {
      if (!this.spec) {
        this.spec = (await parser.validate(this.specPath)) as OpenAPIV3.Document;
      }
    } catch (e) {
      // Spec not valid
      errors.push({ type: ErrorType.SpecNotValid, content: (e as Error).toString() });
      return {
        status: ValidationStatus.Error,
        warnings,
        errors,
      };
    }

    // Spec version not supported
    if (!this.spec.openapi || this.spec.openapi < "3.0.0") {
      errors.push({
        type: ErrorType.VersionNotSupported,
        content: ConstantString.SpecVersionNotSupported,
      });
      return {
        status: ValidationStatus.Error,
        warnings,
        errors,
      };
    }

    // Server information invalid
    if (!this.spec.servers || this.spec.servers.length === 0) {
      errors.push({
        type: ErrorType.NoServerInformation,
        content: ConstantString.NoServerInformation,
      });
    } else if (this.spec.servers.length > 1) {
      errors.push({
        type: ErrorType.MultipleServerInformation,
        content: ConstantString.MultipleServerInformation,
      });
    }

    // Remote reference not supported
    const refPaths = parser.$refs.paths();

    // refPaths [0] is the current spec file path
    if (refPaths.length > 1) {
      console.log("refPaths", refPaths);
      errors.push({
        type: ErrorType.RemoteRefNotSupported,
        content: util.format(ConstantString.RemoteRefNotSupported, refPaths.join(", ")),
      });
    }

    // No supported API
    const apiMap = await this.getAllSupportedApi(this.spec);
    if (Object.keys(apiMap).length === 0) {
      errors.push({
        type: ErrorType.NoSupportedApi,
        content: ConstantString.NoSupportedApi,
      });
    }

    // OperationId missing
    const apisMissingOperationId: string[] = [];
    for (const key in apiMap) {
      const pathObjectItem = apiMap[key];
      if (!pathObjectItem.operationId) {
        apisMissingOperationId.push(key);
      }
    }

    if (apisMissingOperationId.length > 0) {
      warnings.push({
        type: WarningType.OperationIdMissing,
        content: util.format(ConstantString.MissingOperationId, apisMissingOperationId.join(", ")),
      });
    }

    let status = ValidationStatus.Valid;
    if (warnings.length > 0 && errors.length === 0) {
      status = ValidationStatus.Warning;
    } else if (errors.length > 0) {
      status = ValidationStatus.Error;
    }

    return {
      status,
      warnings,
      errors,
    };
  }

  /**
   * Lists all the OpenAPI operations in the specification file.
   * @returns A string array that represents the HTTP method and path of each operation, such as ['GET /pets/{petId}', 'GET /user/{userId}']
   * according to copilot plugin spec, only list get and post method without auth
   */
  async list(): Promise<string[]> {
    try {
      if (!this.spec) {
        this.spec = (await SwaggerParser.validate(this.specPath)) as OpenAPIV3.Document;
      }
      const apiMap = await this.getAllSupportedApi(this.spec);
      return Array.from(Object.keys(apiMap));
    } catch (err) {
      throw new SpecParserError((err as Error).toString(), ErrorType.ListFailed);
    }
  }

  /**
   * Generates and update artifacts from the OpenAPI specification file. Generate Adaptive Cards, update Teams app manifest, and generate a new OpenAPI specification file.
   * @param manifestPath A file path of the Teams app manifest file to update.
   * @param filter An array of strings that represent the filters to apply when generating the artifacts. If filter is empty, it would process nothing.
   * @param specPath An optional file path of the new OpenAPI specification file to generate. If not specified or empty, no spec file will be generated.
   * @param adaptiveCardFolder An optional folder path where the Adaptive Card files will be generated. If not specified or empty, Adaptive Card files will not be generated.
   */
  async generate(
    manifestPath: string,
    filter: string[],
    specPath?: string,
    adaptiveCardFolder?: string,
    signal?: AbortSignal
  ): Promise<void> {
    if (signal?.aborted) {
      throw new SpecParserError(ConstantString.CancelledMessage, ErrorType.Cancelled);
    }

    // TODO: implementation
  }

  private async getAllSupportedApi(
    spec: OpenAPIV3.Document
  ): Promise<{ [key: string]: OpenAPIV3.OperationObject }> {
    if (this.apiMap !== undefined) {
      return this.apiMap;
    }
    const paths = spec.paths;
    const result: { [key: string]: OpenAPIV3.OperationObject } = {};
    for (const path in paths) {
      const methods = paths[path];
      for (const method in methods) {
        // only list get and post method without auth
        if (
          (method === ConstantString.GetMethod || method === ConstantString.PostMethod) &&
          !methods[method]?.security
        ) {
          result[`${method.toUpperCase()} ${path}`] = methods[method] as OpenAPIV3.OperationObject;
        }
      }
    }
    this.apiMap = result;
    return result;
  }
}
