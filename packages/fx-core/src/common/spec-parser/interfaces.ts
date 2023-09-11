// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

/**
 * An interface that represents the result of validating an OpenAPI specification file.
 */
export interface ValidateResult {
  /**
   * The validation status of the OpenAPI specification file.
   */
  status: ValidationStatus;

  /**
   * An array of warning results generated during validation.
   */
  warnings: WarningResult[];

  /**
   * An array of error results generated during validation.
   */
  errors: ErrorResult[];
}

/**
 * An interface that represents a warning result generated during validation.
 */
export interface WarningResult {
  /**
   * The type of warning.
   */
  type: WarningType;

  /**
   * The content of the warning.
   */
  content: string;

  /**
   * data of the warning.
   */
  data?: any;
}

/**
 * An interface that represents an error result generated during validation.
 */
export interface ErrorResult {
  /**
   * The type of error.
   */
  type: ErrorType;

  /**
   * The content of the error.
   */
  content: string;

  /**
   * data of the error.
   */
  data?: any;
}

export interface GenerateResult {
  allSuccess: boolean;
  warnings: WarningResult[];
}

/**
 * An enum that represents the types of errors that can occur during validation.
 */
export enum ErrorType {
  SpecNotValid = "spec-not-valid",
  RemoteRefNotSupported = "remote-ref-not-supported",
  NoServerInformation = "no-server-information",
  UrlProtocolNotSupported = "url-protocol-not-supported",
  RelativeServerUrlNotSupported = "relative-server-url-not-supported",
  NoSupportedApi = "no-supported-api",
  NoExtraAPICanBeAdded = "no-extra-api-can-be-added",
  ResolveServerUrlFailed = "resolve-server-url-failed",

  ListFailed = "list-failed",
  ListOperationMapFailed = "list-operation-map-failed",
  FilterSpecFailed = "filter-spec-failed",
  UpdateManifestFailed = "update-manifest-failed",
  GenerateAdaptiveCardFailed = "generate-adaptive-card-failed",
  GenerateFailed = "generate-failed",
  ValidateFailed = "validate-failed",

  Cancelled = "cancelled",
  Unknown = "unknown",
}

/**
 * An enum that represents the types of warnings that can occur during validation.
 */
export enum WarningType {
  AuthNotSupported = "auth-not-supported",
  MethodNotSupported = "method-not-supported",
  OperationIdMissing = "operationid-missing",
  GenerateCardFailed = "generate-card-failed",
  OperationOnlyContainsOptionalParam = "operation-only-contains-optional-param",
  ConvertSwaggerToOpenAPI = "convert-swagger-to-openapi",
  Unknown = "unknown",
}

/**
 * An enum that represents the validation status of an OpenAPI specification file.
 */
export enum ValidationStatus {
  Valid,
  Warning, // If there are any warnings, the file is still valid
  Error, // If there are any errors, the file is not valid
}

export interface TextBlockElement {
  type: string;
  text: string;
  wrap: boolean;
}

export interface ArrayElement {
  type: string;
  $data: string;
  items: Array<TextBlockElement | ArrayElement>;
}

export interface AdaptiveCard {
  type: string;
  $schema: string;
  version: string;
  body: Array<TextBlockElement | ArrayElement>;
}

export interface Parameter {
  name: string;
  title: string;
  description: string;
}

export interface CheckParamResult {
  requiredNum: number;
  optionalNum: number;
  isValid: boolean;
}
