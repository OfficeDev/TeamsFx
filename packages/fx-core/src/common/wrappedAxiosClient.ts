// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import axios, {
  AxiosInstance,
  AxiosError,
  CreateAxiosDefaults,
  AxiosResponse,
  InternalAxiosRequestConfig,
} from "axios";
import { TOOLS } from "../core/globalVars";
import { APP_STUDIO_API_NAMES, Constants } from "../component/driver/teamsApp/constants";
import {
  TelemetryPropertyKey,
  TelemetryPropertyValue,
} from "../component/driver/teamsApp/utils/telemetry";
import { TelemetryEvent, TelemetryProperty } from "./telemetry";
import { DeveloperPortalAPIFailedError } from "../error/teamsApp";
import { HttpMethod } from "../component/constant/commonConstant";

/**
 * This client will send telemetries to record API request trace
 */
export class WrappedAxiosClient {
  public static create(config?: CreateAxiosDefaults): AxiosInstance {
    const instance = axios.create(config);

    instance.interceptors.request.use((request) => this.onRequest(request));

    // eslint-disable-next-line prettier/prettier
    instance.interceptors.response.use((response) => this.onResponse(response), (error) => this.onRejected(error));

    return instance;
  }

  /**
   * Send API start telemetry
   * @param request
   */
  public static onRequest(request: InternalAxiosRequestConfig) {
    const method = request.method!;
    const fullPath = `${request.baseURL ?? ""}${request.url ?? ""}`;
    const apiName = this.convertUrlToApiName(fullPath, method);

    const properties: { [key: string]: string } = {
      url: `<${apiName}-url>`,
      method: method,
      params: this.generateParameters(request.params),
      ...this.generateExtraProperties(fullPath, request.data),
    };

    let eventName: string;
    if (this.isTDPApi(fullPath)) {
      eventName = TelemetryEvent.AppStudioApi;
    } else {
      eventName = TelemetryEvent.DependencyApi;
    }

    TOOLS?.telemetryReporter?.sendTelemetryEvent(`${eventName}-start`, properties);
    return request;
  }

  /**
   * Send API success telemetry
   * @param response
   * @returns
   */
  public static onResponse(response: AxiosResponse) {
    const method = response.request.method;
    const fullPath = `${(response.request.host as string) ?? ""}${
      (response.request.path as string) ?? ""
    }`;
    const apiName = this.convertUrlToApiName(fullPath, method);

    const properties: { [key: string]: string } = {
      url: `<${apiName}-url>`,
      method: method,
      params: this.generateParameters(response.config.params),
      [TelemetryPropertyKey.success]: TelemetryPropertyValue.success,
      "status-code": response.status.toString(),
      ...this.generateExtraProperties(fullPath, response.data),
    };

    let eventName: string;
    if (this.isTDPApi(fullPath)) {
      eventName = TelemetryEvent.AppStudioApi;
    } else {
      eventName = TelemetryEvent.DependencyApi;
    }
    TOOLS?.telemetryReporter?.sendTelemetryEvent(eventName, properties);
    return response;
  }

  /**
   * Send API failure telemetry
   * @param error
   * @returns
   */
  public static onRejected(error: AxiosError) {
    const method = error.request.method;
    const fullPath = `${(error.request.host as string) ?? ""}${
      (error.request.path as string) ?? ""
    }`;
    const apiName = this.convertUrlToApiName(fullPath, method);

    let requestData: any;
    if (error.config?.data && typeof error.config.data === "string") {
      try {
        requestData = JSON.parse(error.config.data);
      } catch (error) {
        requestData = undefined;
      }
    }
    const properties: { [key: string]: string } = {
      url: `<${apiName}-url>`,
      method: method,
      params: this.generateParameters(error.config!.params),
      [TelemetryPropertyKey.success]: TelemetryPropertyValue.failure,
      [TelemetryPropertyKey.errorMessage]: error.response
        ? JSON.stringify(error.response.data)
        : error.message ?? "undefined",
      "status-code": error.response?.status.toString() ?? "undefined",
      ...this.generateExtraProperties(fullPath, requestData),
    };

    let eventName: string;
    if (this.isTDPApi(fullPath)) {
      const correlationId = error.response?.headers[Constants.CORRELATION_ID] ?? "undefined";
      // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
      const extraData = error.response?.data ? `data: ${JSON.stringify(error.response.data)}` : "";
      const TDPApiFailedError = new DeveloperPortalAPIFailedError(
        error,
        correlationId,
        apiName,
        extraData
      );
      properties[
        TelemetryPropertyKey.errorCode
      ] = `${TDPApiFailedError.source}.${TDPApiFailedError.name}`;
      properties[TelemetryPropertyKey.errorMessage] = TDPApiFailedError.message;
      eventName = TelemetryEvent.AppStudioApi;
    } else {
      eventName = TelemetryEvent.DependencyApi;
    }

    TOOLS?.telemetryReporter?.sendTelemetryErrorEvent(eventName, properties);
    return Promise.reject(error);
  }

  /**
   * Convert request URL to API name, otherwise it will be redacted in telemetry
   * This function should be extended when new API is added
   * @param baseUrl
   * @param path
   * @param method
   * @returns
   */
  public static convertUrlToApiName(fullPath: string, method: string): string {
    if (this.isTDPApi(fullPath)) {
      if (fullPath.match(new RegExp("/api/appdefinitions/partnerCenterAppPackageValidation"))) {
        return APP_STUDIO_API_NAMES.VALIDATE_APP_PACKAGE;
      }
      if (fullPath.match(new RegExp("/api/appdefinitions/v2/import"))) {
        return APP_STUDIO_API_NAMES.CREATE_APP;
      }
      if (fullPath.match(new RegExp("/api/appdefinitions/manifest"))) {
        return APP_STUDIO_API_NAMES.EXISTS_IN_TENANTS;
      }
      if (fullPath.match(new RegExp("/api/appdefinitions/.*/manifest"))) {
        return APP_STUDIO_API_NAMES.GET_APP_PACKAGE;
      }
      if (fullPath.match(new RegExp("/api/appdefinitions/.*/owner"))) {
        return APP_STUDIO_API_NAMES.UPDATE_OWNER;
      }
      if (
        fullPath.match(
          new RegExp(
            /\/api\/appdefinitions\/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/
          )
        )
      ) {
        if (method.toUpperCase() === HttpMethod.GET) {
          return APP_STUDIO_API_NAMES.GET_APP;
        }
        if (method.toUpperCase() === HttpMethod.DELETE) {
          return APP_STUDIO_API_NAMES.DELETE_APP;
        }
      }
      if (fullPath.match(new RegExp("/api/appdefinitions"))) {
        return APP_STUDIO_API_NAMES.LIST_APPS;
      }
      if (fullPath.match(new RegExp("/api/publishing/.*/appdefinitions"))) {
        return APP_STUDIO_API_NAMES.UPDATE_PUBLISHED_APP;
      }
      if (fullPath.match(new RegExp("/api/publishing/.*"))) {
        if (method.toUpperCase() === HttpMethod.GET) {
          return APP_STUDIO_API_NAMES.GET_PUBLISHED_APP;
        }
        if (method.toUpperCase() === HttpMethod.POST) {
          return APP_STUDIO_API_NAMES.PUBLISH_APP;
        }
      }
      if (fullPath.match(new RegExp("/api/v1.0/apiSecretRegistrations/.*"))) {
        if (method.toUpperCase() === HttpMethod.GET) {
          return APP_STUDIO_API_NAMES.GET_API_KEY;
        }
        if (method.toUpperCase() === HttpMethod.POST) {
          return APP_STUDIO_API_NAMES.CREATE_API_KEY;
        }
      }
      if (
        fullPath.match(
          new RegExp(
            /\/api\/botframework\/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/
          )
        )
      ) {
        if (method.toUpperCase() === HttpMethod.GET) {
          return APP_STUDIO_API_NAMES.GET_BOT;
        }
        if (method.toUpperCase() === HttpMethod.POST) {
          return APP_STUDIO_API_NAMES.UPDATE_BOT;
        }
        if (method.toUpperCase() === HttpMethod.DELETE) {
          return APP_STUDIO_API_NAMES.DELETE_BOT;
        }
      }
      if (fullPath.match(new RegExp("/api/botframework"))) {
        if (method.toUpperCase() === HttpMethod.GET) {
          return APP_STUDIO_API_NAMES.LIST_BOT;
        }
        if (method.toUpperCase() === HttpMethod.POST) {
          return APP_STUDIO_API_NAMES.CREATE_BOT;
        }
      }
    }
    if (
      fullPath.match(
        new RegExp(/(^https:\/\/)?authsvc\.teams\.microsoft\.com\/v1\.0\/users\/region/)
      )
    ) {
      return "get-region";
    }
    return fullPath.replace(/\//g, `-`);
  }

  /**
   * Generate extra properties for specific requirements
   * @param baseUrl
   * @param path
   * @param method
   */
  private static generateExtraProperties(fullPath: string, data?: any): { [key: string]: string } {
    const properties: { [key: string]: string } = {};
    if (this.isTDPApi(fullPath)) {
      // Add region property
      properties[TelemetryPropertyKey.region] = String(this.extractRegion(fullPath));

      // Add bot id property
      if (fullPath.match(new RegExp("/api/botframework"))) {
        const regex = new RegExp(/\/api\/botframework\/([0-9a-fA-F-]+)/);
        const matches = regex.exec(fullPath);
        if (matches != null && matches.length > 1) {
          properties[TelemetryProperty.BotId] = matches[1];
        } else if (data?.botId) {
          properties[TelemetryProperty.BotId] = data.botId;
        }
      }
    }
    return properties;
  }

  /**
   * Extract region from baseUrl, E.g. https://dev.teams.microsoft.com/amer => amer
   * @param url
   * @returns
   */
  private static extractRegion(fullPath: string): string | undefined {
    const regex = /dev(-int)?\.teams\.microsoft\.com\/([a-zA-Z-_]+)\/api/;
    const matches = regex.exec(fullPath);
    if (matches != null && matches.length > 1) {
      return matches[2];
    }

    return TelemetryPropertyValue.Global;
  }

  /**
   * Check if it's TDP Api
   * @param baseUrl
   * @returns
   */
  private static isTDPApi(baseUrl: string): boolean {
    const regex = /(^https:\/\/)?dev(-int)?\.teams\.microsoft\.com/;
    const matches = regex.exec(baseUrl);
    return matches != null && matches.length > 0;
  }

  /**
   * Flattern query parameters to string, e.g. {a: 1, b: 2} => a:1;b:2
   * @param params
   * @returns
   */
  private static generateParameters(params?: Record<string, unknown>): string {
    if (!params) {
      return "";
    }
    const parameters: string[] = [];
    for (const [key, value] of Object.entries(params)) {
      parameters.push(`${key}:${value as string}`);
    }
    return parameters.join(";");
  }
}
