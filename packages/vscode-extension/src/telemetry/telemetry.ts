// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { configure, getLogger, Logger } from "log4js";
import * as os from "os";
import * as path from "path";
// eslint-disable-next-line import/no-unresolved
import * as vscode from "vscode";

import {
  ConfigFolderName,
  FxError,
  Stage,
  TelemetryReporter,
  UserError,
} from "@microsoft/teamsfx-api";
import { Correlator, getFixedCommonProjectSettings } from "@microsoft/teamsfx-core";
/* eslint-disable-next-line import/default */
import Reporter from "@vscode/extension-telemetry";

import * as extensionPackage from "../../package.json";
import * as globalVariables from "../globalVariables";
import {
  FeatureFlags,
  getAllFeatureFlags,
  getProjectId,
  isFeatureFlagEnabled,
} from "../utils/commonUtils";
import { TelemetryCache } from "./cache";
import {
  TelemetryComponentType,
  TelemetryErrorType,
  TelemetryEvent,
  TelemetryEventCache,
  TelemetryProperty,
  TelemetrySuccess,
} from "./extTelemetryEvents";

const TelemetryTestLoggerFile = "telemetryTest.log";

/* eslint-disable prefer-const */
let isFromSample: boolean | undefined = undefined;
let settingsVersion: string | undefined = undefined;
let isM365: boolean | undefined = undefined;

export let reporter: VSCodeTelemetryReporter;
// export for UT
export let lastCorrelationId: string | undefined = undefined;

/**
 *  VSCode telemetry reporter used by fx-core and extension.
 *  Usage:
 *    let reporter = new VSCodeTelemetryReporter(key, extensionVersion, extensionId)
 *  Illustrate:
 *    key = <'the application insights key'>, 'aiKey' in package.json
 *    extensionVersion = '<extension version>', extension version will be reported as a property with each event
 *    extensionId = '<your extension unique name>', all events will be prefixed with this event name. eg: 'extensionId/eventname'
 */
export class VSCodeTelemetryReporter extends vscode.Disposable implements TelemetryReporter {
  private readonly reporter: Reporter;
  private readonly logger: Logger | undefined;
  private readonly testFeatureFlag: boolean;
  /**
   * events are cached in memory first and sent later to reduce date loss.
   * strategy: send at least once.
   * assumption: if an event is sent to server 10 seconds ago, it is treated as successfully sent telemetry event.
   */
  private cache: TelemetryCache;

  private sharedProperties: { [key: string]: string } = {};

  constructor(key: string, extensionVersion: string, extensionId: string) {
    super(async () => await this.reporter.dispose());
    this.reporter = new Reporter(extensionId, extensionVersion, key, true);
    this.testFeatureFlag = isFeatureFlagEnabled(FeatureFlags.TelemetryTest);
    if (this.testFeatureFlag) {
      const logFile = path.join(os.homedir(), `.${ConfigFolderName}`, TelemetryTestLoggerFile);
      configure({
        appenders: { everything: { type: "file", filename: logFile } },
        categories: { default: { appenders: ["everything"], level: "debug" } },
      });
      this.logger = getLogger("TelemTest");
    }
    this.cache = new TelemetryCache(this.reporter);
  }

  addSharedProperty(name: string, value?: string): void {
    this.sharedProperties[name] = value ?? "";
  }

  public sendTelemetryEvent(
    eventName: string,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {
    const filledProperties = this.fillProperties(properties);

    if (this.testFeatureFlag) {
      this.logTelemetryEvent(eventName, filledProperties, measurements);
    } else {
      const eventCache: TelemetryEventCache = {
        type: "normal",
        occurTime: new Date(),
        eventName,
        properties: filledProperties,
        measurements,
      };
      this.cache.addEvent(eventCache);
    }
  }

  public sendTelemetryErrorEvent(
    eventName: string,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number },
    errorProps?: string[]
  ): void {
    const filledProperties = this.fillProperties(properties);

    if (this.testFeatureFlag) {
      this.logTelemetryErrorEvent(eventName, filledProperties, measurements, errorProps);
    } else {
      const eventCache: TelemetryEventCache = {
        type: "error",
        occurTime: new Date(),
        eventName,
        properties: filledProperties,
        measurements,
      };
      this.cache.addEvent(eventCache);
    }
  }

  public sendTelemetryException(
    error: Error,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {
    const filledProperties = this.fillProperties(properties);

    if (this.testFeatureFlag) {
      this.logTelemetryException(error, filledProperties, measurements);
    } else {
      this.reporter.sendTelemetryException(error, filledProperties, measurements);
    }
  }

  public async recover() {
    await this.cache.recoverUnsentEventsFromDiskAsync();
  }

  private fillProperties(properties?: { [p: string]: string }): { [p: string]: string } {
    if (!properties) {
      properties = { ...this.sharedProperties };
    } else {
      properties = { ...this.sharedProperties, ...properties };
    }

    this.checkAndOverwriteSharedProperty(properties);

    lastCorrelationId = Correlator.getId();
    if (properties[TelemetryProperty.CorrelationId] === undefined) {
      // deactivate event will set correlation id and should not be overwritten
      properties[TelemetryProperty.CorrelationId] = Correlator.getId();
    }

    if (properties[TelemetryProperty.FeatureFlags] === undefined) {
      const featureFlags = getAllFeatureFlags();
      properties[TelemetryProperty.FeatureFlags] = featureFlags ? featureFlags.join(";") : "";
    }

    return properties;
  }

  private checkAndOverwriteSharedProperty(properties: { [p: string]: string }) {
    if (
      !properties[TelemetryProperty.ProjectId] ||
      !properties[TelemetryProperty.ProgrammingLanguage] ||
      !properties[TelemetryProperty.IsFromSample]
    ) {
      const fixedProjectSettings = getFixedCommonProjectSettings(
        globalVariables.workspaceUri?.fsPath
      );

      if (fixedProjectSettings?.projectId) {
        properties[TelemetryProperty.ProjectId] = fixedProjectSettings?.projectId;
        this.sharedProperties[TelemetryProperty.ProjectId] = fixedProjectSettings?.projectId;
      }

      if (fixedProjectSettings?.programmingLanguage) {
        properties[TelemetryProperty.ProgrammingLanguage] =
          fixedProjectSettings?.programmingLanguage;
        this.sharedProperties[TelemetryProperty.ProgrammingLanguage] =
          fixedProjectSettings?.programmingLanguage;
      }

      if (fixedProjectSettings?.isFromSample) {
        properties[TelemetryProperty.IsFromSample] = fixedProjectSettings?.isFromSample;
        this.sharedProperties[TelemetryProperty.IsFromSample] = fixedProjectSettings?.isFromSample;
      }

      if (fixedProjectSettings?.isM365) {
        properties[TelemetryProperty.IsM365] = fixedProjectSettings?.isM365;
        this.sharedProperties[TelemetryProperty.IsM365] = fixedProjectSettings?.isM365;
      }
    }
  }

  public async dispose() {
    const deactivateEvent: TelemetryEventCache = {
      type: "normal",
      occurTime: new Date(),
      eventName: TelemetryEvent.Deactivate,
      properties: {
        [TelemetryProperty.CorrelationId]: lastCorrelationId || "",
        [TelemetryProperty.ProjectId]: getProjectId() || "",
      },
    };
    this.cache.sendEventsInCache();
    await this.cache.persistUncertainEventsToDiskAsync(deactivateEvent);
    await this.reporter.dispose();
  }

  private logTelemetryEvent(
    eventName: string,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {
    this.logger?.debug(eventName, properties, measurements);
  }

  private logTelemetryErrorEvent(
    eventName: string,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number },
    errorProps?: string[]
  ): void {
    this.logger?.debug(eventName, properties, measurements, errorProps);
  }

  private logTelemetryException(
    error: Error,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {
    this.logger?.debug(error, properties, measurements);
  }
}

export function initializeExtensionTelemetryReporter() {
  reporter = new VSCodeTelemetryReporter(
    extensionPackage.aiKey,
    extensionPackage.version,
    extensionPackage.name
  );
}

export function addSharedProperty(name: string, value: string): void {
  reporter.addSharedProperty(name, value);
}

export function stageToEvent(stage: Stage): string | undefined {
  /* debug telemetry event is not handling here */
  switch (stage) {
    case Stage.create:
      return TelemetryEvent.CreateProject;
    case Stage.init:
      return TelemetryEvent.InitProject;
    case Stage.provision:
      return TelemetryEvent.Provision;
    case Stage.deploy:
      return TelemetryEvent.Deploy;
    case Stage.publish:
      return TelemetryEvent.Publish;
    case Stage.createEnv:
      return TelemetryEvent.CreateNewEnvironment;
    case Stage.grantPermission:
      return TelemetryEvent.GrantPermission;
    default:
      return undefined;
  }
}

export function sendTelemetryEvent(
  eventName: string,
  properties?: { [p: string]: string },
  measurements?: { [p: string]: number }
): void {
  properties = mergeExtensionProperties(properties);

  reporter.sendTelemetryEvent(eventName, properties, measurements);
}

export function sendTelemetryErrorEvent(
  eventName: string,
  error: FxError,
  properties?: { [p: string]: string },
  measurements?: { [p: string]: number },
  errorProps?: string[]
): void {
  properties = mergeExtensionProperties(properties);

  properties[TelemetryProperty.Success] = TelemetrySuccess.No;
  if (error instanceof UserError) {
    properties[TelemetryProperty.ErrorType] = TelemetryErrorType.UserError;
  } else {
    properties[TelemetryProperty.ErrorType] = TelemetryErrorType.SystemError;
  }

  properties[TelemetryProperty.ErrorCode] = `${error.source}.${error.name}`;
  properties[TelemetryProperty.ErrorMessage] = `${error.message}${
    error.stack ? "\nstack:\n" + error.stack : ""
  }`;

  reporter.sendTelemetryErrorEvent(eventName, properties, measurements, errorProps);
}

export function sendTelemetryException(
  error: Error,
  properties?: { [p: string]: string },
  measurements?: { [p: string]: number }
): void {
  properties = mergeExtensionProperties(properties);

  reporter.sendTelemetryException(error, properties, measurements);
}

export async function initializeTelemetry(
  _isFromSample?: boolean,
  _isM365?: boolean,
  _settingsVersion?: string
) {
  isFromSample = _isFromSample;
  isM365 = _isM365;
  settingsVersion = _settingsVersion;
  await reporter.recover();
}

export async function dispose() {
  await reporter.dispose();
}

function mergeExtensionProperties(properties?: { [p: string]: string }): { [p: string]: string } {
  if (!properties) {
    properties = {};
  }

  if (TelemetryProperty.Component in properties === false) {
    properties[TelemetryProperty.Component] = TelemetryComponentType;
  }

  properties[TelemetryProperty.IsExistingUser] = globalVariables.isExistingUser;

  if (globalVariables.workspaceUri) {
    properties[TelemetryProperty.IsSpfx] = globalVariables.isSPFxProject.toString();
  }

  if (isFromSample !== undefined) {
    properties[TelemetryProperty.IsFromSample] = isFromSample.toString();
  }
  if (isM365 !== undefined) {
    properties[TelemetryProperty.IsM365] = isM365.toString();
  }
  if (settingsVersion !== undefined) {
    properties[TelemetryProperty.SettingsVersion] = settingsVersion.toString();
  }

  return properties;
}
