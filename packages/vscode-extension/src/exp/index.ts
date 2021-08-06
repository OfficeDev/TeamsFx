// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";
import {
  IExperimentationService,
  IExperimentationTelemetry,
  TargetPopulation,
  getExperimentationServiceAsync,
} from "vscode-tas-client";
import * as pkg from "./../../package.json";
import { ExtTelemetry } from "./../telemetry/extTelemetry";
import * as dotenv from "dotenv";

dotenv.config();

let expService: IExperimentationService;

export function getExpService(): IExperimentationService {
  return expService;
}

export async function initialize(context: vscode.ExtensionContext) {
  if (process.env.ABTEST_CONTEXT) {
    const data = context.globalState.get("VSCode.ABExp.FeatureData") as Map<string, any>;
    data.set("assignmentContext", process.env.ABTEST_CONTEXT);
    await context.globalState.update("VSCode.ABExp.FeatureData", data);
  }
  expService = await getExperimentationServiceAsync(
    pkg.name,
    pkg.version,
    TargetPopulation.Public,
    new ExperimentationTelemetry(),
    context.globalState
  );
}

class ExperimentationTelemetry implements IExperimentationTelemetry {
  public setSharedProperty(name: string, value: string): void {
    ExtTelemetry.addSharedProperty(name, value);
  }

  public postEvent(eventName: string, props: Map<string, string>): void {
    const properties: { [key: string]: string } = {};
    props.forEach((value, key) => {
      properties[key] = value;
    });
    ExtTelemetry.sendTelemetryEvent(eventName, properties);
  }
}
