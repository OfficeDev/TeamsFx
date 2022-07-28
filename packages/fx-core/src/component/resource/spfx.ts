// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  CloudResource,
  err,
  FxError,
  InputsWithProjectPath,
  ok,
  ResourceContextV3,
  Result,
} from "@microsoft/teamsfx-api";
import "reflect-metadata";
import { Service } from "typedi";
import { SPFxPluginImpl } from "../../plugins/resource/spfx/v3/plugin";
import { ComponentNames } from "../constants";
import { ActionExecutionMW } from "../middleware/actionExecutionMW";

@Service(ComponentNames.SPFx)
export class SpfxResource implements CloudResource {
  readonly name = ComponentNames.SPFx;
  outputs = {};
  finalOutputKeys = [];
  spfxPluginImpl: SPFxPluginImpl = new SPFxPluginImpl();
  @hooks([
    ActionExecutionMW({
      enableTelemetry: true,
      telemetryComponentName: "fx-resource-spfx",
      telemetryEventName: "deploy",
      errorSource: "SPFx",
    }),
  ])
  async deploy(
    context: ResourceContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<undefined, FxError>> {
    const buildRes = await this.spfxPluginImpl.buildSPPackage(context, inputs);
    if (buildRes.isErr()) {
      return err(buildRes.error);
    }
    const deployRes = await this.spfxPluginImpl.deploy(context, inputs, context.tokenProvider!);
    if (deployRes.isErr()) {
      return err(deployRes.error);
    }
    return ok(undefined);
  }
}
