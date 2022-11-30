// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { BaseBuildDriver } from "./baseBuildDriver";
import { Service } from "typedi";
import { DriverContext } from "../interface/commonArgs";
import { FxError, Result } from "@microsoft/teamsfx-api";
import { hooks } from "@feathersjs/hooks";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { TelemetryConstant } from "../../constant/commonConstant";
import { BaseBuildStepDriver } from "./baseBuildStepDriver";

const ACTION_NAME = "dotnet/command";

@Service(ACTION_NAME)
export class DotnetBuildDriver extends BaseBuildStepDriver {
  getImpl(args: unknown, context: DriverContext): BaseBuildDriver {
    return new DotnetBuildDriverImpl(args, context);
  }

  @hooks([addStartAndEndTelemetry(ACTION_NAME, TelemetryConstant.DEPLOY_COMPONENT_NAME)])
  async run(args: unknown, context: DriverContext): Promise<Result<Map<string, string>, FxError>> {
    return super.run(args, context);
  }
}

export class DotnetBuildDriverImpl extends BaseBuildDriver {
  progressBarName = `Building Dotnet project at ${this.workingDirectory}`;
  progressBarSteps = 1;
  buildPrefix = "dotnet";
}
