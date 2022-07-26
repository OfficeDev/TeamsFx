// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Action,
  ContextV3,
  err,
  FxError,
  InputsWithProjectPath,
  MaybePromise,
  ok,
  Platform,
  Result,
} from "@microsoft/teamsfx-api";
import "reflect-metadata";
import { Container, Service } from "typedi";
import { getComponent, runActionByName } from "../workflow";
import "../connection/azureWebAppConfig";
import "../resource/azureSql";
import "../resource/identity";
import { ComponentNames } from "../constants";
import { hasApi } from "../../common/projectSettingsHelperV3";
import { convertToAlphanumericOnly } from "../../common/utils";
import { BicepComponent } from "../bicep";
import { AzureSqlResource } from "../resource/azureSql";
import { UtilFunctions } from "../resource/azureSql/actions/configure";
import { generateConfigBiceps, bicepUtils } from "../utils";
import { cloneDeep } from "lodash";

@Service("sql")
export class Sql {
  name = "sql";

  add(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): MaybePromise<Result<Action | undefined, FxError>> {
    const action: Action = {
      type: "function",
      name: "sql.add",
      question: (context, inputs) => {
        if (inputs.platform == Platform.CLI_HELP) {
          return ok(UtilFunctions.buildQuestionNode());
        }
        return ok(undefined);
      },
      execute: async (context, inputs) => {
        const sqlComponent = getComponent(context.projectSetting, ComponentNames.AzureSQL);
        const hasFunc = hasApi(context.projectSetting);
        if (!hasFunc) {
          const res = await runActionByName("teams-api.add", context, inputs);
          if (res.isErr()) return err(res.error);
        }
        const projectSettings = context.projectSetting;
        const remarks: string[] = ["config 'azure-sql' in projectSettings"];
        projectSettings.components.push({
          name: "azure-sql",
          provision: true,
        });

        // generate bicep
        // bicep.init
        {
          const bicepComponent = Container.get<BicepComponent>("bicep");
          const res = await bicepComponent.init(inputs.projectPath);
          if (res.isErr()) return err(res.error);
        }

        // sql bicep
        {
          const provisionType = sqlComponent ? "database" : "server";
          const clonedInputs = cloneDeep(inputs);
          clonedInputs.provisionType = provisionType;
          const sqlResource = Container.get<AzureSqlResource>(ComponentNames.AzureSQL);
          const res = await sqlResource.generateBicep(context, clonedInputs);
          if (res.isErr()) return err(res.error);
          const bicepRes = await bicepUtils.persistBiceps(
            inputs.projectPath,
            convertToAlphanumericOnly(context.projectSetting.appName),
            res.value
          );
          if (bicepRes.isErr()) return bicepRes;
          remarks.push("generate sql bicep");
        }

        // generate config bicep
        {
          const res = await generateConfigBiceps(context, inputs);
          if (res.isErr()) return err(res.error);
          remarks.push("generate config biceps");
        }

        return ok(remarks);
      },
    };
    return ok(action);
  }
}
