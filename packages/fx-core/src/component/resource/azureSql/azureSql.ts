// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  FxError,
  ok,
  Result,
  CloudResource,
  ContextV3,
  InputsWithProjectPath,
  Bicep,
  ResourceContextV3,
  err,
} from "@microsoft/teamsfx-api";
import "reflect-metadata";
import { Service } from "typedi";
import { AzureSqlOutputs, ComponentNames } from "../../constants";
import { ConfigureActionImplement } from "./actions/configure";
import { ProvisionActionImplement } from "./actions/provision";
import * as path from "path";
import fs from "fs-extra";
import { getTemplatesFolder } from "../../../folder";
import { generateBicepFromFile, getUuid } from "../../../common/tools";
@Service("azure-sql")
export class AzureSqlResource implements CloudResource {
  readonly name = ComponentNames.AzureSQL;
  readonly outputs = AzureSqlOutputs;
  readonly finalOutputKeys = ["sqlResourceId", "endpoint", "databaseName"];
  async generateBicep(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<Bicep[], FxError>> {
    const prefix =
      inputs.provisionType === "database"
        ? "azureSql.provisionDatabase"
        : "azureSql.provisionServer";
    const mPath = path.join(getTemplatesFolder(), "bicep", `${prefix}.module.bicep`);
    const oPath = path.join(getTemplatesFolder(), "bicep", `${prefix}.orchestration.bicep`);
    let module = await fs.readFile(mPath, "utf-8");
    let orch = await fs.readFile(oPath, "utf-8");
    const suffix = getUuid().substring(0, 6);
    const compileCtx = {
      suffix: suffix,
    };
    if (inputs.provisionType === "database") {
      module = await generateBicepFromFile(mPath, compileCtx);
      orch = await generateBicepFromFile(oPath, compileCtx);
    }
    const bicep: Bicep = {
      type: "bicep",
      Provision: {
        Modules: { azureSql: module },
        Orchestration: orch,
      },
    };
    if (inputs.provisionType === "server") {
      bicep.Parameters = await fs.readJson(
        path.join(getTemplatesFolder(), "bicep", "azureSql.parameters.json")
      );
    }
    return ok([bicep]);
  }
  async provision(
    context: ResourceContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<undefined, FxError>> {
    const res = await ProvisionActionImplement.execute(context, inputs);
    if (res.isErr()) return err(res.error);
    return ok(undefined);
  }
  async configure(
    context: ResourceContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<undefined, FxError>> {
    const res = await ConfigureActionImplement.execute(context, inputs);
    if (res.isErr()) return err(res.error);
    return ok(undefined);
  }
}
