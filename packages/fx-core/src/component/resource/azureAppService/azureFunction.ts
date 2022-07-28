// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  FxError,
  InputsWithProjectPath,
  ok,
  ResourceContextV3,
  Result,
} from "@microsoft/teamsfx-api";
import { Service } from "typedi";
import { FunctionOutputs, IdentityOutputs } from "../../constants";
import { AzureAppService } from "./azureAppService";
@Service("azure-function")
export class AzureFunctionResource extends AzureAppService {
  readonly name = "azure-function";
  readonly alias = "FT";
  readonly displayName = "Azure Functions";
  readonly bicepModuleName = "azureFunction";
  outputs = FunctionOutputs;
  finalOutputKeys = ["resourceId", "endpoint"];
  templateContext = {
    identity: {
      resourceId: IdentityOutputs.identityResourceId.bicepVariable,
    },
  };
  async configure(
    context: ResourceContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<undefined, FxError>> {
    return ok(undefined);
  }
}
