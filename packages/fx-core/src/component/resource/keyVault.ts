// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "reflect-metadata";
import { Service } from "typedi";
import { ComponentNames, IdentityOutputs, KeyVaultOutputs } from "../constants";
import { AzureResource } from "./azureResource";
@Service(ComponentNames.KeyVault)
export class KeyVaultResource extends AzureResource {
  readonly name = ComponentNames.KeyVault;
  readonly bicepModuleName = "keyVault";
  outputs = KeyVaultOutputs;
  finalOutputKeys = ["keyVaultResourceId", "m365ClientSecretReference", "botClientSecretReference"];
  templateContext = {
    identity: {
      principalId: IdentityOutputs.identityPrincipalId.bicepVariable,
    },
  };
}
