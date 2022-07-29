// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ContextV3,
  Effect,
  err,
  FxError,
  InputsWithProjectPath,
  ok,
  Result,
} from "@microsoft/teamsfx-api";
import "reflect-metadata";
import { Container, Service } from "typedi";
import { convertToAlphanumericOnly } from "../../common/utils";
import { BicepComponent } from "../bicep";
import "../connection/azureWebAppConfig";
import { ComponentNames } from "../constants";
import { Plans } from "../messages";
import "../resource/azureSql";
import "../resource/identity";
import { KeyVaultResource } from "../resource/keyVault";
import { generateConfigBiceps, bicepUtils } from "../utils";
import { getComponent } from "../workflow";

@Service("key-vault-feature")
export class KeyVaultFeature {
  name = "key-vault-feature";

  /**
   * 1. config keyVault
   * 2. add keyVault provision bicep
   * 3. re-generate resources that connect to key-vault
   * 4. persist bicep
   */
  async add(
    context: ContextV3,
    inputs: InputsWithProjectPath
  ): Promise<Result<undefined, FxError>> {
    const projectSettings = context.projectSetting;
    const keyVaultComponent = getComponent(projectSettings, ComponentNames.KeyVault);
    if (keyVaultComponent) return ok(undefined);
    const effects: Effect[] = [];

    // config
    projectSettings.components.push({
      name: ComponentNames.KeyVault,
      connections: [ComponentNames.Identity],
      provision: true,
    });
    effects.push(Plans.addFeature("key-vault"));
    // bicep.init
    {
      const bicepComponent = Container.get<BicepComponent>("bicep");
      const res = await bicepComponent.init(inputs.projectPath);
      if (res.isErr()) return err(res.error);
    }
    // key-vault provision bicep
    {
      const keyVaultComponent = Container.get<KeyVaultResource>(ComponentNames.KeyVault);
      const res = await keyVaultComponent.generateBicep(context, inputs);
      if (res.isErr()) return err(res.error);
      effects.push("generate key-vault provision bicep");
      const persistRes = await bicepUtils.persistBiceps(
        inputs.projectPath,
        convertToAlphanumericOnly(context.projectSetting.appName),
        res.value
      );
      if (persistRes.isErr()) return persistRes;
    }

    // generate config bicep
    {
      const res = await generateConfigBiceps(context, inputs);
      if (res.isErr()) return err(res.error);
      effects.push("update config biceps");
    }
    return ok(undefined);
  }
}
