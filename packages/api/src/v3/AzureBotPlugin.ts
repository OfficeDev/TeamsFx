////////////////////AzureBotPlugin.ts////////////////

import { ok } from "neverthrow";
import { FxError, Inputs, Result } from "..";
import { Context } from "../v2";
import { ResourcePlugin } from "./plugins";
import { AzureResource } from "./resourceModel";

export interface AzureBot extends AzureResource {
  type: "AzureBot";
  endpoint: string;
  botId: string;
  botPassword: string;
  aadObjectId: string; //bot AAD App Id
  appServicePlan: string; // use for deploy
  botChannelReg: string; // Azure Bot
  botRedirectUri?: string; // ???
}

export class AzureBotPlugin implements ResourcePlugin {
  name = "AzureBotPlugin";
  resourceType = "AzureBot";
  description = "Azure Bot";
  modules: ("tab" | "bot" | "backend")[] = ["bot"];
  async pluginDependencies?(ctx: Context, inputs: Inputs): Promise<Result<string[], FxError>> {
    return ok(["AzureWebAppPlugin"]);
  }
}
