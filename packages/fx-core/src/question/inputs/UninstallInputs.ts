// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/****************************************************************************************
 *                            NOTICE: AUTO-GENERATED                                    *
 ****************************************************************************************
 * This file is automatically generated by script "./src/question/generator.ts".        *
 * Please don't manually change its contents, as any modifications will be overwritten! *
 ***************************************************************************************/

import { Inputs } from "@microsoft/teamsfx-api";

export interface UninstallInputs extends Inputs {
  /** @description Choose uninstall mode */
  mode?: "manifest-id" | "env" | "title-id";
  /** @description Manifest ID */
  "mainfest-id"?: string;
  /** @description Env */
  env?: string;
  /** @description Project path */
  projectPath?: string;
  /** @description Choose resources to uninstall */
  options?: "m365-app" | "app-registration" | "bot-framework-registration"[];
  /** @description Title ID */
  "title-id"?: string;
}
