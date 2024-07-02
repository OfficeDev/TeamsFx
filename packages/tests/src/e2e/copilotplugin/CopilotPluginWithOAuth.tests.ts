// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { Capability } from "../../utils/constants";
import { CaseFactory } from "../caseFactory";
import { ProgrammingLanguage } from "@microsoft/teamsfx-core";

class CopilotPluginOAuthTestCase extends CaseFactory {}

const copilotPluginOAuth: Record<string, string> = {};
copilotPluginOAuth["api-auth"] = "oauth";

new CopilotPluginOAuthTestCase(
  Capability.CopilotPluginFromScratch,
  27569691,
  "huimiao@microsoft.com",
  ["function"],
  ProgrammingLanguage.TS,
  {},
  copilotPluginOAuth
).test();
