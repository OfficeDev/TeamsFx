import { FxError, ok, PluginContext, Result, Void, Plugin } from "@microsoft/teamsfx-api";

export const validManifest = {
  $schema:
    "https://developer.microsoft.com/en-us/json-schemas/teams/v1.9/MicrosoftTeams.schema.json",
  manifestVersion: "1.9",
  version: "1.0.0",
  id: "{appid}",
  packageName: "com.microsoft.teams.extension",
  developer: {
    name: "Teams App, Inc.",
    websiteUrl: "{baseUrl}",
    privacyUrl: "{baseUrl}/index.html#/privacy",
    termsOfUseUrl: "{baseUrl}/index.html#/termsofuse",
  },
  icons: {
    color: "color.png",
    outline: "outline.png",
  },
  name: {
    short: "MyApp",
    full: "This field is not used",
  },
  description: {
    short: "Short description of {appName}.",
    full: "Full description of {appName}.",
  },
  accentColor: "#FFFFFF",
  bots: [],
  composeExtensions: [],
  configurableTabs: [],
  staticTabs: [],
  permissions: ["identity", "messageTeamMembers"],
  validDomains: [],
  webApplicationInfo: {
    id: "{appClientId}",
    resource: "{webApplicationInfoResource}",
  },
};

export function mockPublishThatAlwaysSucceed(plugin: Plugin) {
  plugin.publish = async function (_ctx: PluginContext): Promise<Result<any, FxError>> {
    return ok(Void);
  };
}
