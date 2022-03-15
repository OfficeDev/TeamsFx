// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

/**
 * Config key contract that value is provided by local debug plugin and required by other plugins.
 */
export class LocalDebugConfigKeys {
  public static readonly LocalAuthEndpoint: string = "localAuthEndpoint";

  public static readonly LocalTabEndpoint: string = "localTabEndpoint";
  public static readonly LocalTabDomain: string = "localTabDomain";
  public static readonly TrustDevelopmentCertificate: string = "trustDevCert";

  public static readonly LocalFunctionEndpoint: string = "localFunctionEndpoint";

  public static readonly LocalBotEndpoint: string = "localBotEndpoint";
  public static readonly LocalBotDomain: string = "localBotDomain";
}

export class SolutionPlugin {
  public static readonly Name: string = "solution";
  // public static readonly SelectedPlugins: string = "selectedPlugins";
  public static readonly LocalTeamsAppId: string = "localDebugTeamsAppId";
  public static readonly RemoteTeamsAppId: string = "remoteTeamsAppId";
  public static readonly TeamsAppTenantId: string = "teamsAppTenantId";
  public static readonly ProgrammingLanguage: string = "programmingLanguage";
}

export class AppStudioPlugin {
  public static readonly Name: string = "fx-resource-appstudio";
  public static readonly TeamsAppId: string = "teamsAppId";
}
