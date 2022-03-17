export class Constants {
  public static readonly pluginName: string = "Identity Plugin";
  public static readonly pluginNameShort: string = "msi";
  public static readonly prefix: string = "teamsfx";

  public static readonly apiVersion: string = "2018-11-30";
  public static readonly deployName: string = "user-assigned-identity";

  public static readonly identityName: string = "identityName";
  public static readonly identityClientId: string = "identityClientId";
  public static readonly identityResourceId: string = "identityResourceId";

  public static readonly solution: string = "solution";
  public static readonly subscriptionId: string = "subscriptionId";
  public static readonly resourceGroupName: string = "resourceGroupName";
  public static readonly resourceNameSuffix: string = "resourceNameSuffix";
  public static readonly location: string = "location";
  public static readonly remoteTeamsAppId: string = "remoteTeamsAppId";

  public static readonly resourceProvider: string = "Microsoft.ManagedIdentity";
}

export class Telemetry {
  static readonly componentName = "fx-resource-azure-identity";
  static startSuffix = "-start";
  static valueYes = "yes";
  static valueNo = "no";
  static userError = "user";
  static systemError = "system";

  static readonly stage = {
    generateArmTemplates: "generate-arm-templates",
    updateArmTemplates: "update-arm-templates",
  };

  static readonly properties = {
    component: "component",
    success: "success",
    errorCode: "error-code",
    errorType: "error-type",
    errorMessage: "error-message",
    appid: "appid",
  };
}
export class IdentityBicep {
  static readonly identityName: string = "provisionOutputs.identityOutput.value.identityName";
  static readonly identityClientId: string =
    "provisionOutputs.identityOutput.value.identityClientId";
  static readonly identityResourceId: string =
    "userAssignedIdentityProvision.outputs.identityResourceId";
  static readonly identityPrincipalId: string =
    "userAssignedIdentityProvision.outputs.identityPrincipalId";
}

export class IdentityBicepFile {
  static readonly moduleTempalteFilename: string = "identityProvision.template.bicep";
}
