/* eslint-disable @typescript-eslint/ban-types */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError } from "@microsoft/teamsfx-api";
import { RestError } from "@azure/ms-rest-js";
import path from "path";

/**
 * Void is used to construct Result<Void, FxError>.
 * e.g. return ok(Void);
 * It exists because ok(void) does not compile.
 */
export type Void = {};
export const Void = {};

/**
 * The key of global config visible to all resource plugins.
 */
export const GLOBAL_CONFIG = "solution";
// export const SELECTED_PLUGINS = "selectedPlugins";

/**
 * Used to track whether provision succeeded
 * Set to true when provison succeeds, to false when a new resource is added.
 */
export const SOLUTION_PROVISION_SUCCEEDED = "provisionSucceeded";

/**
 * Config key whose value is either javascript, typescript or csharp.
 */
export const PROGRAMMING_LANGUAGE = "programmingLanguage";

/**
 * Config key whose value is the default function name for adding a new function.
 */
export const DEFAULT_FUNC_NAME = "defaultFunctionName";

/**
 * Config key whose value is output of ARM templates deployment.
 */
export const ARM_TEMPLATE_OUTPUT = "armTemplateOutput";
export const TEAMS_FX_RESOURCE_ID_KEY = "teamsFxPluginId";

/**
 * Config key whose value is the resource group name of project.
 */
export const RESOURCE_GROUP_NAME = "resourceGroupName";

/**
 * Config key whose value is the resource group location of project.
 */
export const LOCATION = "location";

/**
 * Config key whose value is the subscription ID of project.
 */
export const SUBSCRIPTION_ID = "subscriptionId";

/**
 * Config key whose value is the subscription name of project.
 */
export const SUBSCRIPTION_NAME = "subscriptionName";

export const DEFAULT_PERMISSION_REQUEST = [
  {
    resource: "Microsoft Graph",
    delegated: ["User.Read"],
    application: [],
  },
];

export enum PluginNames {
  SQL = "fx-resource-azure-sql",
  MSID = "fx-resource-identity",
  FE = "fx-resource-frontend-hosting",
  SPFX = "fx-resource-spfx",
  BOT = "fx-resource-bot",
  AAD = "fx-resource-aad-app-for-teams",
  FUNC = "fx-resource-function",
  SA = "fx-resource-simple-auth",
  LDEBUG = "fx-resource-local-debug",
  APIM = "fx-resource-apim",
  APPST = "fx-resource-appstudio",
  SOLUTION = "solution",
}

export enum SolutionError {
  InvalidSelectedPluginNames = "InvalidSelectedPluginNames",
  PluginNotFound = "PluginNotFound",
  AADPluginNotEnabled = "AADPluginNotEnabled",
  MissingPermissionsJson = "MissingPermissionsJson",
  DialogIsNotPresent = "DialogIsNotPresent",
  NoResourcePluginSelected = "NoResourcePluginSelected",
  NoAppStudioToken = "NoAppStudioToken",
  NoTeamsAppTenantId = "NoTeamsAppTenantId",
  NoUserName = "NoUserName",
  FailedToCreateResourceGroup = "FailedToCreateResourceGroup",
  FailedToListResourceGroup = "FailedToListResourceGrouop",
  FailedToListResourceGroupLocation = "FailedToListResourceGroupLocation",
  FailedToGetResourceGroupInfoInputs = "FailedToGetResourceGroupInfoInputs",
  ResourceGroupNotFound = "ResourceGroupNotFound",
  SubscriptionNotFound = "SubscriptionNotFound",
  NotLoginToAzure = "NotLoginToAzure",
  AzureAccountExtensionNotInitialized = "AzureAccountExtensionNotInitialized",
  LocalTabEndpointMissing = "LocalTabEndpointMissing",
  LocalTabDomainMissing = "LocalTabDomainMissing",
  LocalClientIDMissing = "LocalDebugClientIDMissing",
  LocalApplicationIdUrisMissing = "LocalApplicationIdUrisMissing",
  LocalClientSecretMissing = "LocalClientSecretMissing",
  CannotUpdatePermissionForSPFx = "CannotUpdatePermissionForSPFx",
  CannotAddResourceForSPFx = "CannotAddResourceForSPFx",
  FailedToParseAzureTenantId = "FailedToParseAzureTenantId",
  CannotRunProvisionInSPFxProject = "CannotRunProvisionInSPFxProject",
  CannotRunThisTaskInSPFxProject = "CannotRunThisTaskInSPFxProject",
  FrontendEndpointAndDomainNotFound = "FrontendEndpointAndDomainNotFound",
  RemoteClientIdNotFound = "RemoteClientIdNotFound",
  AddResourceNotSupport = "AddResourceNotSupport",
  AddCapabilityNotSupport = "AddCapabilityNotSupport",
  FailedToAddCapability = "FailedToAddCapability",
  NoResourceToDeploy = "NoResourceToDeploy",
  ProvisionInProgress = "ProvisionInProgress",
  DeploymentInProgress = "DeploymentInProgress",
  PublishInProgress = "PublishInProgress",
  UnknownSolutionRunningState = "UnknownSolutionRunningState",
  CannotDeployBeforeProvision = "CannotDeployBeforeProvision",
  CannotPublishBeforeProvision = "CannotPublishBeforeProvision",
  CannotLocalDebugInDifferentTenant = "CannotLocalDebugInDifferentTenant",
  NoSubscriptionFound = "NoSubscriptionFound",
  NoSubscriptionSelected = "NoSubscriptionSelected",
  FailedToGetParamForRegisterTeamsAppAndAad = "FailedToGetParamForRegisterTeamsAppAndAad",
  BotInternalError = "BotInternalError",
  InternelError = "InternelError",
  RegisterTeamsAppAndAadError = "RegisterTeamsAppAndAadError",
  GetLocalDebugConfigError = "GetLocalDebugConfigError",
  GetRemoteConfigError = "GetRemoteConfigError",
  UnsupportedPlatform = "UnsupportedPlatform",
  InvalidInput = "InvalidInput",
  FailedToCompileBicepFiles = "FailedToCompileBicepFiles",
  FailedToGetAzureCredential = "FailedToGetAzureCredential",
  FailedToGenerateArmTemplates = "FailedToGenerateArmTemplates",
  FailedToUpdateArmParameters = "FailedToUpdateArmTemplates",
  FailedToDeployArmTemplatesToAzure = "FailedToDeployArmTemplatesToAzure",
  FailedToPollArmDeploymentStatus = "FailedToPollArmDeploymentStatus",
  FailedToValidateArmTemplates = "FailedToValidateArmTemplates",
  FailedToRetrieveUserInfo = "FailedToRetrieveUserInfo",
  FeatureNotSupported = "FeatureNotSupported",
  CannotFindUserInCurrentTenant = "CannotFindUserInCurrentTenant",
  FailedToGrantPermission = "FailedToGrantPermission",
  FailedToCheckPermission = "FailedToCheckPermission",
  FailedToListCollaborator = "FailedToListCollaborator",
  EmailCannotBeEmptyOrSame = "EmailCannotBeEmptyOrSame",
  FailedToExecuteTasks = "FailedToExecuteTasks",
  FailedToGetEnvName = "FailedToGetEnvName",
  TeamsAppTenantIdNotRight = "TeamsAppTenantIdNotRight",
  AddSsoNotSupported = "AddSsoNotSupported",
  NeedEnableFeatureFlag = "NeedEnableFeatureFlag",
  SsoEnabled = "SsoEnabled",
  InvalidSsoProject = "InvalidSsoProject",
  InvalidProjectPath = "InvalidProjectPath",
  FailedToCreateAuthFiles = "FailedToCreateAuthFiles",
}

export const LOCAL_DEBUG_TAB_ENDPOINT = "localTabEndpoint";
export const LOCAL_DEBUG_TAB_DOMAIN = "localTabDomain";
export const LOCAL_DEBUG_BOT_DOMAIN = "localBotDomain";
export const BOT_DOMAIN = "validDomain";
export const BOT_SECTION = "bots";
export const COMPOSE_EXTENSIONS_SECTION = "composeExtensions";
export const LOCAL_WEB_APPLICATION_INFO_SOURCE = "local_applicationIdUris";
export const WEB_APPLICATION_INFO_SOURCE = "applicationIdUris";
export const LOCAL_DEBUG_AAD_ID = "local_clientId";
export const REMOTE_AAD_ID = "clientId";
export const LOCAL_APPLICATION_ID_URIS = "local_applicationIdUris";
export const REMOTE_APPLICATION_ID_URIS = "applicationIdUris";
export const LOCAL_CLIENT_SECRET = "local_clientSecret";
export const REMOTE_CLIENT_SECRET = "clientSecret";
export const REMOTE_TEAMS_APP_TENANT_ID = "teamsAppTenantId";
export const LOCAL_TENANT_ID = "local_tenantId";
// Teams App Id for local debug
export const LOCAL_DEBUG_TEAMS_APP_ID = "localDebugTeamsAppId";
// Teams App Id for remote
export const REMOTE_TEAMS_APP_ID = "remoteTeamsAppId";
export const TEAMS_APP_ID = "teamsAppId";

export const AzureRoleAssignmentsHelpLink =
  "https://aka.ms/teamsfx-azure-role-assignments-help-link";
export const SharePointManageSiteAdminHelpLink =
  "https://aka.ms/teamsfx-sharepoint-manage-site-admin-help-link";

export const ViewAadAppHelpLink = "https://aka.ms/teamsfx-view-aad-app";

export const DoProvisionFirstError = new UserError(
  "DoProvisionFirst",
  "DoProvisionFirst",
  "Solution"
);
export const CancelError = new UserError("Solution", "UserCancel", "UserCancel");
// This is the max length specified in
// https://developer.microsoft.com/en-us/json-schemas/teams/v1.7/MicrosoftTeams.schema.json

export enum SolutionTelemetryEvent {
  CreateStart = "create-start",
  Create = "create",

  AddResourceStart = "add-resource-start",
  AddResource = "add-resource",

  AddCapabilityStart = "add-capability-start",
  AddCapability = "add-capability",

  GrantPermissionStart = "grant-permission-start",
  GrantPermission = "grant-permission",

  CheckPermissionStart = "check-permission-start",
  CheckPermission = "check-permission",

  ListCollaboratorStart = "list-collaborator-start",
  ListCollaborator = "list-collaborator",

  GenerateArmTemplateStart = "generate-armtemplate-start",
  GenerateArmTemplate = "generate-armtemplate",

  ArmDeploymentStart = "deploy-armtemplate-start",
  ArmDeployment = "deploy-armtemplate",

  AddSsoStart = "add-sso-start",
  AddSso = "add-sso",
  AddSsoReadme = "add-sso-readme",

  DeployStart = "deploy-start",
  Deploy = "deploy",

  ProvisionStart = "provision-start",
  Provision = "provision",
}

export enum SolutionTelemetryProperty {
  Component = "component",
  Resources = "resources",
  Capabilities = "capabilities",
  Success = "success",
  CollaboratorCount = "collaborator-count",
  AadOwnerCount = "aad-owner-count",
  AadPermission = "aad-permission",
  ArmDeploymentError = "arm-deployment-error",
  TeamsAppPermission = "teams-app-permission",
  ProgrammingLanguage = "programming-language",
  Env = "env",
  IncludeAadManifest = "include-aad-manifest",
  ErrorCode = "error-code",
  ErrorMessage = "error-message",
  HostType = "host-type",
  SubscriptionId = "subscription-id",
  AddTabSso = "tab-sso",
  AddBotSso = "bot-sso",
}

export enum SolutionTelemetrySuccess {
  Yes = "yes",
  No = "no",
}

export const SolutionTelemetryComponentName = "solution";
export const SolutionSource = "Solution";

export class UnauthorizedToCheckResourceGroupError extends UserError {
  constructor(resourceGroupName: string, subscriptionId: string, subscriptionName: string) {
    const subscriptionInfoString =
      subscriptionId + (subscriptionName.length > 0 ? `(${subscriptionName})` : "");
    super(
      SolutionSource,
      new.target.name,
      `Unauthorized to check the existence of resource group '${resourceGroupName}' in subscription '${subscriptionInfoString}'. Please check your Azure subscription.`
    );
  }
}

export class FailedToCheckResourceGroupExistenceError extends UserError {
  constructor(
    error: unknown,
    resourceGroupName: string,
    subscriptionId: string,
    subscriptionName: string
  ) {
    const subscriptionInfoString =
      subscriptionId + (subscriptionName.length > 0 ? `(${subscriptionName})` : "");
    const baseErrorMessage = `Failed to check the existence of resource group '${resourceGroupName}' in subscription '${subscriptionInfoString}'`;

    if (error instanceof RestError) {
      // Avoid sensitive information like request headers in the error message.
      const rawErrorString = JSON.stringify({
        code: error.code,
        statusCode: error.statusCode,
        body: error.body,
        name: error.name,
        message: error.message,
      });

      super(SolutionSource, new.target.name, `${baseErrorMessage}, error: '${rawErrorString}'`);
    } else if (error instanceof Error) {
      // Reuse the original error object to prevent losing the stack info
      error.message = `${baseErrorMessage}, error: '${error.message}'`;
      super({ error, source: SolutionSource });
    } else {
      super(
        SolutionSource,
        new.target.name,
        `${baseErrorMessage}, error: '${JSON.stringify(error)}'`
      );
    }
  }
}

export enum Language {
  JavaScript = "javascript",
  TypeScript = "typescript",
  CSharp = "csharp",
}

export class AddSsoParameters {
  static readonly filePath = path.join("plugins", "resource", "aad", "auth");
  static readonly Bot = "bot";
  static readonly Tab = "tab";
  static readonly Readme = "README.md";
  static readonly LearnMore = "Learn More";
  static readonly LearnMoreUrl = "https://aka.ms/teamsfx-add-sso-readme";
  static readonly AddSso = "addSso";
}

export class UserTaskFunctionName {
  static readonly ConnectExistingApi = "connectExistingApi";
}

export interface ProvisionSubscriptionCheckResult {
  hasSwitchedSubscription: boolean;
}

export type FillInAzureConfigsResult = ProvisionSubscriptionCheckResult;
