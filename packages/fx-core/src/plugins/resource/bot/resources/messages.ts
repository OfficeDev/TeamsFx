// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { Links } from "../constants";

export class Messages {
    public static readonly SomethingIsInvalidWithValue = (something: string, value: string): string => `'${something}' is invalid with '${value}'.`;
    public static readonly InputValidValueForSomething = (something: string): string => `Please select valid values for '${something}'.`;
    public static readonly SomethingIsMissing = (something: string): string => `'${something}' is missing.`;
    public static readonly SomethingIsNotFound = (something: string): string => `'${something}' is not found.`;
    public static readonly SomethingIsNotExisting = (something: string): string => `'${something}' is not existing.`;
    public static readonly SomethingIsInWrongFormat = (something: string): string => `'${something}' is in wrong format.`;
    public static readonly FailToCreateSomeClient = (clientName: string): string => `Failed to create '${clientName}'.`;
    public static readonly FailToProvisionSomeResource = (resource: string): string => `Failed to provision '${resource}'.`;
    public static readonly FailToUpdateConfigs = (something: string): string => `Failed to update configs for '${something}'.`;
    public static readonly FailToListPublishingCredentials = "Failed to list publishing credentials.";
    public static readonly FailToDoZipDeploy = "Failed to deploy zip file.";
    public static readonly FailToUpdateMessageEndpoint = (endpoint: string): string => `Failed to update message endpoint with '${endpoint}'.`;
    public static readonly FailToDownloadFrom = (url: string): string => `Failed to download from '${url}'.`;
    public static readonly FailToFindSomethingFor = (something: string, forsth: string): string => `Failed to retrieve '${something}' for '${forsth}'.`;
    public static readonly ReferToIssueLink = `Refer to ${Links.ISSUE_LINK}.`;
    public static readonly ReferToHelpLink = `Refer to ${Links.HELP_LINK}.`;
    public static readonly CommandFailWithMessage = (command: string, message: string): string => `Run '${command}' failed with message: ${message}`;
    public static readonly DoSthBeforeSth = (sth: string, beforeSth: string): string => `Perform command '${sth}' before '${beforeSth}'.`;
    public static readonly FailToCallAppStudio = (apiName: string): string => `Failed to execute '${apiName}'.`;
    public static readonly SuccessfullyRetrievedTemplateZip = (zipUrl: string): string => `Successfully retrieved zip package from ${zipUrl}.`;
    public static readonly FallingBackToUseLocalTemplateZip = "Falling back to use local template zip.";

    public static readonly WorkingDirIsMissing = "Working directory is missing.";
    public static readonly FailToGetAzureCreds = "Failed to retrieve Azure credentials.";
    public static readonly TryLoginAzure = "Login to Azure.";
    public static readonly SkipDeployNoUpdates = "Skipping deployment: no updates found.";

    public static readonly PreScaffoldingBot = "Pre-scaffolding bot.";
    public static readonly ScaffoldingBot = "Scaffolding bot.";
    public static readonly SuccessfullyScaffoldedBot = "Successfully scaffolded bot.";

    public static readonly PreProvisioningBot = "Pre-provisioning bot.";
    public static readonly ProvisioningBot = "Provisioning bot.";
    public static readonly SuccessfullyProvisionedBot = "Successfully provisioned bot.";

    public static readonly PreDeployingBot = "Pre-deploying bot.";
    public static readonly DeployingBot = "Deploying bot.";
    public static readonly SuccessfullyDeployedBot = "Successfully deployed bot.";
}