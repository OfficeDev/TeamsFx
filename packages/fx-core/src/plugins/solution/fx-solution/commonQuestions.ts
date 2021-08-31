/* eslint-disable @typescript-eslint/ban-types */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  ok,
  err,
  returnSystemError,
  returnUserError,
  FxError,
  Result,
  SolutionConfig,
  SolutionContext,
  AzureAccountProvider,
  SubscriptionInfo,
  QTreeNode,
  traverse,
  Inputs,
  UserInteraction,
  OptionItem,
} from "@microsoft/teamsfx-api";
import { GLOBAL_CONFIG, LOCATION, RESOURCE_GROUP_NAME, SolutionError } from "./constants";
import { v4 as uuidv4 } from "uuid";
import { ResourceManagementClient } from "@azure/arm-resources";
import { PluginDisplayName } from "../../../common/constants";
import { isMultiEnvEnabled } from "../../../common";
import {
  CoreQuestionNames,
  QuestionNewResourceGroupLocation,
  QuestionNewResourceGroupName,
  QuestionSelectResourceGroup,
} from "../../../core/question";
import { desensitize } from "../../../core/middleware/questionModel";
import { SOLUTION } from "../../resource/appstudio/constants";

export type AzureSubscription = {
  displayName: string;
  subscriptionId: string;
};

const DefaultResourceGroupLocation = "East US";
type ResourceGroupInfo = {
  createNewResourceGroup: boolean;
  name: string;
  location: string;
};

// TODO: use the emoji plus sign like Azure Functions extension
const newResourceGroupOption = "+ New resource group";

class CommonQuestions {
  resourceNameSuffix = "";
  resourceGroupName = "";
  tenantId = "";
  subscriptionId = "";
  // default to East US for now
  location = "East US";
  teamsAppTenantId = "";
}

/**
 * make sure subscription is correct
 *
 */
export async function checkSubscription(
  ctx: SolutionContext
): Promise<Result<SubscriptionInfo, FxError>> {
  if (ctx.azureAccountProvider === undefined) {
    return err(
      returnSystemError(
        new Error("azureAccountProvider is undefined"),
        "Solution",
        SolutionError.InternelError
      )
    );
  }
  const askSubRes = await ctx.azureAccountProvider!.getSelectedSubscription(true);
  return ok(askSubRes!);
}

async function getQuestionsForResourceGroup(
  defaultResourceGroupName: string,
  existingResourceGroupNameLocations: [string, string][],
  availableLocations: string[]
) {
  const selectResourceGroup = QuestionSelectResourceGroup;

  const staticOptions: OptionItem[] = [
    { id: newResourceGroupOption, label: newResourceGroupOption },
  ];
  selectResourceGroup.staticOptions = staticOptions.concat(
    existingResourceGroupNameLocations.map((item) => {
      return {
        id: item[0],
        label: item[0],
        description: item[1],
      };
    })
  );

  const node = new QTreeNode(selectResourceGroup);

  const inputNewResourceGroupName = QuestionNewResourceGroupName;
  inputNewResourceGroupName.default = defaultResourceGroupName;
  const newResourceGroupNameNode = new QTreeNode(inputNewResourceGroupName);
  newResourceGroupNameNode.condition = { equals: newResourceGroupOption };
  node.addChild(newResourceGroupNameNode);

  const selectLocation = QuestionNewResourceGroupLocation;
  // TODO: maybe lazily load locations
  selectLocation.staticOptions = availableLocations;
  const newResourceGroupLocationNode = new QTreeNode(selectLocation);
  newResourceGroupNameNode.addChild(newResourceGroupLocationNode);

  return node.trim();
}

/**
 * Ask user to create a new resource group or use an exsiting resource group
 */
export async function askResourceGroupInfo(
  ctx: SolutionContext,
  rmClient: ResourceManagementClient,
  inputs: Inputs,
  ui: UserInteraction,
  appName: string
): Promise<Result<ResourceGroupInfo, FxError>> {
  const defaultResourceGroupName = `${appName.replace(" ", "_")}-rg`;
  if (!isMultiEnvEnabled()) {
    return ok({
      createNewResourceGroup: true,
      name: defaultResourceGroupName,
      location: DefaultResourceGroupLocation,
    });
  }

  // TODO: support pagination
  let resourceGroupResults;
  try {
    resourceGroupResults = await rmClient.resourceGroups.list();
  } catch (error) {
    ctx.logProvider?.error(`Failed to list resource group: error '${error}'`);
    return err(
      returnSystemError(
        new Error(`Failed to list resource group`),
        "Solution",
        SolutionError.FailedToListResourceGroup
      )
    );
  }
  const resourceGroupNameLocations = resourceGroupResults
    .filter((item) => item.name)
    .map((item) => [item.name, item.location] as [string, string]);

  // TODO: call Azure API directly to list all locations because ARM SDK does not wrap this API.
  // And then filter by the 'resourceGroup' resource provider.
  // https://github.com/microsoft/vscode-azuretools/blob/cda6548af53a1c0f538a5ef7542c0eba1d5fa566/ui/src/wizard/LocationListStep.ts#L173
  const availableLocations = ["East US", "West US"];
  const node = await getQuestionsForResourceGroup(
    defaultResourceGroupName,
    resourceGroupNameLocations,
    availableLocations
  );
  if (node) {
    const res = await traverse(node, inputs, ui);
    if (res.isErr()) {
      ctx.logProvider?.debug(
        `[${PluginDisplayName.Solution}] failed to run question model for target resource group.`
      );
      return err(res.error);
    }

    const desensitized = desensitize(node, inputs);
    ctx.logProvider?.info(
      `[${
        PluginDisplayName.Solution
      }] success to run question model for resource group, answers:${JSON.stringify(desensitized)}`
    );
  }

  const resourceGroupName = inputs[CoreQuestionNames.TargetResourceGroupName];
  if (resourceGroupName === newResourceGroupOption) {
    return ok({
      name: inputs[CoreQuestionNames.NewResourceGroupName],
      location: inputs[CoreQuestionNames.NewResourceGroupLocation],
      createNewResourceGroup: true,
    });
  } else {
    const targetResourceGroupName = inputs[CoreQuestionNames.TargetResourceGroupName];
    if (typeof targetResourceGroupName !== "string") {
      return err(
        returnSystemError(
          new Error(`Failed to get user input for resource group info`),
          "Solution",
          SolutionError.FailedToListResourceGroup
        )
      );
    }

    const target = resourceGroupNameLocations.find((item) => item[0] == targetResourceGroupName);
    const location = target![1]; // location must exist because the user can only select from this list.
    return ok({
      createNewResourceGroup: false,
      name: targetResourceGroupName,
      location: location,
    });
  }
}

/**
 * Asks common questions and puts the answers in the global namespace of SolutionConfig
 *
 */
async function askCommonQuestions(
  ctx: SolutionContext,
  appName: string,
  config: SolutionConfig,
  azureAccountProvider?: AzureAccountProvider,
  appstudioTokenJson?: object
): Promise<Result<CommonQuestions, FxError>> {
  if (appstudioTokenJson === undefined) {
    return err(
      returnSystemError(
        new Error("Graph token json is undefined"),
        "Solution",
        SolutionError.NoAppStudioToken
      )
    );
  }

  const commonQuestions = new CommonQuestions();

  //1. check subscriptionId
  const subscriptionResult = await checkSubscription(ctx);
  if (subscriptionResult.isErr()) {
    return err(subscriptionResult.error);
  }
  const subscriptionId = subscriptionResult.value.subscriptionId;
  commonQuestions.subscriptionId = subscriptionId;
  commonQuestions.tenantId = subscriptionResult.value.tenantId;
  ctx.logProvider?.info(
    `[${PluginDisplayName.Solution}] askCommonQuestions, step 1 - check subscriptionId pass!`
  );

  // Note setSubscription here will change the token returned by getAccountCredentialAsync according to the subscription selected.
  // So getting azureToken needs to precede setSubscription.
  const azureToken = await azureAccountProvider?.getAccountCredentialAsync();
  if (azureToken === undefined) {
    return err(
      returnUserError(
        new Error("Login to Azure using the Azure Account extension"),
        "Solution",
        SolutionError.NotLoginToAzure
      )
    );
  }

  //2. check resource group
  const rmClient = new ResourceManagementClient(azureToken, subscriptionId);

  // Resource group info precedence are:
  //   1. env config (config.{envName}.json), for user customization
  //   2. publish profile (profile.{envName}.json), for reprovision
  //   3. asking user with a popup
  const resourceGroupNameFromEnvConfig = ctx.envInfo.config.azure.resourceGroupName;
  const resourceGroupNameFromProfile = ctx.envInfo.profile.get(SOLUTION).get(RESOURCE_GROUP_NAME);
  const resourceGroupLocationFromProfile = ctx.envInfo.profile.get(SOLUTION).get(LOCATION);
  let resourceGroupInfo: ResourceGroupInfo;
  if (resourceGroupNameFromEnvConfig) {
    const checkRes = await rmClient.resourceGroups.checkExistence(resourceGroupNameFromEnvConfig);
    if (checkRes.body) {
      resourceGroupInfo = {
        createNewResourceGroup: false,
        name: resourceGroupNameFromEnvConfig,
        location: DefaultResourceGroupLocation, // TODO: retrieve location using ARM SDK
      };
    } else {
      // Currently we do not support creating resource group by input config, so just throw an error.
      return err(
        returnUserError(
          new Error("Resource group does not exist"),
          "Solution",
          SolutionError.ResourceGroupNotFound
        )
      );
    }
  } else if (resourceGroupNameFromProfile && resourceGroupLocationFromProfile) {
    const checkRes = await rmClient.resourceGroups.checkExistence(resourceGroupNameFromProfile);
    if (checkRes.body) {
      resourceGroupInfo = {
        createNewResourceGroup: false,
        name: resourceGroupNameFromProfile,
        location: resourceGroupLocationFromProfile,
      };
    } else {
      resourceGroupInfo = {
        createNewResourceGroup: true,
        name: resourceGroupNameFromProfile,
        location: resourceGroupLocationFromProfile,
      };
    }
  } else if (ctx.answers && ctx.ui) {
    const resourceGroupInfoResult = await askResourceGroupInfo(
      ctx,
      rmClient,
      ctx.answers,
      ctx.ui,
      appName
    );
    if (resourceGroupInfoResult.isErr()) {
      return err(resourceGroupInfoResult.error);
    }
    resourceGroupInfo = resourceGroupInfoResult.value;
  } else {
    // fall back to default values when user interaction is not available
    resourceGroupInfo = {
      createNewResourceGroup: true,
      name: `${appName.replace(" ", "_")}-rg`,
      location: DefaultResourceGroupLocation,
    };
  }
  if (resourceGroupInfo.createNewResourceGroup) {
    const response = await rmClient.resourceGroups.createOrUpdate(resourceGroupInfo.name, {
      location: resourceGroupInfo.location,
    });

    if (response.name === undefined) {
      return err(
        returnSystemError(
          new Error(`Failed to create resource group ${resourceGroupInfo.name}`),
          "Solution",
          SolutionError.FailedToCreateResourceGroup
        )
      );
    }
    resourceGroupInfo.name = response.name;
    ctx.logProvider?.info(
      `[${PluginDisplayName.Solution}] askCommonQuestions - resource group:'${resourceGroupInfo.name}' created!`
    );
  }
  commonQuestions.resourceGroupName = resourceGroupInfo.name;
  commonQuestions.location = resourceGroupInfo.location;
  ctx.logProvider?.info(
    `[${PluginDisplayName.Solution}] askCommonQuestions, step 2 - check resource group pass!`
  );

  // teamsAppTenantId
  const teamsAppTenantId = (appstudioTokenJson as any).tid;
  if (
    teamsAppTenantId === undefined ||
    !(typeof teamsAppTenantId === "string") ||
    teamsAppTenantId.length === 0
  ) {
    return err(
      returnSystemError(
        new Error("Cannot find Teams app tenant id"),
        "Solution",
        SolutionError.NoTeamsAppTenantId
      )
    );
  } else {
    commonQuestions.teamsAppTenantId = teamsAppTenantId;
  }
  ctx.logProvider?.info(
    `[${PluginDisplayName.Solution}] askCommonQuestions, step 3 - check teamsAppTenantId pass!`
  );

  //resourceNameSuffix
  const resourceNameSuffix = config.get(GLOBAL_CONFIG)?.getString("resourceNameSuffix");
  if (!resourceNameSuffix) commonQuestions.resourceNameSuffix = uuidv4().substr(0, 6);
  else commonQuestions.resourceNameSuffix = resourceNameSuffix;
  ctx.logProvider?.info(
    `[${PluginDisplayName.Solution}] askCommonQuestions, step 4 - check resourceNameSuffix pass!`
  );

  ctx.logProvider?.info(
    `[${PluginDisplayName.Solution}] askCommonQuestions, step 5 - check tenantId pass!`
  );

  return ok(commonQuestions);
}

/**
 * Asks for userinput and fills the answers in global config.
 *
 * @param config reference to solution config
 * @param dialog communication channel to Core Module
 */
export async function fillInCommonQuestions(
  ctx: SolutionContext,
  appName: string,
  config: SolutionConfig,
  azureAccountProvider?: AzureAccountProvider,
  // eslint-disable-next-line @typescript-eslint/ban-types
  appStudioJson?: object
): Promise<Result<SolutionConfig, FxError>> {
  const result = await askCommonQuestions(
    ctx,
    appName,
    config,
    azureAccountProvider,
    appStudioJson
  );
  if (result.isOk()) {
    // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
    const globalConfig = config.get(GLOBAL_CONFIG)!;
    result.map((commonQuestions) => {
      for (const [k, v] of Object.entries(commonQuestions)) {
        globalConfig.set(k, v);
      }
    });
    return ok(config);
  }
  return result.map((_) => config);
}
