// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Di Lin <dilin@microsoft.com>
 */

import path from "path";
import {
  AadValidator,
  BotValidator,
  FrontendValidator,
  SimpleAuthValidator,
} from "../../commonlib";
import { environmentManager, isFeatureFlagEnabled } from "@microsoft/teamsfx-core";
import {
  execAsync,
  execAsyncWithRetry,
  getSubscriptionId,
  getTestFolder,
  getUniqueAppName,
  cleanUp,
  setBotSkuNameToB1Bicep,
  readContextMultiEnv,
  setSimpleAuthSkuNameToB1Bicep,
} from "../commonUtils";
import AppStudioLogin from "../../../src/commonlib/appStudioLogin";
import { FeatureFlagName } from "@microsoft/teamsfx-core/src/common/constants";
import "mocha";
import { Capability, getWebappServicePlan, Resource } from "../../commonlib/utilities";
import * as fs from "fs-extra";
import MockAzureAccountProvider from "../../../src/commonlib/azureLoginUserPassword";
import { expect } from "chai";

describe("Add capabilities", function () {
  //  Only test when insider feature flag enabled
  if (!isFeatureFlagEnabled(FeatureFlagName.InsiderPreview, true)) {
    return;
  }

  const testFolder = getTestFolder();
  const subscription = getSubscriptionId();
  let appName: string, projectPath: string;

  before(async () => {
    appName = getUniqueAppName();
    projectPath = path.resolve(testFolder, appName);
  });

  after(async () => {
    await cleanUp(appName, projectPath, true, false, false, true);
  });

  it("tab project can add bot capability and provision", async () => {
    // Arrange
    await createProjectWithCapability(appName, testFolder, Capability.Tab);

    // Act
    await addCapabilityToProject(projectPath, Capability.Bot);

    await setSimpleAuthSkuNameToB1Bicep(projectPath, environmentManager.getDefaultEnvName());
    await setBotSkuNameToB1Bicep(projectPath, environmentManager.getDefaultEnvName());
    await SetSubAndProvisionProject(subscription, projectPath);

    // Assert
    await validateTabAndBotProjectProvision(projectPath);
  });

  it("tab project can add messaging extension capability and provision", async () => {
    // Arrange
    await createProjectWithCapability(appName, testFolder, Capability.Tab);

    // Act
    await addCapabilityToProject(projectPath, Capability.MessagingExtension);

    await setSimpleAuthSkuNameToB1Bicep(projectPath, environmentManager.getDefaultEnvName());
    await setBotSkuNameToB1Bicep(projectPath, environmentManager.getDefaultEnvName());
    await SetSubAndProvisionProject(subscription, projectPath);

    // Assert
    await validateTabAndBotProjectProvision(projectPath);
  });

  it("bot project can add tab capability and provision", async () => {
    // Arrange
    await createProjectWithCapability(appName, testFolder, Capability.Bot);

    // Act
    await addCapabilityToProject(projectPath, Capability.Tab);

    await setSimpleAuthSkuNameToB1Bicep(projectPath, environmentManager.getDefaultEnvName());
    await setBotSkuNameToB1Bicep(projectPath, environmentManager.getDefaultEnvName());
    await SetSubAndProvisionProject(subscription, projectPath);

    // Assert
    await validateTabAndBotProjectProvision(projectPath);
  });

  it("messaging extnsion project can add tab capability and provision", async () => {
    // Arrange
    await createProjectWithCapability(appName, testFolder, Capability.MessagingExtension);

    // Act
    await addCapabilityToProject(projectPath, Capability.Tab);

    await setSimpleAuthSkuNameToB1Bicep(projectPath, environmentManager.getDefaultEnvName());
    await setBotSkuNameToB1Bicep(projectPath, environmentManager.getDefaultEnvName());
    await SetSubAndProvisionProject(subscription, projectPath);

    // Assert
    await validateTabAndBotProjectProvision(projectPath);
  });

  async function validateTabAndBotProjectProvision(projectPath: string) {
    const context = await readContextMultiEnv(projectPath, environmentManager.getDefaultEnvName());

    // Validate Aad App
    const aad = AadValidator.init(context, false, AppStudioLogin);
    await AadValidator.validate(aad);

    // Validate Simple Auth
    const simpleAuth = SimpleAuthValidator.init(context);
    await SimpleAuthValidator.validate(simpleAuth, aad);

    // Validate Tab Frontend
    const frontend = FrontendValidator.init(context, true);
    await FrontendValidator.validateProvision(frontend);

    // Validate Bot Provision
    const bot = BotValidator.init(context, true);
    await BotValidator.validateProvision(bot, true);
  }
});

describe("User can customize Bicep files", function () {
  //  Only test when insider feature flag enabled
  if (!isFeatureFlagEnabled(FeatureFlagName.InsiderPreview, true)) {
    return;
  }

  const testFolder = getTestFolder();
  const subscription = getSubscriptionId();
  let appName: string, projectPath: string;

  before(async () => {
    appName = getUniqueAppName();
    projectPath = path.resolve(testFolder, appName);
  });

  after(async () => {
    await cleanUp(appName, projectPath, true, false, false, true);
  });

  it("user customized Bicep file is used when provision", async () => {
    // Arrange
    await createProjectWithCapability(appName, testFolder, Capability.Tab);

    // Act
    await setSimpleAuthSkuNameToB1Bicep(projectPath, environmentManager.getDefaultEnvName());
    const customizedServicePlans: string[] = await customizeBicepFile(projectPath);
    await SetSubAndProvisionProject(subscription, projectPath);

    const resourceGroup: string = await getRGAfterProvision(projectPath);
    chai.assert.exists(resourceGroup);
    expect(resourceGroup).to.be.a("string");

    // Assert
    customizedServicePlans.forEach(async (servicePlanName) => {
      await validateServicePlan(servicePlanName, resourceGroup);
    });
  });

  it("Regenerate Bicep will not affect user's customized Bicep code", async () => {
    // Arrange
    await createProjectWithCapability(appName, testFolder, Capability.Tab);

    // Act
    await setSimpleAuthSkuNameToB1Bicep(projectPath, environmentManager.getDefaultEnvName());
    const customizedServicePlans: string[] = await customizeBicepFile(projectPath);

    // Add capability and cloud resource
    await addCapabilityToProject(projectPath, Capability.Bot);
    await setBotSkuNameToB1Bicep(projectPath, environmentManager.getDefaultEnvName());
    await addResourceToProject(projectPath, Resource.AzureFunction);

    await SetSubAndProvisionProject(subscription, projectPath);

    const resourceGroup = await getRGAfterProvision(projectPath);
    chai.assert.exists(resourceGroup);
    expect(resourceGroup).to.be.a("string");

    // Assert
    customizedServicePlans.forEach(async (servicePlanName) => {
      await validateServicePlan(servicePlanName, resourceGroup);
    });
  });

  async function getRGAfterProvision(projectPath: string): Promise<string | undefined> {
    const context = await readContextMultiEnv(projectPath, environmentManager.getDefaultEnvName());
    if (context["solution"] && context["solution"]["resourceGroupName"]) {
      return context["solution"]["resourceGroupName"];
    }
    return undefined;
  }

  async function customizeBicepFile(projectPath: string): Promise<string[]> {
    const newServerFarms: string[] = [];
    const bicepFileFolder = path.join(projectPath, "templates", "azure");

    const simpleAuthTestServerFarm = "simpleAuth_testResource";
    await fs.appendFile(
      path.join(bicepFileFolder, "provision", "simpleAuth.bicep"),
      `
resource customizedServerFarms 'Microsoft.Web/serverfarms@2021-02-01' = {
  name: '${simpleAuthTestServerFarm}'
  location: resourceGroup().location
  sku: {
    name: 'B1'
  }
  kind: 'app'
}
`
    );
    newServerFarms.push(simpleAuthTestServerFarm);

    const provisionTestServerFarm = "provision_testResource";
    await fs.appendFile(
      path.join(bicepFileFolder, "provision.bicep"),
      `
resource customizedServerFarms 'Microsoft.Web/serverfarms@2021-02-01' = {
  name: '${provisionTestServerFarm}'
  location: resourceGroup().location
  sku: {
    name: 'B1'
  }
  kind: 'app'
}
`
    );
    newServerFarms.push(provisionTestServerFarm);

    const configTestServerFarm = "config_testResource";
    await fs.appendFile(
      path.join(bicepFileFolder, "config.bicep"),
      `
resource customizedServerFarms 'Microsoft.Web/serverfarms@2021-02-01' = {
  name: '${configTestServerFarm}'
  location: resourceGroup().location
  sku: {
    name: 'B1'
  }
  kind: 'app'
}
`
    );
    newServerFarms.push(configTestServerFarm);

    const mainTestServerFarm = "main_testResource";
    await fs.appendFile(
      path.join(bicepFileFolder, "main.bicep"),
      `
resource customizedServerFarms 'Microsoft.Web/serverfarms@2021-02-01' = {
  name: '${mainTestServerFarm}'
  location: resourceGroup().location
  sku: {
    name: 'B1'
  }
  kind: 'app'
}
`
    );
    newServerFarms.push(mainTestServerFarm);

    return newServerFarms;
  }

  async function validateServicePlan(servicePlanName: string, resourceGroup: string) {
    console.log(`Start to validate server farm ${servicePlanName}.`);

    const tokenProvider = MockAzureAccountProvider;
    const tokenCredential = await tokenProvider.getAccountCredentialAsync();
    const token = (await tokenCredential?.getToken())?.accessToken;

    const serivcePlanResponse = await getWebappServicePlan(
      subscription,
      resourceGroup,
      servicePlanName,
      token as string
    );
    chai.assert(serivcePlanResponse, "B1");
  }
});

async function createProjectWithCapability(
  appName: string,
  testFolder: string,
  capability: string
) {
  await execAsync(
    `teamsfx new --interactive false --app-name ${appName} --capabilities ${capability} `,
    {
      cwd: testFolder,
      env: process.env,
      timeout: 0,
    }
  );
  console.log(
    `[Successfully] scaffold project to ${path.resolve(
      testFolder,
      appName
    )} with capability ${capability}`
  );
}

async function addCapabilityToProject(projectPath: string, capabilityToAdd: string) {
  await execAsync(`teamsfx capability add ${capabilityToAdd}`, {
    cwd: projectPath,
    env: process.env,
    timeout: 0,
  });
  console.log(`[Successfully] add capability ${capabilityToAdd} to ${projectPath}`);
}

async function addResourceToProject(projectPath: string, resourceToAdd: string) {
  await execAsync(`teamsfx resource add ${resourceToAdd}`, {
    cwd: projectPath,
    env: process.env,
    timeout: 0,
  });
  console.log(`[Successfully] add resource ${resourceToAdd} to ${projectPath}`);
}

async function SetSubAndProvisionProject(subscription: string, projectPath: string) {
  // set subscription
  await execAsync(`teamsfx account set --subscription ${subscription}`, {
    cwd: projectPath,
    env: process.env,
    timeout: 0,
  });
  console.log(`[Successfully] set subscription for ${projectPath}`);

  // provision
  await execAsyncWithRetry(`teamsfx provision`, {
    cwd: projectPath,
    env: process.env,
    timeout: 0,
  });
  console.log(`[Successfully] provision for ${projectPath}`);
}
