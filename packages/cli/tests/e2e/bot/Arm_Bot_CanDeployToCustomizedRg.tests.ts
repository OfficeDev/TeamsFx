// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Di Lin <dilin@microsoft.com>
 */

import path from "path";

import { AadValidator, BotValidator } from "../../commonlib";
import {
  getSubscriptionId,
  getTestFolder,
  getUniqueAppName,
  cleanUp,
  readContextMultiEnv,
  createResourceGroup,
  deleteResourceGroupByName,
  setBotSkuNameToB1Bicep,
} from "../commonUtils";
import AppStudioLogin from "../../../src/commonlib/appStudioLogin";
import { environmentManager, isFeatureFlagEnabled } from "@microsoft/teamsfx-core";
import { FeatureFlagName } from "@microsoft/teamsfx-core/src/common/constants";
import { CliHelper } from "../../commonlib/cliHelper";
import { Capability, ResourceToDeploy } from "../../commonlib/constants";
import { customizeBicepFilesToCustomizedRg } from "../commonUtils";

describe("Deploy to customized resource group", function () {
  //  Only test when insider feature flag enabled
  if (!isFeatureFlagEnabled(FeatureFlagName.InsiderPreview, true)) {
    return;
  }

  const testFolder = getTestFolder();
  const subscription = getSubscriptionId();
  let appName: string, projectPath: string;

  beforeEach(async () => {
    appName = getUniqueAppName();
    projectPath = path.resolve(testFolder, appName);
  });

  afterEach(async () => {
    await cleanUp(appName, projectPath, true, false, false, true);
  });

  it(`bot project can deploy bot resource to customized resource group and successfully provision / deploy`, async function () {
    // Create new bot project
    await CliHelper.createProjectWithCapability(appName, testFolder, Capability.Bot);

    // Create empty resource group
    const customizedRgName = `${appName}-customized-rg`;
    await createResourceGroup(customizedRgName, "eastus");

    // Customize simple auth bicep files
    await customizeBicepFilesToCustomizedRg(
      customizedRgName,
      projectPath,
      `name: 'botProvision'`,
      `name: 'addTeamsFxBotConfiguration'`
    );

    // Provision
    setBotSkuNameToB1Bicep(projectPath, environmentManager.getDefaultEnvName());
    await CliHelper.setSubscription(subscription, projectPath);
    await CliHelper.provisionProject(projectPath);

    // deploy
    await CliHelper.deployProject(ResourceToDeploy.Bot, projectPath);

    // Assert
    {
      const context = await readContextMultiEnv(
        projectPath,
        environmentManager.getDefaultEnvName()
      );

      // Validate Aad App
      const aad = AadValidator.init(context, false, AppStudioLogin);
      await AadValidator.validate(aad);

      // Validate Bot
      const bot = BotValidator.init(context, true);
      await BotValidator.validateProvision(bot, true);
      await BotValidator.validateDeploy(bot);
    }

    await deleteResourceGroupByName(customizedRgName);
  });
});
