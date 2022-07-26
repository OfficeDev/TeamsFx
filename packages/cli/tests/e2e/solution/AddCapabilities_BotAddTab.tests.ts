// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Di Lin <dilin@microsoft.com>
 */

import path from "path";
import { environmentManager, isPreviewFeaturesEnabled } from "@microsoft/teamsfx-core";
import {
  getSubscriptionId,
  getTestFolder,
  getUniqueAppName,
  cleanUp,
  setBotSkuNameToB1Bicep,
  setSimpleAuthSkuNameToB1Bicep,
  validateTabAndBotProjectProvision,
} from "../commonUtils";
import "mocha";
import { CliHelper } from "../../commonlib/cliHelper";
import { Capability } from "../../commonlib/constants";
import mockedEnv, { RestoreFn } from "mocked-env";
describe("Add capabilities", function () {
  const testFolder = getTestFolder();
  const subscription = getSubscriptionId();
  let appName: string | undefined;
  let projectPath: string | undefined;
  let mockedEnvRestore: RestoreFn | undefined;
  const env = environmentManager.getDefaultEnvName();
  afterEach(async () => {
    if (appName && projectPath) {
      await cleanUp(appName, projectPath, true, true, false);
    }
    if (mockedEnvRestore) {
      mockedEnvRestore();
    }
  });
  for (const v3flag of ["false", "true"]) {
    it("bot project can add tab capability and provision", async () => {
      mockedEnvRestore = mockedEnv({ TEAMSFX_APIV3: v3flag });
      appName = getUniqueAppName();
      projectPath = path.resolve(testFolder, appName);

      // Arrange
      await CliHelper.createProjectWithCapability(appName, testFolder, Capability.Bot);

      // Act
      if (isPreviewFeaturesEnabled()) {
        await CliHelper.addCapabilityToProject(projectPath, Capability.SSOTab);
      } else {
        await CliHelper.addCapabilityToProject(projectPath, Capability.Tab);
      }

      await setSimpleAuthSkuNameToB1Bicep(projectPath, env);
      await setBotSkuNameToB1Bicep(projectPath, env);
      await CliHelper.setSubscription(subscription, projectPath);
      await CliHelper.provisionProject(projectPath);

      // Assert
      await validateTabAndBotProjectProvision(projectPath, env);
    });

    it("message extension project can add tab capability and provision", async () => {
      mockedEnvRestore = mockedEnv({ TEAMSFX_APIV3: v3flag });
      appName = getUniqueAppName();
      projectPath = path.resolve(testFolder, appName);

      // Arrange
      await CliHelper.createProjectWithCapability(appName, testFolder, Capability.MessageExtension);

      // Act
      if (isPreviewFeaturesEnabled()) {
        await CliHelper.addCapabilityToProject(projectPath, Capability.SSOTab);
      } else {
        await CliHelper.addCapabilityToProject(projectPath, Capability.Tab);
      }

      await setSimpleAuthSkuNameToB1Bicep(projectPath, env);
      await setBotSkuNameToB1Bicep(projectPath, env);
      await CliHelper.setSubscription(subscription, projectPath);
      await CliHelper.provisionProject(projectPath);

      // Assert
      await validateTabAndBotProjectProvision(projectPath, env);
    });
  }
});
