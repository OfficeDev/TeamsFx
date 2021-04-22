// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import path from "path";
import { expect } from "chai";

import { AadValidator, deleteAadApp, MockAzureAccountProvider } from "fx-api";

import { execAsync, getTestFolder, getUniqueAppName } from "./commonUtils";

describe("Provision", function() {
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);

  it(`Provision Resource: Update Permission for AAD - Test Plan Id 9729543`, async function() {
    // new a project
    const newResult = await execAsync(`teamsfx new --app-name ${appName} --interactive false --verbose false`, {
      cwd: testFolder,
      env: process.env,
      timeout: 0
    });
    expect(newResult.stdout).to.eq("");
    expect(newResult.stderr).to.eq("");

    {
      // set fx-resource-simple-auth.skuName as B1
      const context = await fs.readJSON(`${projectPath}/.fx/env.default.json`);
      context["fx-resource-simple-auth"]["skuName"] = "B1";
      await fs.writeJSON(`${projectPath}/.fx/env.default.json`, context, { spaces: 4 });
    }

    {
      // update permission
      const permission = "[{\"resource\":\"Microsoft Graph\",\"scopes\": [\"User.Read\",\"User.Read.All\"]}]";
      await fs.writeJSON(`${projectPath}/permission.json`, permission, { spaces: 4 });
    }

    // provision
    const provisionResult = await execAsync(
      `teamsfx provision --subscription 1756abc0-3554-4341-8d6a-46674962ea19 --verbose false`,
      {
        cwd: projectPath,
        env: process.env,
        timeout: 0
      }
    );
    expect(provisionResult.stdout).to.eq("");
    expect(provisionResult.stderr).to.eq("");

    // Get context
    const expectedPermission = "[{\"resourceAppId\":\"00000003-0000-0000-c000-000000000000\",\"resourceAccess\": [{\"id\": \"e1fe6dd8-ba31-4d61-89e7-88639da4683d\",\"type\": \"Scope\"},{\"id\": \"a154be20-db9c-4678-8ab7-66f6cc099a59\",\"type\": \"Scope\"}]}]";
    const context = await fs.readJSON(`${projectPath}/.fx/env.default.json`);

    // Validate Aad App
    const aad = AadValidator.init(context);
    await AadValidator.validate(aad, expectedPermission);
  });

  this.afterAll(async () => {
    // delete aad app
    const context = await fs.readJSON(`${projectPath}/.fx/env.default.json`);
    await deleteAadApp(context);

    // remove resouce
    await MockAzureAccountProvider.getInstance().deleteResourceGroup(`${appName}-rg`);

    // remove project
    await fs.remove(projectPath);
  });
});
