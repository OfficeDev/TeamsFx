// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import path from "path";
import * as chai from "chai";

import {
  AadValidator,
  AppStudioValidator,
  FrontendValidator,
  FunctionValidator,
  SimpleAuthValidator,
} from "../../commonlib";

import {
  execAsync,
  execAsyncWithRetry,
  getSubscriptionId,
  getTestFolder,
  getUniqueAppName,
  setSimpleAuthSkuNameToB1,
  setBotSkuNameToB1,
  cleanUp,
  TEN_MEGA_BYTE,
} from "../commonUtils";
import AppStudioLogin from "../../../src/commonlib/appStudioLogin";

describe("Azure App Happy Path", function () {
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const subscription = getSubscriptionId();
  const projectPath = path.resolve(testFolder, appName);

  it(`Tab + Bot (Create New) + Function + SQL + Apim`, async function () {
    // new a project ( tab + function + sql )
    await execAsync(
      `teamsfx new --interactive false --app-name ${appName} --capabilities tab --azure-resources function sql`,
      {
        cwd: testFolder,
        env: process.env,
        timeout: 0,
        maxBuffer: TEN_MEGA_BYTE,
      }
    );
    console.log(`[Successfully] scaffold to ${projectPath}`);

    await setSimpleAuthSkuNameToB1(projectPath);

    // capability add bot
    await execAsync(`teamsfx capability add bot`, {
      cwd: projectPath,
      env: process.env,
      timeout: 0,
      maxBuffer: TEN_MEGA_BYTE,
    });

    await setBotSkuNameToB1(projectPath);

    // set subscription
    await execAsync(`teamsfx account set --subscription ${subscription}`, {
      cwd: projectPath,
      env: process.env,
      timeout: 0,
      maxBuffer: TEN_MEGA_BYTE,
    });

    // resource add apim
    await execAsync(`teamsfx resource add azure-apim --function-name testApim`, {
      cwd: projectPath,
      env: process.env,
      timeout: 0,
      maxBuffer: TEN_MEGA_BYTE,
    });

    {
      /// TODO: add check for scaffold
    }

    // provision
    await execAsyncWithRetry(
      `teamsfx provision --sql-admin-name Abc123321 --sql-password Cab232332`,
      {
        cwd: projectPath,
        env: process.env,
        timeout: 0,
        maxBuffer: TEN_MEGA_BYTE,
      }
    );

    {
      // Validate provision
      // Get context
      const context = await fs.readJSON(`${projectPath}/.fx/env.default.json`);

      // Validate Aad App
      const aad = AadValidator.init(context, false, AppStudioLogin);
      await AadValidator.validate(aad);

      // Validate Simple Auth
      const simpleAuth = SimpleAuthValidator.init(context);
      await SimpleAuthValidator.validate(simpleAuth, aad);

      // Validate Function App
      const func = FunctionValidator.init(context);
      await FunctionValidator.validateProvision(func);

      // Validate Tab Frontend
      const frontend = FrontendValidator.init(context);
      await FrontendValidator.validateProvision(frontend);
    }

    // deploy
    await execAsyncWithRetry(
      `teamsfx deploy --open-api-document openapi/openapi.json --api-prefix qwed --api-version v1`,
      {
        cwd: projectPath,
        env: process.env,
        timeout: 0,
        maxBuffer: TEN_MEGA_BYTE,
      }
    );

    {
      /// TODO: add check for deploy
    }

    // validate the manifest
    const validationResult = await execAsyncWithRetry(`teamsfx validate`, {
      cwd: projectPath,
      env: process.env,
      timeout: 0,
    });

    {
      chai.assert.isEmpty(validationResult.stderr);
    }

    // build
    await execAsyncWithRetry(`teamsfx build`, {
      cwd: projectPath,
      env: process.env,
      timeout: 0,
    });

    {
      // Validate built package
      const file = `${projectPath}/.fx/appPackage.zip`;
      chai.assert.isTrue(await fs.pathExists(file));
    }

    // publish
    await execAsyncWithRetry(`teamsfx publish`, {
      cwd: projectPath,
      env: process.env,
      timeout: 0,
    });

    {
      // Validate publish result
      const context = await fs.readJSON(`${projectPath}/.fx/env.default.json`);
      const aad = AadValidator.init(context, false, AppStudioLogin);
      const appId = aad.clientId;

      AppStudioValidator.init();
      await AppStudioValidator.validatePublish(appId);
    }
  });

  after(async () => {
    // clean up
    await cleanUp(appName, projectPath, true, true, true);
  });
});
