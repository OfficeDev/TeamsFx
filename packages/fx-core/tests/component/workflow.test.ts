// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  FuncQuestion,
  InputsWithProjectPath,
  ok,
  Platform,
  ResourceContextV3,
  Void,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import * as os from "os";
import * as path from "path";
import sinon from "sinon";
import { setTools } from "../../src/core/globalVars";
import * as templateAction from "../../src/common/template-utils/templatesActions";
import "../../src/component/core";
import "../../src/component/feature/bot";
import "../../src/component/feature/sql";
import "../../src/component/resource/botService";
import { createContextV3 } from "../../src/component/utils";
import { runActionByName } from "../../src/component/workflow";
import { MockTools, randomAppName } from "../core/utils";
import * as provisionV3 from "../../src/plugins/solution/fx-solution/v3/provision";
import { AppStudioClient } from "../../src/plugins/resource/appstudio/appStudio";
import * as clientFactory from "../../src/plugins/resource/bot/clientFactory";
import { AADRegistration } from "../../src/plugins/resource/bot/aadRegistration";
import { TestHelper } from "../plugins/resource/frontend/helper";
import arm from "../../src/plugins/solution/fx-solution/arm";
import { FrontendDeployment } from "../../src/plugins/resource/frontend/ops/deploy";
import { newEnvInfoV3 } from "../../src/core/environment";
import { Utils } from "../../src/plugins/resource/spfx/utils/utils";
import { YoChecker } from "../../src/plugins/resource/spfx/depsChecker/yoChecker";
import { GeneratorChecker } from "../../src/plugins/resource/spfx/depsChecker/generatorChecker";
import { cpUtils } from "../../src/plugins/solution/fx-solution/utils/depsChecker/cpUtils";
import * as uuid from "uuid";
import * as aadManifest from "../../src/core/generateAadManifestTemplate";
import {
  SPFXQuestionNames,
  versionCheckQuestion,
} from "../../src/plugins/resource/spfx/utils/questions";
import { DefaultManifestProvider } from "../../src/component/resource/appManifest/manifestProvider";
import { ComponentNames } from "../../src/component/constants";
import { AzureSolutionQuestionNames } from "../../src";
import { FunctionScaffold } from "../../src/plugins/resource/function/ops/scaffold";
import { TeamsfxCore } from "../../src/component/core";
import Container from "typedi";
import { AzureStorageResource } from "../../src/component/resource/azureStorage";
describe("Workflow test for v3", () => {
  const sandbox = sinon.createSandbox();
  const tools = new MockTools();
  setTools(tools);
  const appName = `unittest${randomAppName()}`;
  const projectPath = path.join(os.homedir(), "TeamsApps", appName);
  const context = createContextV3();
  const fx = Container.get<TeamsfxCore>("fx");
  beforeEach(() => {
    sandbox.stub(tools.ui, "showMessage").resolves(ok("Confirm"));
  });

  afterEach(() => {
    sandbox.restore();
  });

  after(async () => {
    await fs.remove(projectPath);
  });

  it("fx.init", async () => {
    const inputs: InputsWithProjectPath = {
      projectPath: projectPath,
      platform: Platform.VSCode,
      "app-name": appName,
      folder: path.join(os.homedir(), "TeamsApps"),
    };
    const res = await fx.init(context, inputs);
    assert.isTrue(res.isOk());
    assert.equal(context.projectSetting!.appName, appName);
    assert.deepEqual(context.projectSetting.components, []);
  });

  it("teams-bot.add", async () => {
    const inputs: InputsWithProjectPath = {
      projectPath: projectPath,
      platform: Platform.VSCode,
      [AzureSolutionQuestionNames.Features]: "Bot",
      language: "typescript",
    };
    sandbox.stub(templateAction, "scaffoldFromTemplates").resolves();
    const res = await runActionByName("teams-bot.add", context, inputs);
    if (res.isErr()) {
      console.log(res.error);
    }
    assert.isTrue(res.isOk());
  });
  it("spfx-tab.add", async () => {
    sandbox.stub(Utils, "configure");
    sandbox.stub(fs, "stat").resolves();
    sandbox.stub(YoChecker.prototype, "isInstalled").resolves(true);
    sandbox.stub(GeneratorChecker.prototype, "isInstalled").resolves(true);
    sandbox.stub(cpUtils, "executeCommand").resolves("succeed");
    const manifestId = uuid.v4();
    sandbox.stub(fs, "readFile").resolves(new Buffer(`{"id": "${manifestId}"}`));
    sandbox.stub(fs, "writeFile").resolves();
    sandbox.stub(fs, "rename").resolves();
    sandbox.stub(fs, "copyFile").resolves();
    sandbox.stub(versionCheckQuestion as FuncQuestion, "func").resolves(undefined);
    sinon.stub(DefaultManifestProvider.prototype, "updateCapability").resolves(ok(Void));

    const inputs: InputsWithProjectPath = {
      projectPath: projectPath,
      platform: Platform.CLI,
      language: "typescript",
      [SPFXQuestionNames.webpart_name]: "hello",
      [SPFXQuestionNames.framework_type]: "none",
    };
    const res = await runActionByName("spfx-tab.add", context, inputs);
    if (res.isErr()) {
      console.log(res.error);
    }
    assert.isTrue(res.isOk());
  });
  it("sql.add", async () => {
    sandbox.stub(FunctionScaffold, "scaffoldFunction").resolves();
    const inputs: InputsWithProjectPath = {
      projectPath: projectPath,
      platform: Platform.VSCode,
      ["function-name"]: "getUserProfile",
    };
    const res = await runActionByName("sql.add", context, inputs);
    if (res.isErr()) {
      console.log(res.error);
    }
    assert.isTrue(res.isOk());
  });
  it("sso.add", async () => {
    const inputs: InputsWithProjectPath = {
      projectPath: projectPath,
      platform: Platform.VSCode,
    };
    const res = await runActionByName("sso.add", context, inputs);
    if (res.isErr()) {
      console.log(res.error);
    }
    assert.isTrue(res.isOk());
  });
  it("apim-feature.add", async () => {
    sandbox.stub(FunctionScaffold, "scaffoldFunction").resolves();
    const inputs: InputsWithProjectPath = {
      projectPath: projectPath,
      platform: Platform.VSCode,
      ["function-name"]: "getUserProfile",
    };
    const res = await runActionByName("apim-feature.add", context, inputs);
    if (res.isErr()) {
      console.log(res.error);
    }
    assert.isTrue(res.isOk());
  });
  it("fx.provision", async () => {
    sandbox.stub(templateAction, "scaffoldFromTemplates").resolves();
    sandbox.stub(tools.tokenProvider.m365TokenProvider, "getAccessToken").resolves(ok("fakeToken"));
    sandbox
      .stub(tools.tokenProvider.m365TokenProvider, "getJsonObject")
      .resolves(ok({ tid: "mockTid" }));
    sandbox
      .stub(tools.tokenProvider.azureAccountProvider, "getAccountCredentialAsync")
      .resolves(TestHelper.fakeCredential);
    sandbox
      .stub(provisionV3, "fillInAzureConfigs")
      .resolves(ok({ hasSwitchedSubscription: false }));
    sandbox.stub(provisionV3, "askForProvisionConsent").resolves(ok(Void));
    sandbox.stub(AppStudioClient, "getApp").onFirstCall().throws({}).onSecondCall().resolves({});
    sandbox.stub(AppStudioClient, "importApp").resolves({ teamsAppId: "mockTeamsAppId" });
    sandbox.stub(clientFactory, "createResourceProviderClient").resolves({});
    sandbox.stub(clientFactory, "ensureResourceProvider").resolves();
    sandbox.stub(AADRegistration, "registerAADAppAndGetSecretByGraph").resolves({
      clientId: "00000000-0000-0000-0000-000000000000",
      clientSecret: "mockClientSecret",
      objectId: "00000000-0000-0000-0000-000000000000",
    });
    sandbox.stub(arm, "deployArmTemplates").resolves(ok(undefined));
    const appName = `unittest${randomAppName()}`;
    const inputs: InputsWithProjectPath = {
      projectPath: projectPath,
      platform: Platform.VSCode,
      [AzureSolutionQuestionNames.Features]: "Bot",
      language: "typescript",
      "app-name": appName,
      folder: path.join(os.homedir(), "TeamsApps"),
    };
    const initRes = await fx.init(context, inputs);
    if (initRes.isErr()) {
      console.log(initRes.error);
    }
    assert.isTrue(initRes.isOk());

    const addBotRes = await runActionByName("teams-bot.add", context, inputs);
    if (addBotRes.isErr()) {
      console.log(addBotRes.error);
    }
    assert.isTrue(addBotRes.isOk());
    context.envInfo = newEnvInfoV3();
    context.tokenProvider = tools.tokenProvider;
    context.envInfo.state = {
      solution: {
        provisionSucceeded: true,
        needCreateResourceGroup: false,
        resourceGroupName: "mockRG",
        location: "eastasia",
        resourceNameSuffix: "3bf854123",
        teamsAppTenantId: "mockTid",
        subscriptionId: "mockSid",
        subscriptionName: "mockSname",
        tenantId: "mockAzureTid",
      },
      "azure-web-app": {
        sku: "F1",
        appName: "testwebApp",
        domain: "testwebApp.azurewebsites.net",
        appServicePlanName: "testwebAppPlan",
        resourceId:
          "/subscriptions/mockSid/resourceGroups/jay-texas/providers/Microsoft.Web/sites/testwebApp",
        endpoint: "https://testwebApp.azurewebsites.net",
      },
      [ComponentNames.BotService]: {
        botId: "00000000-0000-0000-0000-000000000000",
      },
      [ComponentNames.AadApp]: {
        clientId: "00000000-0000-0000-0000-000000000000",
      },
    };
    const provisionRes = await runActionByName("fx.provision", context, inputs);
    if (provisionRes.isErr()) {
      console.log(provisionRes.error);
    }
    assert.isTrue(provisionRes.isOk());
  });

  it("fx.provision after switching subscription", async () => {
    sandbox.stub(templateAction, "scaffoldFromTemplates").resolves();
    sandbox.stub(tools.tokenProvider.m365TokenProvider, "getAccessToken").resolves(ok("fakeToken"));
    sandbox
      .stub(tools.tokenProvider.m365TokenProvider, "getJsonObject")
      .resolves(ok({ tid: "mockTid" }));
    sandbox
      .stub(tools.tokenProvider.azureAccountProvider, "getAccountCredentialAsync")
      .resolves(TestHelper.fakeCredential);
    sandbox.stub(provisionV3, "fillInAzureConfigs").resolves(ok({ hasSwitchedSubscription: true }));
    sandbox.stub(AppStudioClient, "getApp").onFirstCall().throws({}).onSecondCall().resolves({});
    sandbox.stub(AppStudioClient, "importApp").resolves({ teamsAppId: "mockTeamsAppId" });
    sandbox.stub(clientFactory, "createResourceProviderClient").resolves({});
    sandbox.stub(clientFactory, "ensureResourceProvider").resolves();
    sandbox.stub(AADRegistration, "registerAADAppAndGetSecretByGraph").resolves({
      clientId: "mockClientId",
      clientSecret: "mockClientSecret",
      objectId: "mockObjectId",
    });
    sandbox.stub(arm, "deployArmTemplates").resolves(ok(undefined));
    const appName = `unittest${randomAppName()}`;
    const inputs: InputsWithProjectPath = {
      projectPath: projectPath,
      platform: Platform.VSCode,
      [AzureSolutionQuestionNames.Features]: "Bot",
      language: "typescript",
      "app-name": appName,
      folder: path.join(os.homedir(), "TeamsApps"),
    };
    const initRes = await fx.init(context, inputs);
    if (initRes.isErr()) {
      console.log(initRes.error);
    }
    assert.isTrue(initRes.isOk());

    const addBotRes = await runActionByName("teams-bot.add", context, inputs);
    if (addBotRes.isErr()) {
      console.log(addBotRes.error);
    }
    assert.isTrue(addBotRes.isOk());
    context.envInfo = newEnvInfoV3();
    context.tokenProvider = tools.tokenProvider;
    context.envInfo.state = {
      solution: {
        provisionSucceeded: true,
        needCreateResourceGroup: false,
        resourceGroupName: "mockRG",
        location: "eastasia",
        resourceNameSuffix: "3bf854123",
        teamsAppTenantId: "mockTid",
        subscriptionId: "mockSid",
        subscriptionName: "mockSname",
        tenantId: "mockAzureTid",
      },
      "azure-web-app": {
        sku: "F1",
        appName: "testwebApp",
        domain: "testwebApp.azurewebsites.net",
        appServicePlanName: "testwebAppPlan",
        resourceId:
          "/subscriptions/mockSid/resourceGroups/jay-texas/providers/Microsoft.Web/sites/testwebApp",
        endpoint: "https://testwebApp.azurewebsites.net",
      },
    };
    const provisionRes = await runActionByName("fx.provision", context, inputs);
    if (provisionRes.isErr()) {
      console.log(provisionRes.error);
    }
    assert.isTrue(provisionRes.isOk());
  });

  it("fx.provision local debug after switching m365 tenant", async () => {
    sandbox.stub(templateAction, "scaffoldFromTemplates").resolves();
    sandbox.stub(tools.tokenProvider.m365TokenProvider, "getAccessToken").resolves(ok("fakeToken"));
    sandbox
      .stub(tools.tokenProvider.m365TokenProvider, "getJsonObject")
      .resolves(ok({ tid: "mockSwitchedTid", upn: "mockUpn" }));
    sandbox
      .stub(tools.tokenProvider.azureAccountProvider, "getAccountCredentialAsync")
      .resolves(TestHelper.fakeCredential);
    sandbox.stub(AppStudioClient, "getApp").onFirstCall().throws({}).onSecondCall().resolves({});
    sandbox.stub(AppStudioClient, "importApp").resolves({ teamsAppId: "mockTeamsAppId" });
    sandbox.stub(clientFactory, "createResourceProviderClient").resolves({});
    sandbox.stub(clientFactory, "ensureResourceProvider").resolves();
    sandbox.stub(AADRegistration, "registerAADAppAndGetSecretByGraph").resolves({
      clientId: "mockClientId",
      clientSecret: "mockClientSecret",
      objectId: "mockObjectId",
    });
    const appName = `unittest${randomAppName()}`;
    const inputs: InputsWithProjectPath = {
      projectPath: projectPath,
      platform: Platform.VSCode,
      features: "Bot",
      language: "typescript",
      "app-name": appName,
      folder: path.join(os.homedir(), "TeamsApps"),
      checkerInfo: {
        skipNgrok: true,
      },
    };
    const initRes = await fx.init(context, inputs);
    if (initRes.isErr()) {
      console.log(initRes.error);
    }
    assert.isTrue(initRes.isOk());

    context.projectSetting.components = [
      {
        name: "teams-bot",
        build: true,
        capabilities: ["bot"],
        deploy: true,
        folder: "bot",
        hosting: "azure-web-app",
      },
    ];
    context.envInfo = newEnvInfoV3("local");
    context.tokenProvider = tools.tokenProvider;
    context.envInfo.state = {
      solution: {
        provisionSucceeded: true,
        teamsAppTenantId: "mockTid",
      },
      "app-manifest": {
        tenantId: "mockTid",
        teamsAppId: "mockTeamsAppId",
      },
    };
    context.envInfo.config.bot = {
      siteEndpoint: "https://localtest:3978",
    };
    const provisionRes = await runActionByName("fx.provision", context, inputs);
    if (provisionRes.isErr()) {
      console.log(provisionRes.error);
    }
    assert.isTrue(provisionRes.isOk());
    assert.isTrue(context.envInfo.state.solution.teamsAppTenantId === "mockSwitchedTid");
    assert.isTrue(context.envInfo.state.solution.provisionSucceeded);
    assert.isTrue(context.envInfo.state["app-manifest"]["tenantId"] === "mockSwitchedTid");
  });

  it("azure-storage.deploy", async () => {
    sandbox.stub(templateAction, "scaffoldFromTemplates").resolves();
    sandbox.stub(FrontendDeployment, "doFrontendDeploymentV3").resolves();
    sandbox.stub(aadManifest, "generateAadManifestTemplate").resolves();
    sandbox
      .stub(tools.tokenProvider.azureAccountProvider, "getAccountCredentialAsync")
      .resolves(TestHelper.fakeCredential);
    const inputs: InputsWithProjectPath = {
      projectPath: projectPath,
      platform: Platform.VSCode,
      folder: path.join(projectPath, "tabs"),
      componentId: "teams-tab",
    };
    context.envInfo = newEnvInfoV3();
    context.tokenProvider = tools.tokenProvider;
    context.envInfo.state = {
      solution: {
        provisionSucceeded: true,
        needCreateResourceGroup: false,
        resourceGroupName: "mockRG",
        location: "eastasia",
        resourceNameSuffix: "3bf854123",
        teamsAppTenantId: "mockTid",
        subscriptionId: "mockSid",
        subscriptionName: "mockSname",
        tenantId: "mockAzureTid",
      },
      "teams-tab": {
        location: "centreus",
        storageResourceId:
          "/subscriptions/mockSid/resourceGroups/jay-texas/providers/Microsoft.Storage/storageAccounts/testaccount",
        endpoint: "https://testaccount.azurewebsites.net",
      },
    };
    const addTabRes = await runActionByName("teams-tab.add", context, inputs);
    if (addTabRes.isErr()) {
      console.log(addTabRes.error);
    }
    assert.isTrue(addTabRes.isOk());
    const azureStorage = Container.get<AzureStorageResource>(ComponentNames.AzureStorage);
    const res = await azureStorage.deploy(context as ResourceContextV3, inputs);
    if (res.isErr()) {
      console.log(res.error);
    }
    assert.isTrue(res.isOk());
  });
});
