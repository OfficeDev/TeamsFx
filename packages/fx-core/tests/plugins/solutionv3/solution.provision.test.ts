// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  AzureAccountProvider,
  FxError,
  Inputs,
  ok,
  Platform,
  ProjectSettings,
  Result,
  SubscriptionInfo,
  TokenProvider,
  TokenRequest,
  v2,
  v3,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import sinon from "sinon";
import * as uuid from "uuid";
import fs from "fs-extra";
import arm from "../../../src/plugins/solution/fx-solution/arm";
import {
  BuiltInFeaturePluginNames,
  BuiltInSolutionNames,
} from "../../../src/plugins/solution/fx-solution/v3/constants";
import {
  getQuestionsForProvision,
  provisionResources,
} from "../../../src/plugins/solution/fx-solution/v3/provision";
import { MockedM365Provider, MockedAzureAccountProvider, MockedV2Context } from "../solution/util";
import { MockFeaturePluginNames } from "./mockPlugins";
import * as path from "path";
import * as os from "os";
import { randomAppName } from "../../core/utils";
import { resourceGroupHelper } from "../../../src/plugins/solution/fx-solution/utils/ResourceGroupHelper";
import { ResourceManagementClient } from "@azure/arm-resources";
import * as appStudio from "../../../src/component/resource/appManifest/appStudio";
import {
  publishApplication,
  getQuestionsForPublish,
} from "../../../src/plugins/solution/fx-solution/v3/publish";
import { TestHelper } from "../solution/helper";
import { TestFilePath } from "../../constants";
describe("SolutionV3 - provision", () => {
  const sandbox = sinon.createSandbox();
  beforeEach(async () => {
    sandbox
      .stub<any, any>(arm, "deployArmTemplates")
      .callsFake(
        async (
          ctx: v2.Context,
          inputs: v2.InputsWithProjectPath,
          envInfo: v3.EnvInfoV3,
          azureAccountProvider: AzureAccountProvider
        ): Promise<Result<void, FxError>> => {
          return ok(undefined);
        }
      );
    sandbox
      .stub<any, any>(resourceGroupHelper, "askResourceGroupInfo")
      .callsFake(
        async (
          ctx: v2.Context,
          inputs: Inputs,
          azureAccountProvider: AzureAccountProvider,
          rmClient: ResourceManagementClient,
          defaultResourceGroupName: string
        ): Promise<Result<any, FxError>> => {
          return ok({
            createNewResourceGroup: false,
            name: "mockRG",
            location: "mockLoc",
          });
        }
      );
  });
  afterEach(async () => {
    sandbox.restore();
  });
  it("provision", async () => {
    const projectSettings: ProjectSettings = {
      appName: "my app",
      projectId: uuid.v4(),
      solutionSettings: {
        name: BuiltInSolutionNames.azure,
        version: "3.0.0",
        capabilities: ["Tab"],
        hostType: "Azure",
        azureResources: [],
        activeResourcePlugins: [MockFeaturePluginNames.tab],
      },
    };
    const ctx = new MockedV2Context(projectSettings);
    const inputs: v2.InputsWithProjectPath = {
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir(), randomAppName()),
    };
    const mockedTokenProvider: TokenProvider = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      m365TokenProvider: new MockedM365Provider(),
    };
    const mockSub: SubscriptionInfo = {
      subscriptionId: "mockSubId",
      subscriptionName: "mockSubName",
      tenantId: "mockTenantId",
    };
    sandbox
      .stub<any, any>(mockedTokenProvider.azureAccountProvider, "listSubscriptions")
      .callsFake(async (): Promise<SubscriptionInfo[]> => {
        return [mockSub];
      });
    sandbox
      .stub<any, any>(mockedTokenProvider.m365TokenProvider, "getJsonObject")
      .callsFake(
        async (tokenRequest: TokenRequest): Promise<Result<Record<string, unknown>, FxError>> => {
          return ok({ tid: "mock-tenant-id" });
        }
      );
    sandbox
      .stub<any, any>(ctx.userInteraction, "showMessage")
      .callsFake(
        async (
          level: "info" | "warn" | "error",
          message: string,
          modal: boolean,
          ...items: string[]
        ): Promise<Result<string | undefined, FxError>> => {
          return ok("Provision");
        }
      );
    sandbox.stub(appStudio, "createTeamsApp").resolves(ok(uuid.v4()));
    sandbox.stub(appStudio, "updateTeamsApp").resolves(ok(uuid.v4()));

    const envInfoV3: v3.EnvInfoV3 = {
      envName: "dev",
      state: { solution: { ...mockSub }, "fx-resource-appstudio": {} },
      config: {},
    };
    const res = await provisionResources(ctx, inputs, envInfoV3, mockedTokenProvider);
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.isTrue(res.value.state["fx-resource-appstudio"].tenantId === "mock-tenant-id");
    }
  });

  it("provision - has provisioned before in same account", async () => {
    const parameterFileNameTemplate = (env: string) => `azure.parameters.${env}.json`;
    const configDir = path.join(TestHelper.rootDir, TestFilePath.configFolder);
    const targetEnvName = "dev";
    const originalResourceBaseName = "originalResourceBaseName";
    const paramContent = TestHelper.getParameterFileContent(
      {
        resourceBaseName: originalResourceBaseName,
        param1: "value1",
        param2: "value2",
      },
      {
        userParam1: "userParamValue1",
        userParam2: "userParamValue2",
      }
    );

    after(async () => {
      await fs.remove(TestHelper.rootDir);
    });
    await fs.ensureDir(configDir);

    await fs.writeFile(
      path.join(configDir, parameterFileNameTemplate(targetEnvName)),
      paramContent
    );

    const projectSettings: ProjectSettings = {
      appName: "my app",
      projectId: uuid.v4(),
      solutionSettings: {
        name: BuiltInSolutionNames.azure,
        version: "3.0.0",
        capabilities: ["Tab"],
        hostType: "Azure",
        azureResources: [BuiltInFeaturePluginNames.identity],
        activeResourcePlugins: ["fx-resource-identity"],
      },
    };
    const ctx = new MockedV2Context(projectSettings);
    const inputs: v2.InputsWithProjectPath = {
      platform: Platform.VSCode,
      projectPath: TestHelper.rootDir,
    };
    const mockedTokenProvider: TokenProvider = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      m365TokenProvider: new MockedM365Provider(),
    };
    const mockSub: SubscriptionInfo = {
      subscriptionId: "mockSubId",
      subscriptionName: "mockSubName",
      tenantId: "mockTenantId",
    };
    const mockSwitchedSub: SubscriptionInfo = {
      subscriptionId: "mockNewSubId",
      subscriptionName: "mockNewSubName",
      tenantId: "mockNewTenantId",
    };
    sandbox
      .stub<any, any>(mockedTokenProvider.azureAccountProvider, "getSelectedSubscription")
      .callsFake(async (): Promise<SubscriptionInfo> => {
        return mockSwitchedSub;
      });
    sandbox
      .stub<any, any>(mockedTokenProvider.azureAccountProvider, "listSubscriptions")
      .callsFake(async (): Promise<SubscriptionInfo[]> => {
        return [mockSub, mockSwitchedSub];
      });
    sandbox
      .stub<any, any>(mockedTokenProvider.m365TokenProvider, "getJsonObject")
      .callsFake(
        async (tokenRequest: TokenRequest): Promise<Result<Record<string, unknown>, FxError>> => {
          return ok({ tid: "mock-tenant-id" });
        }
      );
    sandbox
      .stub<any, any>(ctx.userInteraction, "showMessage")
      .callsFake(
        async (
          level: "info" | "warn" | "error",
          message: string,
          modal: boolean,
          ...items: string[]
        ): Promise<Result<string | undefined, FxError>> => {
          return ok("Provision");
        }
      );
    sandbox.stub(appStudio, "createTeamsApp").resolves(ok(uuid.v4()));
    sandbox.stub(appStudio, "updateTeamsApp").resolves(ok(uuid.v4()));

    const envInfoV3: v3.EnvInfoV3 = {
      envName: "dev",
      state: { solution: { ...mockSub }, "fx-resource-appstudio": {} },
      config: {},
    };
    const res = await provisionResources(ctx, inputs, envInfoV3, mockedTokenProvider);
    assert.isTrue(res.isOk());

    if (res.isOk()) {
      assert.isTrue(res.value.state["fx-resource-appstudio"].tenantId === "mock-tenant-id");
      assert.isTrue(res.value.state.solution.subscriptionId === "mockNewSubId");
      assert.isTrue(res.value.state.solution.subscriptionName === "mockNewSubName");
      assert.isTrue(res.value.state.solution.tenantId === "mockNewTenantId");
    }
  });

  it("getQuestionsForProvision", async () => {
    const projectSettings: ProjectSettings = {
      appName: "my app",
      projectId: uuid.v4(),
      solutionSettings: {
        name: BuiltInSolutionNames.azure,
        version: "3.0.0",
        capabilities: ["Tab"],
        hostType: "Azure",
        azureResources: [],
        activeResourcePlugins: [MockFeaturePluginNames.tab],
      },
    };
    const ctx = new MockedV2Context(projectSettings);
    const inputs: v2.InputsWithProjectPath = {
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir(), randomAppName()),
    };
    const mockedTokenProvider: TokenProvider = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      m365TokenProvider: new MockedM365Provider(),
    };
    const envInfoV3: v2.DeepReadonly<v3.EnvInfoV3> = {
      envName: "dev",
      config: {},
      state: { solution: {} },
    };
    const res = await getQuestionsForProvision(ctx, inputs, envInfoV3, mockedTokenProvider);
    assert.isTrue(res.isOk());
  });
});
