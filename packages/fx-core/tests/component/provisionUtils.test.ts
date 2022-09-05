import { ok, Platform, v2 } from "@microsoft/teamsfx-api";
import chai from "chai";
import * as sinon from "sinon";
import { provisionUtils } from "../../src/component/provisionUtils";
import { createContextV3 } from "../../src/component/utils";
import { SolutionError } from "../../src/plugins/solution";
import { resourceGroupHelper } from "../../src/plugins/solution/fx-solution/utils/ResourceGroupHelper";
import { MockAzureAccountProvider } from "../core/utils";
import { TestHelper } from "../plugins/resource/frontend/helper";

const expect = chai.expect;

describe("checkProvisionSubscription", () => {
  const mocker = sinon.createSandbox();

  afterEach(() => {
    mocker.restore();
  });

  it("provision with CLI parameters succeeds", async () => {
    const context = createContextV3();
    const azureAccountProvider = new MockAzureAccountProvider();
    const envInfo = {
      envName: "test",
      config: {},
      state: { solution: {} },
    };
    mocker.stub(context.logProvider, "log").resolves(true);
    mocker
      .stub(azureAccountProvider, "getAccountCredentialAsync")
      .resolves(TestHelper.fakeCredential);
    mocker.stub(azureAccountProvider, "listSubscriptions").resolves([
      {
        subscriptionName: "mockSubName",
        subscriptionId: "cli-sub",
        tenantId: "mockTenantId",
      },
    ]);

    const res = await provisionUtils.checkProvisionSubscription(
      context,
      envInfo,
      azureAccountProvider,
      "cli-sub",
      "test"
    );

    expect(res.isOk()).equal(true);
    if (res.isErr()) {
      console.log(res.error);
    }
    expect((envInfo.state.solution as any).subscriptionId).equal("cli-sub");
  });

  it("provision with CLI parameters error", async () => {
    const context = createContextV3();
    const azureAccountProvider = new MockAzureAccountProvider();
    const envInfo = {
      envName: "test",
      config: {},
      state: { solution: {} },
    };
    mocker.stub(context.logProvider, "log").resolves(true);
    mocker
      .stub(azureAccountProvider, "getAccountCredentialAsync")
      .resolves(TestHelper.fakeCredential);
    mocker.stub(azureAccountProvider, "listSubscriptions").resolves([
      {
        subscriptionName: "mockSubName",
        subscriptionId: "sub",
        tenantId: "mockTenantId",
      },
    ]);

    const res = await provisionUtils.checkProvisionSubscription(
      context,
      envInfo,
      azureAccountProvider,
      "cli-sub",
      "test"
    );

    expect(res.isErr()).equal(true);
    if (res.isErr()) {
      console.log(res.error);
      expect(res.error.name).equals(SolutionError.SubscriptionNotFound);
    }
  });

  it("provision with config error", async () => {
    const context = createContextV3();
    const azureAccountProvider = new MockAzureAccountProvider();
    const envInfo = {
      envName: "test",
      config: {
        azure: {
          subscriptionId: "mockSub",
        },
      },
      state: { solution: {} },
    };
    mocker.stub(context.logProvider, "log").resolves(true);
    mocker
      .stub(azureAccountProvider, "getAccountCredentialAsync")
      .resolves(TestHelper.fakeCredential);
    mocker.stub(azureAccountProvider, "listSubscriptions").resolves([
      {
        subscriptionName: "mockSubName",
        subscriptionId: "sub",
        tenantId: "mockTenantId",
      },
    ]);

    const res = await provisionUtils.checkProvisionSubscription(
      context,
      envInfo,
      azureAccountProvider,
      undefined,
      "test"
    );

    expect(res.isErr()).equal(true);
    if (res.isErr()) {
      console.log(res.error);
      expect(res.error.name).equals(SolutionError.SubscriptionNotFound);
    }
  });

  it("provision with config succeeds", async () => {
    const context = createContextV3();
    const azureAccountProvider = new MockAzureAccountProvider();
    const envInfo = {
      envName: "test",
      config: {
        azure: {
          subscriptionId: "mockSub",
        },
      },
      state: { solution: {} },
    };
    mocker.stub(context.logProvider, "log").resolves(true);
    mocker
      .stub(azureAccountProvider, "getAccountCredentialAsync")
      .resolves(TestHelper.fakeCredential);
    mocker.stub(azureAccountProvider, "listSubscriptions").resolves([
      {
        subscriptionName: "mockSubName",
        subscriptionId: "mockSub",
        tenantId: "mockTenantId",
      },
    ]);

    const res = await provisionUtils.checkProvisionSubscription(
      context,
      envInfo,
      azureAccountProvider,
      undefined,
      "test"
    );

    expect(res.isOk()).equal(true);
    if (res.isErr()) {
      console.log(res.error);
    }
    expect((envInfo.state.solution as any).subscriptionId).equal("mockSub");
  });
});

describe("fillInAzureConfigs", () => {
  const mocker = sinon.createSandbox();

  afterEach(() => {
    mocker.restore();
  });

  it("provision with CLI parameters succeeds", async () => {
    const context = createContextV3();
    const azureAccountProvider = new MockAzureAccountProvider();
    const envInfo = {
      envName: "test",
      config: {},
      state: { solution: {} },
    };
    const inputs: v2.InputsWithProjectPath = {
      platform: Platform.CLI,
      projectPath: "path",
      targetSubscriptionId: "cli-sub",
      targetResourceGroupName: "cli-rg",
    };
    mocker.stub(context.logProvider, "log").resolves(true);
    mocker
      .stub(azureAccountProvider, "getAccountCredentialAsync")
      .resolves(TestHelper.fakeCredential);
    mocker.stub(resourceGroupHelper, "getResourceGroupInfo").resolves(
      ok({
        createNewResourceGroup: false,
        name: "cli-rg",
        location: "East US",
      })
    );
    mocker.stub(azureAccountProvider, "listSubscriptions").resolves([
      {
        subscriptionName: "mockSubName",
        subscriptionId: "cli-sub",
        tenantId: "mockTenantId",
      },
    ]);
    const tokenProvider = { azureAccountProvider };

    const res = await provisionUtils.fillInAzureConfigs(
      context,
      inputs,
      envInfo,
      tokenProvider as any
    );

    if (res.isErr()) {
      console.log(res.error);
    }
    expect(res.isOk()).equal(true);
    expect((envInfo.state.solution as any).subscriptionId).equal("cli-sub");
    expect((envInfo.state.solution as any).resourceGroupName).equal("cli-rg");
  });

  it("provision with CLI parameters resource group not exist", async () => {
    const context = createContextV3();
    const azureAccountProvider = new MockAzureAccountProvider();
    const envInfo = {
      envName: "test",
      config: {},
      state: { solution: {} },
    };
    const inputs: v2.InputsWithProjectPath = {
      platform: Platform.CLI,
      projectPath: "path",
      targetSubscriptionId: "cli-sub",
      targetResourceGroupName: "cli-rg",
    };
    mocker.stub(context.logProvider, "log").resolves(true);
    mocker
      .stub(azureAccountProvider, "getAccountCredentialAsync")
      .resolves(TestHelper.fakeCredential);
    mocker.stub(resourceGroupHelper, "getResourceGroupInfo").resolves(ok(undefined));
    mocker.stub(azureAccountProvider, "listSubscriptions").resolves([
      {
        subscriptionName: "mockSubName",
        subscriptionId: "cli-sub",
        tenantId: "mockTenantId",
      },
    ]);
    const tokenProvider = { azureAccountProvider };

    const res = await provisionUtils.fillInAzureConfigs(
      context,
      inputs,
      envInfo,
      tokenProvider as any
    );
    expect(res.isErr()).equal(true);
    if (res.isErr()) {
      expect(res.error.name).equal(SolutionError.ResourceGroupNotFound);
    }
  });

  it("provision with resource group name from config file succeeds", async () => {
    const context = createContextV3();
    const azureAccountProvider = new MockAzureAccountProvider();
    const envInfo = {
      envName: "test",
      config: {
        azure: {
          subscriptionId: "mockSub",
          resourceGroupName: "mockRg",
        },
      },
      state: { solution: {} },
    };
    const inputs: v2.InputsWithProjectPath = {
      platform: Platform.CLI,
      projectPath: "path",
    };
    mocker.stub(context.logProvider, "log").resolves(true);
    mocker
      .stub(azureAccountProvider, "getAccountCredentialAsync")
      .resolves(TestHelper.fakeCredential);
    mocker.stub(resourceGroupHelper, "getResourceGroupInfo").resolves(
      ok({
        createNewResourceGroup: false,
        name: "mockRg",
        location: "East US",
      })
    );
    mocker.stub(azureAccountProvider, "listSubscriptions").resolves([
      {
        subscriptionName: "mockSubName",
        subscriptionId: "mockSub",
        tenantId: "mockTenantId",
      },
    ]);
    const tokenProvider = { azureAccountProvider };

    const res = await provisionUtils.fillInAzureConfigs(
      context,
      inputs,
      envInfo,
      tokenProvider as any
    );

    expect(res.isOk()).equal(true);
    expect((envInfo.state.solution as any).subscriptionId).equal("mockSub");
    expect((envInfo.state.solution as any).resourceGroupName).equal("mockRg");
  });

  it("provision with resource group name from config file not exist", async () => {
    const context = createContextV3();
    const azureAccountProvider = new MockAzureAccountProvider();
    const envInfo = {
      envName: "test",
      config: {
        azure: {
          subscriptionId: "mockSub",
          resourceGroupName: "mockRg",
        },
      },
      state: { solution: {} },
    };
    const inputs: v2.InputsWithProjectPath = {
      platform: Platform.CLI,
      projectPath: "path",
    };
    mocker.stub(context.logProvider, "log").resolves(true);
    mocker
      .stub(azureAccountProvider, "getAccountCredentialAsync")
      .resolves(TestHelper.fakeCredential);
    mocker.stub(resourceGroupHelper, "getResourceGroupInfo").resolves(ok(undefined));
    mocker.stub(azureAccountProvider, "listSubscriptions").resolves([
      {
        subscriptionName: "mockSubName",
        subscriptionId: "mockSub",
        tenantId: "mockTenantId",
      },
    ]);
    const tokenProvider = { azureAccountProvider };

    const res = await provisionUtils.fillInAzureConfigs(
      context,
      inputs,
      envInfo,
      tokenProvider as any
    );

    expect(res.isErr()).equal(true);
    if (res.isErr()) {
      expect(res.error.name).equal(SolutionError.ResourceGroupNotFound);
    }
  });

  it("provision with resource group name from config file but missing subscription id", async () => {
    const context = createContextV3();
    const azureAccountProvider = new MockAzureAccountProvider();
    const envInfo = {
      envName: "test",
      config: {
        azure: {
          resourceGroupName: "mockRg",
        },
      },
      state: { solution: {} },
    };
    const inputs: v2.InputsWithProjectPath = {
      platform: Platform.CLI,
      projectPath: "path",
    };
    mocker.stub(context.logProvider, "log").resolves(true);
    mocker
      .stub(azureAccountProvider, "getAccountCredentialAsync")
      .resolves(TestHelper.fakeCredential);
    mocker.stub(resourceGroupHelper, "getResourceGroupInfo").resolves(ok(undefined));
    mocker.stub(azureAccountProvider, "listSubscriptions").resolves([
      {
        subscriptionName: "mockSubName",
        subscriptionId: "mockSub",
        tenantId: "mockTenantId",
      },
    ]);
    mocker.stub(azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionName: "mockSubName",
      subscriptionId: "mockSub",
      tenantId: "mockTenantId",
    });
    const tokenProvider = { azureAccountProvider };

    const res = await provisionUtils.fillInAzureConfigs(
      context,
      inputs,
      envInfo,
      tokenProvider as any
    );

    expect(res.isErr()).equal(true);
    if (res.isErr()) {
      expect(res.error.name).equal(SolutionError.MissingSubscriptionIdInConfig);
    }
  });
});
