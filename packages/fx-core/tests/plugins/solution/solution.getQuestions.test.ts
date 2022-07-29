// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  AzureSolutionSettings,
  Func,
  Inputs,
  MultiSelectQuestion,
  ok,
  OptionItem,
  Platform,
  ProjectSettings,
  SingleSelectQuestion,
  Stage,
  TokenProvider,
  v2,
} from "@microsoft/teamsfx-api";
import { EnvInfoV2 } from "@microsoft/teamsfx-api/build/v2";
import chai, { assert } from "chai";
import chaiAsPromised from "chai-as-promised";
import { it } from "mocha";
import * as sinon from "sinon";
import Container from "typedi";
import * as uuid from "uuid";
import "../../../src/plugins/resource/apim/v2";
import "../../../src/plugins/resource/appstudio/v2";
import { AppStudioPluginV3 } from "../../../src/plugins/resource/appstudio/v3";
import "../../../src/plugins/resource/bot/v2";
import "../../../src/plugins/resource/frontend/v2";
import "../../../src/plugins/resource/function/v2";
import "../../../src/plugins/resource/localdebug/v2";
import "../../../src/plugins/resource/spfx/v2";
import "../../../src/plugins/resource/sql/v2";
import * as tool from "../../../src/common/tools";
import * as featureFlags from "../../../src/common/featureFlags";
import {
  GLOBAL_CONFIG,
  SOLUTION_PROVISION_SUCCEEDED,
} from "../../../src/plugins/solution/fx-solution/constants";
import {
  AzureResourceApimNewUI,
  AzureResourceFunctionNewUI,
  AzureResourceKeyVaultNewUI,
  AzureResourceSQLNewUI,
  BotNewUIOptionItem,
  BotOptionItem,
  CicdOptionItem,
  CommandAndResponseOptionItem,
  HostTypeOptionAzure,
  HostTypeOptionSPFx,
  TabNewUIOptionItem,
  MessageExtensionItem,
  MessageExtensionNewUIItem,
  NotificationOptionItem,
  ApiConnectionOptionItem,
  TabNonSsoItem,
  TabOptionItem,
  SingleSignOnOptionItem,
} from "../../../src/plugins/solution/fx-solution/question";
import { ResourcePluginsV2 } from "../../../src/plugins/solution/fx-solution/ResourcePluginContainer";
import {
  getQuestions,
  getQuestionsForAddFeature,
  getQuestionsForScaffolding,
  getQuestionsForUserTask,
} from "../../../src/plugins/solution/fx-solution/v2/getQuestions";
import { BuiltInFeaturePluginNames } from "../../../src/plugins/solution/fx-solution/v3/constants";
import { MockedM365Provider, MockedAzureAccountProvider, MockedV2Context } from "./util";
import { BotCapabilities, PluginBot } from "../../../src/plugins/resource/bot/resources/strings";
import { BotHostTypes } from "../../../src";

chai.use(chaiAsPromised);
const expect = chai.expect;
const functionPluginV2 = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.FunctionPlugin);
const sqlPluginV2 = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.SqlPlugin);
const spfxPluginV2 = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.SpfxPlugin);
const frontendPluginV2 = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.FrontendPlugin);
const botPluginV2 = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.BotPlugin);
const cicdPlugin = Container.get<v2.ResourcePlugin>(ResourcePluginsV2.CICDPlugin);
const mockedProvider: TokenProvider = {
  azureAccountProvider: new MockedAzureAccountProvider(),
  m365TokenProvider: new MockedM365Provider(),
};
const envInfo: EnvInfoV2 = {
  envName: "default",
  config: {},
  state: { solution: {} },
};

describe("getQuestionsForScaffolding()", async () => {
  const sandbox = sinon.createSandbox();
  const projectSettings: ProjectSettings = {
    appName: "my app",
    projectId: uuid.v4(),
    solutionSettings: {
      hostType: HostTypeOptionAzure.id,
      name: "test",
      version: "1.0",
      activeResourcePlugins: ["fx-resource-frontend-hosting", "fx-resource-aad-app-for-teams"],
      capabilities: [],
      azureResources: [],
    },
  };

  beforeEach(() => {
    spfxPluginV2.getQuestionsForScaffolding = async function () {
      return ok(undefined);
    };
    frontendPluginV2.getQuestionsForScaffolding = async function () {
      return ok(undefined);
    };
    functionPluginV2.getQuestionsForScaffolding = async function () {
      return ok(undefined);
    };
    sqlPluginV2.getQuestionsForScaffolding = async function () {
      return ok(undefined);
    };
    botPluginV2.getQuestionsForScaffolding = async function () {
      return ok(undefined);
    };
    cicdPlugin.getQuestionsForUserTask = async function () {
      return ok(undefined);
    };
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("getQuestionsForScaffolding", async () => {
    const mockedCtx = new MockedV2Context(projectSettings);
    const mockedInputs: Inputs = {
      platform: Platform.VSCode,
    };
    const result = await getQuestionsForScaffolding(mockedCtx, mockedInputs);
    expect(result.isOk()).to.be.true;
  });

  it("getQuestions - provision", async () => {
    const mockedCtx = new MockedV2Context(projectSettings);
    const mockedInputs: Inputs = {
      platform: Platform.VSCode,
      stage: Stage.provision,
    };
    const result = await getQuestions(mockedCtx, mockedInputs, envInfo, mockedProvider);
    assert.isTrue(result.isOk());
    if (result.isOk()) {
      const node = result.value;
      assert.isTrue(node !== undefined && node.data !== undefined);
    }
  });

  it("getQuestions - deploy", async () => {
    (projectSettings.solutionSettings as AzureSolutionSettings).capabilities.push(TabOptionItem.id);
    const mockedCtx = new MockedV2Context(projectSettings);
    const mockedInputs: Inputs = {
      platform: Platform.VSCode,
      stage: Stage.deploy,
    };
    envInfo.state[GLOBAL_CONFIG][SOLUTION_PROVISION_SUCCEEDED] = false;
    const result1 = await getQuestions(mockedCtx, mockedInputs, envInfo, mockedProvider);
    assert.isTrue(result1.isErr());
    envInfo.state[GLOBAL_CONFIG][SOLUTION_PROVISION_SUCCEEDED] = true;
    const result2 = await getQuestions(mockedCtx, mockedInputs, envInfo, mockedProvider);
    assert.isTrue(result2.isOk());
    if (result2.isOk()) {
      const node = result2.value as any;
      assert.isTrue(
        node !== undefined &&
          node.children[0].data.default.length === 1 &&
          node.children[0].data.default.includes("fx-resource-frontend-hosting")
      );
    }

    // Not show AAD plugin
    sandbox.stub<any, any>(tool, "isAadManifestEnabled").returns(true);
    const result3 = await getQuestions(mockedCtx, mockedInputs, envInfo, mockedProvider);
    chai.assert.isTrue(result3.isOk());
    if (result3.isOk()) {
      const node = result3.value as any;
      assert.isTrue(
        node !== undefined &&
          node.children[0].data.default.length === 1 &&
          node.children[0].data.default.includes("fx-resource-frontend-hosting")
      );
    }
  });

  it("getQuestions - publish", async () => {
    const mockedCtx = new MockedV2Context(projectSettings);
    const mockedInputs: Inputs = {
      platform: Platform.VSCode,
      stage: Stage.publish,
    };
    envInfo.state[GLOBAL_CONFIG][SOLUTION_PROVISION_SUCCEEDED] = false;
    const result1 = await getQuestions(mockedCtx, mockedInputs, envInfo, mockedProvider);
    assert.isTrue(result1.isErr());

    (mockedCtx.projectSetting.solutionSettings as AzureSolutionSettings).hostType =
      HostTypeOptionSPFx.id;
    const result11 = await getQuestions(mockedCtx, mockedInputs, envInfo, mockedProvider);
    assert.isTrue(result11.isErr());

    envInfo.state[GLOBAL_CONFIG][SOLUTION_PROVISION_SUCCEEDED] = true;
    const result2 = await getQuestions(mockedCtx, mockedInputs, envInfo, mockedProvider);
    assert.isTrue(result2.isOk());
    if (result2.isOk()) {
      const node = result2.value;
      assert.isTrue(node !== undefined && node.data !== undefined);
    }
  });

  it("getQuestions - grantPermission", async () => {
    const mockedCtx = new MockedV2Context(projectSettings);
    const mockedInputs: Inputs = {
      platform: Platform.VSCode,
      stage: Stage.grantPermission,
    };
    const result2 = await getQuestions(mockedCtx, mockedInputs, envInfo, mockedProvider);
    assert.isTrue(result2.isOk());
    if (result2.isOk()) {
      const node = result2.value;
      assert.isTrue(node !== undefined && node.data !== undefined);
    }
  });
  it("getQuestionsForUserTask - addCapability for SPFx failed", async () => {
    const mockedCtx = new MockedV2Context(projectSettings);
    const mockedInputs: Inputs = {
      platform: Platform.VSCode,
      stage: Stage.grantPermission,
    };
    const func: Func = {
      method: "addCapability",
      namespace: "fx-solution-azure",
    };
    {
      (mockedCtx.projectSetting.solutionSettings as AzureSolutionSettings).hostType =
        HostTypeOptionSPFx.id;
      const res = await getQuestionsForUserTask(
        mockedCtx,
        mockedInputs,
        func,
        envInfo,
        mockedProvider
      );
      assert.isTrue(res.isErr());
    }
  });

  it("getQuestionsForUserTask - addCapability success", async () => {
    const mockedCtx = new MockedV2Context(projectSettings);
    const mockedInputs: Inputs = {
      platform: Platform.VSCode,
      stage: Stage.grantPermission,
    };
    const func: Func = {
      method: "addCapability",
      namespace: "fx-solution-azure",
    };
    const appStudioPlugin = Container.get<AppStudioPluginV3>(BuiltInFeaturePluginNames.appStudio);
    sandbox.stub<any, any>(featureFlags, "isBotNotificationEnabled").returns(false);
    sandbox.stub<any, any>(tool, "isAadManifestEnabled").returns(false);
    sandbox
      .stub<any, any>(appStudioPlugin, "capabilityExceedLimit")
      .callsFake(
        async (
          ctx: v2.Context,
          inputs: v2.InputsWithProjectPath,
          capability: "staticTab" | "configurableTab" | "Bot" | "MessageExtension"
        ) => {
          return ok(false);
        }
      );
    (mockedCtx.projectSetting.solutionSettings as AzureSolutionSettings).hostType =
      HostTypeOptionAzure.id;
    const res = await getQuestionsForUserTask(
      mockedCtx,
      mockedInputs,
      func,
      envInfo,
      mockedProvider
    );
    assert.isTrue(res.isOk() && res.value && res.value.data !== undefined);
    if (res.isOk()) {
      const node = res.value;
      assert.isTrue(
        node &&
          node.data &&
          node.data.type === "multiSelect" &&
          node.data.staticOptions.length === 3
      );
      if (node && node.data && node.data.type === "multiSelect") {
        assert.deepEqual((node.data as MultiSelectQuestion).staticOptions as OptionItem[], [
          BotOptionItem,
          TabOptionItem,
          MessageExtensionItem,
        ]);
      }
    }
  });

  it("getQuestionsForUserTask - addCapability failed because of capabilityExceedLimit", async () => {
    const mockedCtx = new MockedV2Context(projectSettings);
    const mockedInputs: Inputs = {
      platform: Platform.VSCode,
      stage: Stage.grantPermission,
    };
    const func: Func = {
      method: "addCapability",
      namespace: "fx-solution-azure",
    };
    const appStudioPlugin = Container.get<AppStudioPluginV3>(BuiltInFeaturePluginNames.appStudio);
    sandbox
      .stub<any, any>(appStudioPlugin, "capabilityExceedLimit")
      .callsFake(
        async (
          ctx: v2.Context,
          inputs: v2.InputsWithProjectPath,
          capability: "staticTab" | "configurableTab" | "Bot" | "MessageExtension"
        ) => {
          return ok(true);
        }
      );
    (mockedCtx.projectSetting.solutionSettings as AzureSolutionSettings).hostType =
      HostTypeOptionAzure.id;
    const res = await getQuestionsForUserTask(
      mockedCtx,
      mockedInputs,
      func,
      envInfo,
      mockedProvider
    );
    assert.isTrue(res.isOk() && res.value === undefined);
  });

  it("getQuestionsForUserTask - addResource", async () => {
    const mockedCtx = new MockedV2Context(projectSettings);
    const mockedInputs: Inputs = {
      platform: Platform.VSCode,
      stage: Stage.grantPermission,
    };
    const func: Func = {
      method: "addResource",
      namespace: "fx-solution-azure",
    };
    {
      (mockedCtx.projectSetting.solutionSettings as AzureSolutionSettings).hostType =
        HostTypeOptionSPFx.id;
      const res = await getQuestionsForUserTask(
        mockedCtx,
        mockedInputs,
        func,
        envInfo,
        mockedProvider
      );
      assert.isTrue(res.isErr());
    }
    {
      (mockedCtx.projectSetting.solutionSettings as AzureSolutionSettings).hostType =
        HostTypeOptionAzure.id;
      const res = await getQuestionsForUserTask(
        mockedCtx,
        mockedInputs,
        func,
        envInfo,
        mockedProvider
      );
      assert.isTrue(res.isOk());
    }
    {
      (mockedCtx.projectSetting.solutionSettings as AzureSolutionSettings).hostType =
        HostTypeOptionAzure.id;
      (mockedCtx.projectSetting.solutionSettings as AzureSolutionSettings).capabilities = [
        TabOptionItem.id,
      ];
      const res = await getQuestionsForUserTask(
        mockedCtx,
        mockedInputs,
        func,
        envInfo,
        mockedProvider
      );
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        const node = res.value;
        assert.isTrue(node !== undefined && node.data !== undefined);
      }
    }
  });

  it("getQuestionsForAddFeature - CLI_HELP", async () => {
    sandbox.stub<any, any>(tool, "canAddCICDWorkflows").resolves(true);
    const mockedCtx = new MockedV2Context(projectSettings);
    const mockedInputs: Inputs = {
      platform: Platform.CLI_HELP,
      stage: Stage.grantPermission,
      projectPath: "test path",
    };
    const func: Func = {
      method: "addFeature",
      namespace: "fx-solution-azure",
    };
    const appStudioPlugin = Container.get<AppStudioPluginV3>(BuiltInFeaturePluginNames.appStudio);
    sandbox
      .stub<any, any>(appStudioPlugin, "capabilityExceedLimit")
      .callsFake(
        async (
          ctx: v2.Context,
          inputs: v2.InputsWithProjectPath,
          capability: "staticTab" | "configurableTab" | "Bot" | "MessageExtension"
        ) => {
          return ok(false);
        }
      );
    (mockedCtx.projectSetting.solutionSettings as AzureSolutionSettings).hostType =
      HostTypeOptionAzure.id;
    const res = await getQuestionsForAddFeature(
      mockedCtx,
      mockedInputs,
      func,
      envInfo,
      mockedProvider
    );
    assert.isTrue(res.isOk() && res.value && res.value !== undefined);
    if (res.isOk()) {
      const node = res.value;
      assert.isTrue(
        node && node.data && node.data.type === "singleSelect",
        "option item count check"
      );
      if (node && node.data && node.data.type === "singleSelect") {
        const options = (node.data as SingleSelectQuestion).staticOptions as OptionItem[];
        assert.deepEqual(
          options,
          [
            NotificationOptionItem,
            CommandAndResponseOptionItem,
            TabNewUIOptionItem,
            TabNonSsoItem,
            BotNewUIOptionItem,
            MessageExtensionNewUIItem,
            AzureResourceFunctionNewUI,
            AzureResourceApimNewUI,
            AzureResourceSQLNewUI,
            AzureResourceKeyVaultNewUI,
            SingleSignOnOptionItem,
            ApiConnectionOptionItem,
            CicdOptionItem,
          ],
          "option item should match"
        );
      }
    }
  });

  it("getQuestionsForUserTask - addFeature success", async () => {
    sandbox.stub<any, any>(featureFlags, "isPreviewFeaturesEnabled").returns(true);
    sandbox.stub<any, any>(tool, "canAddCICDWorkflows").resolves(true);
    const mockedCtx = new MockedV2Context(projectSettings);
    const mockedInputs: Inputs = {
      platform: Platform.VSCode,
      stage: Stage.grantPermission,
      projectPath: "test path",
    };
    const func: Func = {
      method: "addFeature",
      namespace: "fx-solution-azure",
    };
    const appStudioPlugin = Container.get<AppStudioPluginV3>(BuiltInFeaturePluginNames.appStudio);
    sandbox
      .stub<any, any>(appStudioPlugin, "capabilityExceedLimit")
      .callsFake(
        async (
          ctx: v2.Context,
          inputs: v2.InputsWithProjectPath,
          capability: "staticTab" | "configurableTab" | "Bot" | "MessageExtension"
        ) => {
          return ok(false);
        }
      );
    (mockedCtx.projectSetting.solutionSettings as AzureSolutionSettings).hostType =
      HostTypeOptionAzure.id;
    const res = await getQuestionsForUserTask(
      mockedCtx,
      mockedInputs,
      func,
      envInfo,
      mockedProvider
    );
    assert.isTrue(res.isOk() && res.value && res.value.data !== undefined);
    if (res.isOk()) {
      const node = res.value;
      assert.isTrue(
        node &&
          node.data &&
          node.data.type === "singleSelect" &&
          node.data.staticOptions.length === 10,
        "option item count check"
      );
      if (node && node.data && node.data.type === "singleSelect") {
        const options = (node.data as SingleSelectQuestion).staticOptions as OptionItem[];
        assert.deepEqual(
          options,
          [
            NotificationOptionItem,
            CommandAndResponseOptionItem,
            TabNonSsoItem,
            BotNewUIOptionItem,
            MessageExtensionNewUIItem,
            AzureResourceFunctionNewUI,
            AzureResourceApimNewUI,
            AzureResourceSQLNewUI,
            AzureResourceKeyVaultNewUI,
            CicdOptionItem,
          ],
          "option item should match"
        );
      }
    }
  });

  it("getQuestionsForUserTask - addFeature: SPFx can add CICD", async () => {
    sandbox.stub(featureFlags, "isPreviewFeaturesEnabled").returns(true);
    sandbox.stub(featureFlags, "isBotNotificationEnabled").returns(true);
    sandbox.stub<any, any>(tool, "canAddCICDWorkflows").resolves(true);
    const spfxProjectSettings = {
      appName: "my app",
      projectId: uuid.v4(),
      solutionSettings: {
        hostType: HostTypeOptionSPFx.id,
        name: "test",
        version: "1.0",
        activeResourcePlugins: [
          "fx-resource-spfx",
          "fx-resource-local-debug",
          "fx-resource-appstudio",
          "fx-resource-cicd",
        ],
        capabilities: [TabNewUIOptionItem.id],
        azureResources: [],
      },
    };
    const mockedCtx = new MockedV2Context(spfxProjectSettings);
    const mockedInputs: Inputs = {
      platform: Platform.VSCode,
      stage: Stage.addFeature,
      projectPath: "test path",
    };
    const func: Func = {
      method: "addFeature",
      namespace: "fx-solution-azure",
    };
    const res = await getQuestionsForUserTask(
      mockedCtx,
      mockedInputs,
      func,
      envInfo,
      mockedProvider
    );
    assert.isTrue(res.isOk() && res.value && res.value.data !== undefined);
    if (res.isOk()) {
      const node = res.value;
      assert.isTrue(
        node && node.data && node.data.type === "singleSelect",
        "result should be singleSelect"
      );
      if (node && node.data && node.data.type === "singleSelect") {
        const options = (node.data as SingleSelectQuestion).staticOptions as OptionItem[];
        assert.deepEqual(options, [CicdOptionItem], "option item should match");
      }
    }
  });

  it("getQuestionsForUserTask - addFeature: message extension", async () => {
    sandbox.stub(featureFlags, "isPreviewFeaturesEnabled").returns(true);
    sandbox.stub(featureFlags, "isBotNotificationEnabled").returns(true);
    sandbox.stub<any, any>(tool, "canAddCICDWorkflows").resolves(true);
    const projectSettingsWithNotification = {
      appName: "my app",
      projectId: uuid.v4(),
      solutionSettings: {
        hostType: HostTypeOptionAzure.id,
        name: "test",
        version: "1.0",
        activeResourcePlugins: [
          "fx-resource-frontend-hosting",
          "fx-resource-aad-app-for-teams",
          BuiltInFeaturePluginNames.bot,
        ],
        capabilities: [BotOptionItem.id],
        azureResources: [],
      },
      pluginSettings: {
        [BuiltInFeaturePluginNames.bot]: {
          [PluginBot.BOT_CAPABILITIES]: [BotCapabilities.NOTIFICATION],
          [PluginBot.HOST_TYPE]: BotHostTypes.AzureFunctions,
        },
      },
    };
    const mockedCtx = new MockedV2Context(projectSettingsWithNotification);
    const mockedInputs: Inputs = {
      platform: Platform.VSCode,
      stage: Stage.addFeature,
      projectPath: "test path",
    };
    const func: Func = {
      method: "addFeature",
      namespace: "fx-solution-azure",
    };
    const appStudioPlugin = Container.get<AppStudioPluginV3>(BuiltInFeaturePluginNames.appStudio);
    sandbox
      .stub<any, any>(appStudioPlugin, "capabilityExceedLimit")
      .callsFake(
        async (
          ctx: v2.Context,
          inputs: v2.InputsWithProjectPath,
          capability: "staticTab" | "configurableTab" | "Bot" | "MessageExtension"
        ) => {
          if (capability === "Bot") {
            return ok(true);
          } else {
            return ok(false);
          }
        }
      );
    (mockedCtx.projectSetting.solutionSettings as AzureSolutionSettings).hostType =
      HostTypeOptionAzure.id;
    const res = await getQuestionsForUserTask(
      mockedCtx,
      mockedInputs,
      func,
      envInfo,
      mockedProvider
    );
    assert.isTrue(res.isOk() && res.value && res.value.data !== undefined);
    if (res.isOk()) {
      const node = res.value;
      assert.isTrue(
        node && node.data && node.data.type === "singleSelect",
        "result should be singleSelect"
      );
      if (node && node.data && node.data.type === "singleSelect") {
        const options = (node.data as SingleSelectQuestion).staticOptions as OptionItem[];
        assert.deepEqual(
          options,
          [
            TabNewUIOptionItem,
            TabNonSsoItem,
            AzureResourceFunctionNewUI,
            AzureResourceApimNewUI,
            AzureResourceSQLNewUI,
            AzureResourceKeyVaultNewUI,
            ApiConnectionOptionItem,
            CicdOptionItem,
          ],
          "option item should match"
        );
      }
    }
  });
  it("getQuestionsForUserTask - addFeature: can add message extension for legacy bot", async () => {
    sandbox.stub(featureFlags, "isPreviewFeaturesEnabled").returns(true);
    sandbox.stub(featureFlags, "isBotNotificationEnabled").returns(true);
    sandbox.stub<any, any>(tool, "canAddCICDWorkflows").resolves(true);
    const projectSettingsWithNotification = {
      appName: "my app",
      projectId: uuid.v4(),
      solutionSettings: {
        hostType: HostTypeOptionAzure.id,
        name: "test",
        version: "1.0",
        activeResourcePlugins: [
          "fx-resource-frontend-hosting",
          "fx-resource-aad-app-for-teams",
          BuiltInFeaturePluginNames.bot,
        ],
        capabilities: [BotOptionItem.id],
        azureResources: [],
      },
      pluginSettings: {
        [BuiltInFeaturePluginNames.bot]: {
          [PluginBot.BOT_CAPABILITIES]: [],
          [PluginBot.HOST_TYPE]: BotHostTypes.AzureFunctions,
        },
      },
    };
    const mockedCtx = new MockedV2Context(projectSettingsWithNotification);
    const mockedInputs: Inputs = {
      platform: Platform.VSCode,
      stage: Stage.addFeature,
      projectPath: "test path",
    };
    const func: Func = {
      method: "addFeature",
      namespace: "fx-solution-azure",
    };
    const appStudioPlugin = Container.get<AppStudioPluginV3>(BuiltInFeaturePluginNames.appStudio);
    sandbox
      .stub<any, any>(appStudioPlugin, "capabilityExceedLimit")
      .callsFake(
        async (
          ctx: v2.Context,
          inputs: v2.InputsWithProjectPath,
          capability: "staticTab" | "configurableTab" | "Bot" | "MessageExtension"
        ) => {
          if (capability === "Bot") {
            return ok(true);
          } else {
            return ok(false);
          }
        }
      );
    (mockedCtx.projectSetting.solutionSettings as AzureSolutionSettings).hostType =
      HostTypeOptionAzure.id;
    const res = await getQuestionsForUserTask(
      mockedCtx,
      mockedInputs,
      func,
      envInfo,
      mockedProvider
    );
    assert.isTrue(res.isOk() && res.value && res.value.data !== undefined);
    if (res.isOk()) {
      const node = res.value;
      assert.isTrue(
        node && node.data && node.data.type === "singleSelect",
        "result should be singleSelect"
      );
      if (node && node.data && node.data.type === "singleSelect") {
        const options = (node.data as SingleSelectQuestion).staticOptions as OptionItem[];
        assert.deepEqual(
          options,
          [
            TabNewUIOptionItem,
            TabNonSsoItem,
            MessageExtensionNewUIItem,
            AzureResourceFunctionNewUI,
            AzureResourceApimNewUI,
            AzureResourceSQLNewUI,
            AzureResourceKeyVaultNewUI,
            ApiConnectionOptionItem,
            CicdOptionItem,
          ],
          "option item should match"
        );
      }
    }
  });
});
