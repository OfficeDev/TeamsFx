import "mocha";
import * as chai from "chai";
import * as fs from "fs-extra";
import * as os from "os";
import * as path from "path";
import { ConfigFolderName, LocalSettings, LocalSettingsProvider } from "@microsoft/teamsfx-api";
import {
  LocalSettingsAuthKeys,
  LocalSettingsBackendKeys,
  LocalSettingsBotKeys,
  LocalSettingsFrontendKeys,
  LocalSettingsTeamsAppKeys,
} from "../../src/common/localSettingsConstants";
import {
  LocalEnvBackendKeys,
  LocalEnvFrontendKeys,
} from "../../src/plugins/resource/localdebug/constants";

describe("LocalSettings provider APIs", () => {
  const workspaceFolder = path.resolve(__dirname, "./data/");
  const testFilePath = path.resolve(__dirname, `./data/.${ConfigFolderName}/localSettings.json`);

  let hasFrontend: boolean;
  let hasBackend: boolean;
  let hasBot: boolean;
  let localSettingsProvider: LocalSettingsProvider;

  beforeEach(() => {
    localSettingsProvider = new LocalSettingsProvider(workspaceFolder);
    fs.emptyDirSync(workspaceFolder);
  });

  describe("init localSettings", () => {
    it("should init with tab and backaned", () => {
      hasFrontend = true;
      hasBackend = true;
      hasBot = false;

      const localSettings = localSettingsProvider.init(hasFrontend, hasBackend, hasBot);
      assertLocalSettings(localSettings, hasFrontend, hasBackend, hasBot);
    });

    it("should init with tab and without backaned", () => {
      hasFrontend = true;
      hasBackend = false;
      hasBot = false;

      const localSettings = localSettingsProvider.init(hasFrontend, hasBackend, hasBot);
      assertLocalSettings(localSettings, hasFrontend, hasBackend, hasBot);
    });

    it("should init with bot", () => {
      hasFrontend = false;
      hasBackend = false;
      hasBot = true;

      const localSettings = localSettingsProvider.init(hasFrontend, hasBackend, hasBot);
      assertLocalSettings(localSettings, hasFrontend, hasBackend, hasBot);
    });
  });

  describe("save localSettings", () => {
    it("should create with default settings", async () => {
      hasFrontend = true;
      hasBackend = true;
      hasBot = true;

      const localSettings: LocalSettings = localSettingsProvider.init(
        hasFrontend,
        hasBackend,
        hasBot
      );
      await localSettingsProvider.save(localSettings);

      chai.assert.isTrue(await fs.pathExists(testFilePath));
      const expectedContent = JSON.stringify(localSettings, null, 4);
      const actualContent = await fs.readFile(testFilePath, "utf8");
      chai.assert.equal(actualContent, expectedContent);
    });
  });

  describe("load localSettings", () => {
    it("should load after save", async () => {
      const localSettings = localSettingsProvider.init(true, true, true);
      const updateValue = "http://localhost:5000";
      localSettings.auth?.set(LocalSettingsAuthKeys.SimpleAuthServiceEndpoint, updateValue);

      await localSettingsProvider.save(localSettings);
      const updatedLocalSettings = await localSettingsProvider.load();

      assertLocalSettings(updatedLocalSettings, true, true, true);
      chai.assert.equal(
        updatedLocalSettings!.auth?.get(LocalSettingsAuthKeys.SimpleAuthServiceEndpoint),
        updateValue
      );
    });

    it("should return undefined if file doesn't exist", async () => {
      const localSettings = await localSettingsProvider.load();
      chai.assert.isUndefined(localSettings);
    });
  });

  function assertLocalSettings(
    localSettings: LocalSettings | undefined,
    hasFrontend: boolean,
    hasBackend: boolean,
    hasBot: boolean
  ) {
    chai.assert.isDefined(localSettings);

    // Teams app settings is always required.
    chai.assert.isDefined(localSettings!.teamsApp);

    const expectedTeamsAppKeys = Object.values(LocalSettingsTeamsAppKeys);
    for (const key of expectedTeamsAppKeys) {
      chai.assert.isTrue(localSettings!.teamsApp?.has(key));
    }

    // Verify frontend related settings.
    if (hasFrontend) {
      chai.assert.isDefined(localSettings!.frontend);
      chai.assert.isDefined(localSettings!.auth);

      const expectedTeamsAppKeys = Object.values(LocalSettingsTeamsAppKeys);
      const expectedFrontendKeys = Object.values(LocalSettingsFrontendKeys);
      const expectedAuthKeys = Object.values(LocalSettingsAuthKeys);

      for (const key of expectedTeamsAppKeys) {
        chai.assert.isTrue(localSettings!.teamsApp?.has(key));
      }

      for (const key of expectedAuthKeys) {
        chai.assert.isTrue(localSettings!.auth?.has(key));
      }

      for (const key of expectedFrontendKeys) {
        chai.assert.isTrue(localSettings?.frontend?.has(key));
      }
    }

    // Verify backend related settings.
    if (hasBackend) {
      chai.assert.isDefined(localSettings!.backend);

      const expectedBackendKeys = Object.values(LocalSettingsBackendKeys);
      for (const key of expectedBackendKeys) {
        chai.assert.isTrue(localSettings!.backend?.has(key));
      }
    }

    // Verify bot related settings.
    if (hasBot) {
      chai.assert.isDefined(localSettings!.bot);

      const expectedBotKeys = Object.values(LocalSettingsBotKeys);
      for (const key of expectedBotKeys) {
        chai.assert.isTrue(localSettings!.bot?.has(key));
      }
    }
  }
});
