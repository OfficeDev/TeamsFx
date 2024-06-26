import { err, ok, UserError } from "@microsoft/teamsfx-api";
import { ProgressHandler } from "@microsoft/vscode-ui";
import * as sinon from "sinon";
import VsCodeLogInstance from "../../src/commonlib/log";
import * as errorCommon from "../../src/error/common";
import {
  migrateTeamsManifestHandler,
  migrateTeamsTabAppHandler,
} from "../../src/handlers/migrationHandler";
import { TeamsAppMigrationHandler } from "../../src/migration/migrationHandler";
import * as vsc_ui from "../../src/qm/vsc_ui";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import * as localizeUtils from "../../src/utils/localizeUtils";

describe("Migration handlers", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("migrateTeamsTabAppHandler", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("happy path", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent").returns();
      sandbox.stub(localizeUtils, "localize").callsFake((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        showMessage: () => Promise.resolve(ok("teamstoolkit.migrateTeamsTabApp.upgrade")),
        selectFolder: () => Promise.resolve(ok({ type: "success", result: "test" })),
        createProgressBar: () => progressHandler,
      });
      sandbox.stub(VsCodeLogInstance, "info").returns();
      sandbox.stub(TeamsAppMigrationHandler.prototype, "updatePackageJson").resolves(ok(true));
      sandbox.stub(TeamsAppMigrationHandler.prototype, "updateCodes").resolves(ok([]));

      const result = await migrateTeamsTabAppHandler();

      chai.assert.deepEqual(result, ok(null));
    });

    it("happy path: failed files", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent").returns();
      sandbox.stub(localizeUtils, "localize").callsFake((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        showMessage: () => Promise.resolve(ok("teamstoolkit.migrateTeamsTabApp.upgrade")),
        selectFolder: () => Promise.resolve(ok({ type: "success", result: "test" })),
        createProgressBar: () => progressHandler,
      });
      sandbox.stub(VsCodeLogInstance, "info").returns();
      const warningStub = sandbox.stub(VsCodeLogInstance, "warning");
      sandbox.stub(TeamsAppMigrationHandler.prototype, "updatePackageJson").resolves(ok(true));
      sandbox
        .stub(TeamsAppMigrationHandler.prototype, "updateCodes")
        .resolves(ok(["test1", "test2"]));

      const result = await migrateTeamsTabAppHandler();

      chai.assert.deepEqual(result, ok(null));
      chai.expect(warningStub.calledOnce).to.be.true;
    });

    it("error", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent").returns();
      const sendTelemetryErrorEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sandbox.stub(localizeUtils, "localize").callsFake((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        showMessage: () => Promise.resolve(ok("teamstoolkit.migrateTeamsTabApp.upgrade")),
        selectFolder: () => Promise.resolve(ok({ type: "success", result: "test" })),
        createProgressBar: () => progressHandler,
      });
      sandbox.stub(VsCodeLogInstance, "info").returns();
      sandbox.stub(TeamsAppMigrationHandler.prototype, "updatePackageJson").resolves(ok(true));
      sandbox
        .stub(TeamsAppMigrationHandler.prototype, "updateCodes")
        .resolves(err({ foo: "bar" } as any));

      const result = await migrateTeamsTabAppHandler();

      chai.assert.isTrue(result.isErr());
      chai.expect(sendTelemetryErrorEventStub.calledOnce).to.be.true;
    });

    it("user cancel", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent").returns();
      sandbox.stub(localizeUtils, "localize").callsFake((key: string) => key);
      const sendTelemetryErrorEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        showMessage: () => Promise.resolve(ok("teamstoolkit.migrateTeamsTabApp.upgrade")),
        selectFolder: () => Promise.resolve(ok({ type: "skip" })),
      });

      const result = await migrateTeamsTabAppHandler();

      chai.assert.deepEqual(result, ok(null));
      chai.expect(sendTelemetryErrorEventStub.calledOnce).to.be.true;
    });

    it("user cancel: skip folder selection", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent").returns();
      sandbox.stub(localizeUtils, "localize").callsFake((key: string) => key);
      const sendTelemetryErrorEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        showMessage: () => Promise.resolve(ok("cancel")),
      });

      const result = await migrateTeamsTabAppHandler();

      chai.assert.deepEqual(result, ok(null));
      chai.expect(sendTelemetryErrorEventStub.calledOnce).to.be.true;
    });

    it("no change in package.json", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent").returns();
      sandbox.stub(localizeUtils, "localize").callsFake((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        showMessage: () => Promise.resolve(ok("teamstoolkit.migrateTeamsTabApp.upgrade")),
        selectFolder: () => Promise.resolve(ok({ type: "success", result: "test" })),
        createProgressBar: () => progressHandler,
      });
      sandbox.stub(VsCodeLogInstance, "info").returns();
      sandbox.stub(VsCodeLogInstance, "warning").returns();
      sandbox.stub(TeamsAppMigrationHandler.prototype, "updatePackageJson").resolves(ok(false));

      const result = await migrateTeamsTabAppHandler();

      chai.assert.deepEqual(result, ok(null));
    });
  });

  describe("migrateTeamsManifestHandler", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("happy path", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent").returns();
      sandbox.stub(localizeUtils, "localize").callsFake((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        showMessage: () => Promise.resolve(ok("teamstoolkit.migrateTeamsManifest.upgrade")),
        selectFile: () => Promise.resolve(ok({ type: "success", result: "test" })),
        createProgressBar: () => progressHandler,
      });
      sandbox.stub(VsCodeLogInstance, "info").returns();
      sandbox.stub(TeamsAppMigrationHandler.prototype, "updateManifest").resolves(ok(null));

      const result = await migrateTeamsManifestHandler();

      chai.assert.deepEqual(result, ok(null));
    });

    it("user cancel: skip file selection", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent").returns();
      const sendTelemetryErrorEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sandbox.stub(localizeUtils, "localize").callsFake((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        showMessage: () => Promise.resolve(ok("teamstoolkit.migrateTeamsManifest.upgrade")),
        selectFile: () => Promise.resolve(ok({ type: "skip" })),
        createProgressBar: () => progressHandler,
      });
      sandbox.stub(VsCodeLogInstance, "info").returns();
      sandbox.stub(TeamsAppMigrationHandler.prototype, "updateManifest").resolves(ok(null));

      const result = await migrateTeamsManifestHandler();

      chai.assert.deepEqual(result, ok(null));
      chai.expect(sendTelemetryErrorEventStub.calledOnce).to.be.true;
    });

    it("error", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent").returns();
      sandbox.stub(localizeUtils, "localize").callsFake((key: string) => key);
      const sendTelemetryErrorEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const progressHandler = new ProgressHandler("title", 1);
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        showMessage: () => Promise.resolve(ok("teamstoolkit.migrateTeamsManifest.upgrade")),
        selectFile: () => Promise.resolve(ok({ type: "success", result: "test" })),
        createProgressBar: () => progressHandler,
      });
      sandbox.stub(VsCodeLogInstance, "info").returns();
      sandbox
        .stub(TeamsAppMigrationHandler.prototype, "updateManifest")
        .resolves(err(new UserError("source", "name", "")));
      sandbox.stub(errorCommon, "showError").callsFake(async () => {});

      const result = await migrateTeamsManifestHandler();

      chai.assert.isTrue(result.isErr());
      chai.expect(sendTelemetryErrorEventStub.calledOnce).to.be.true;
    });
  });
});
