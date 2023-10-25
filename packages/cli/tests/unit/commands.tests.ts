import { CLIContext, err, ok } from "@microsoft/teamsfx-api";
import {
  CapabilityOptions,
  CollaborationStateResult,
  FuncToolChecker,
  FxCore,
  ListCollaboratorResult,
  LocalCertificateManager,
  LtsNodeChecker,
  PackageService,
  PermissionsResult,
  QuestionNames,
  UserCancelError,
  envUtil,
} from "@microsoft/teamsfx-core";
import { assert } from "chai";
import "mocha";
import * as sinon from "sinon";
import * as activate from "../../src/activate";
import * as accountUtils from "../../src/cmds/account";
import * as m365 from "../../src/cmds/m365/m365";
import { localTelemetryReporter } from "../../src/cmds/preview/localTelemetryReporter";
import {
  accountLoginAzureCommand,
  accountLoginM365Command,
  accountLogoutCommand,
  accountShowCommand,
  addSPFxWebpartCommand,
  configGetCommand,
  configSetCommand,
  getCreateCommand,
  createSampleCommand,
  deployCommand,
  envAddCommand,
  envListCommand,
  listTemplatesCommand,
  listSamplesCommand,
  m365LaunchInfoCommand,
  m365SideloadingCommand,
  m365UnacquireCommand,
  packageCommand,
  permissionGrantCommand,
  permissionStatusCommand,
  previewCommand,
  printGlobalConfig,
  provisionCommand,
  publishCommand,
  updateAadAppCommand,
  updateTeamsAppCommand,
  upgradeCommand,
  validateCommand,
  helpCommand,
} from "../../src/commands/models";
import AzureTokenProvider from "../../src/commonlib/azureLogin";
import * as codeFlowLogin from "../../src/commonlib/codeFlowLogin";
import { signedIn, signedOut } from "../../src/commonlib/common/constant";
import { logger } from "../../src/commonlib/logger";
import M365TokenProvider from "../../src/commonlib/m365Login";
import { UserSettings } from "../../src/userSetttings";
import * as utils from "../../src/utils";
import { MissingRequiredOptionError } from "../../src/error";
import mockedEnv, { RestoreFn } from "mocked-env";
import { teamsappUpdateCommand } from "../../src/commands/models/teamsapp/update";
import { teamsappPackageCommand } from "../../src/commands/models/teamsapp/package";
import { teamsappValidateCommand } from "../../src/commands/models/teamsapp/validate";
import { teamsappPublishCommand } from "../../src/commands/models/teamsapp/publish";
import { DoctorChecker, teamsappDoctorCommand } from "../../src/commands/models/teamsapp/doctor";
import M365TokenInstance from "../../src/commonlib/m365Login";
import * as tools from "@microsoft/teamsfx-core/build/common/tools";

describe("CLI commands", () => {
  const sandbox = sinon.createSandbox();

  let mockedEnvRestore: RestoreFn;

  beforeEach(() => {
    sandbox.stub(logger, "info").resolves(true);
    sandbox.stub(logger, "error").resolves(true);
  });

  afterEach(() => {
    sandbox.restore();
    if (mockedEnvRestore) {
      mockedEnvRestore();
    }
  });

  describe("getCreateCommand", async () => {
    it("happy path", async () => {
      mockedEnvRestore = mockedEnv({
        COPILOT_PLUGIN: "false",
      });
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "createProject").resolves(ok({ projectPath: "..." }));

      const ctx: CLIContext = {
        command: { ...getCreateCommand(), fullName: "teamsfx new" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };

      const copilotPluginQuestionNames = [QuestionNames.OpenAIPluginManifest.toString()];
      assert.isTrue(
        ctx.command.options?.filter((o) => copilotPluginQuestionNames.includes(o.name)).length === 0
      );
      const res = await getCreateCommand().handler!(ctx);
      assert.isTrue(res.isOk());
    });

    it("createProjectOptions - API copilot plugin enabled", async () => {
      mockedEnvRestore = mockedEnv({
        COPILOT_PLUGIN: "true",
        API_COPILOT_PLUGIN: "true",
      });
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "createProject").resolves(ok({ projectPath: "..." }));
      const ctx: CLIContext = {
        command: { ...getCreateCommand(), fullName: "teamsfx new" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await getCreateCommand().handler!(ctx);
      const copilotPluginQuestionNames = [QuestionNames.OpenAIPluginManifest.toString()];
      assert.isTrue(
        ctx.command.options?.filter((o) => copilotPluginQuestionNames.includes(o.name)).length === 1
      );
      assert.isTrue(res.isOk());
    });

    it("createProjectOptions - API copilot plugin disabled but bot Copilot plugin enabled", async () => {
      mockedEnvRestore = mockedEnv({
        COPILOT_PLUGIN: "true",
        API_COPILOT_PLUGIN: "false",
      });
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "createProject").resolves(ok({ projectPath: "..." }));

      const ctx: CLIContext = {
        command: { ...getCreateCommand(), fullName: "teamsfx new" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };

      const copilotPluginQuestionNames = [QuestionNames.OpenAIPluginManifest.toString()];
      assert.isTrue(
        ctx.command.options?.filter((o) => copilotPluginQuestionNames.includes(o.name)).length === 0
      );
      const res = await getCreateCommand().handler!(ctx);
      assert.isTrue(res.isOk());
    });

    it("core return error", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "createProject").resolves(err(new UserCancelError()));
      const ctx: CLIContext = {
        command: { ...getCreateCommand(), fullName: "teamsfx new" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await getCreateCommand().handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });

  describe("createSampleCommand", async () => {
    it("happy path", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "createSampleProject").resolves(ok({ projectPath: "..." }));
      const ctx: CLIContext = {
        command: { ...createSampleCommand, fullName: "teamsfx new sample" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await createSampleCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("core return error", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "createProject").resolves(err(new UserCancelError()));
      const ctx: CLIContext = {
        command: { ...createSampleCommand, fullName: "teamsfx new sample" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await createSampleCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });
  describe("listSampleCommand", async () => {
    it("happy path", async () => {
      sandbox.stub(utils, "getTemplates").resolves([]);
      const ctx: CLIContext = {
        command: { ...listSamplesCommand, fullName: "teamsfx list samples" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await listSamplesCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("accountLoginAzureCommand", async () => {
    it("should success when service-principal = false", async () => {
      sandbox.stub(AzureTokenProvider, "signout");
      sandbox.stub(accountUtils, "outputAzureInfo").resolves();
      const ctx: CLIContext = {
        command: { ...accountLoginAzureCommand, fullName: "teamsapp auth login azure" },
        optionValues: { "service-principal": false },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountLoginAzureCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("should fail when service-principal = true", async () => {
      sandbox.stub(AzureTokenProvider, "signout");
      sandbox.stub(accountUtils, "outputAzureInfo").resolves();
      const ctx: CLIContext = {
        command: { ...accountLoginAzureCommand, fullName: "teamsapp auth login azure" },
        optionValues: { "service-principal": true },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountLoginAzureCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
    it("should fail service-principal = false", async () => {
      sandbox.stub(AzureTokenProvider, "signout");
      sandbox.stub(accountUtils, "outputAzureInfo").resolves();
      const ctx: CLIContext = {
        command: { ...accountLoginAzureCommand, fullName: "teamsapp auth login azure" },
        optionValues: { "service-principal": false, username: "abc" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountLoginAzureCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });
  describe("accountLoginM365Command", async () => {
    it("should success", async () => {
      sandbox.stub(M365TokenProvider, "signout");
      sandbox.stub(accountUtils, "outputM365Info").resolves();
      const ctx: CLIContext = {
        command: { ...accountLoginM365Command, fullName: "teamsapp auth login m365" },
        optionValues: { "service-principal": false },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountLoginM365Command.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });

  describe("addSPFxWebpartCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "addWebpart").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...addSPFxWebpartCommand, fullName: "teamsfx add spfx-web-part" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await addSPFxWebpartCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("deployCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "deployArtifacts").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...deployCommand, fullName: "teamsfx" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await deployCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("envAddCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "createEnv").resolves(ok(undefined));
      sandbox.stub(utils, "isWorkspaceSupported").returns(true);
      const ctx: CLIContext = {
        command: { ...envAddCommand, fullName: "teamsfx" },
        optionValues: { projectPath: "." },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await envAddCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("isWorkspaceSupported: false", async () => {
      sandbox.stub(FxCore.prototype, "createEnv").resolves(ok(undefined));
      sandbox.stub(utils, "isWorkspaceSupported").returns(false);
      const ctx: CLIContext = {
        command: { ...envAddCommand, fullName: "teamsfx" },
        optionValues: { projectPath: "." },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await envAddCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });
  describe("envListCommand", async () => {
    it("success", async () => {
      sandbox.stub(utils, "isWorkspaceSupported").returns(true);
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev"]));
      const ctx: CLIContext = {
        command: { ...envListCommand, fullName: "teamsfx" },
        optionValues: { projectPath: "." },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await envListCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("isWorkspaceSupported: false", async () => {
      sandbox.stub(utils, "isWorkspaceSupported").returns(false);
      const ctx: CLIContext = {
        command: { ...envListCommand, fullName: "teamsfx" },
        optionValues: { projectPath: "." },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await envListCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
    it("listEnv error", async () => {
      sandbox.stub(utils, "isWorkspaceSupported").returns(true);
      sandbox.stub(envUtil, "listEnv").resolves(err(new UserCancelError()));
      const ctx: CLIContext = {
        command: { ...envListCommand, fullName: "teamsfx" },
        optionValues: { projectPath: "." },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await envListCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });
  describe("provisionCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "provisionResources").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...provisionCommand, fullName: "teamsfx" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await provisionCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("non interactive mode", async () => {
      sandbox.stub(FxCore.prototype, "provisionResources").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...provisionCommand, fullName: "teamsfx" },
        optionValues: { nonInteractive: true, region: "East US" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await provisionCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("packageCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "createAppPackage").resolves(ok({ state: "OK" }));
      const ctx: CLIContext = {
        command: { ...packageCommand, fullName: "teamsfx" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await packageCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("permissionGrantCommand", async () => {
    it("success interactive = false", async () => {
      sandbox
        .stub(FxCore.prototype, "grantPermission")
        .resolves(ok({ state: "OK" } as PermissionsResult));
      const ctx: CLIContext = {
        command: { ...permissionGrantCommand, fullName: "teamsfx" },
        optionValues: { "manifest-file-path": "abc" },
        globalOptionValues: { interactive: false },
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await permissionGrantCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("success interactive = true", async () => {
      sandbox
        .stub(FxCore.prototype, "grantPermission")
        .resolves(ok({ state: "OK" } as PermissionsResult));
      const ctx: CLIContext = {
        command: { ...permissionGrantCommand, fullName: "teamsfx" },
        optionValues: {},
        globalOptionValues: { interactive: true },
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await permissionGrantCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("missing option", async () => {
      sandbox
        .stub(FxCore.prototype, "grantPermission")
        .resolves(ok({ state: "OK" } as PermissionsResult));
      const ctx: CLIContext = {
        command: { ...permissionGrantCommand, fullName: "teamsfx" },
        optionValues: {},
        globalOptionValues: { interactive: false },
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await permissionGrantCommand.handler!(ctx);
      assert.isTrue(res.isErr() && res.error instanceof MissingRequiredOptionError);
    });
  });
  describe("permissionStatusCommand", async () => {
    it("listCollaborator", async () => {
      sandbox
        .stub(FxCore.prototype, "listCollaborator")
        .resolves(ok({ state: "OK" } as ListCollaboratorResult));
      const ctx: CLIContext = {
        command: { ...permissionStatusCommand, fullName: "teamsfx" },
        optionValues: { all: true },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await permissionStatusCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("checkPermission", async () => {
      sandbox
        .stub(FxCore.prototype, "checkPermission")
        .resolves(ok({ state: "OK" } as CollaborationStateResult));
      const ctx: CLIContext = {
        command: { ...permissionStatusCommand, fullName: "teamsfx" },
        optionValues: { all: false },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await permissionStatusCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("publishCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "publishApplication").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...publishCommand, fullName: "teamsfx" },
        optionValues: { env: "local" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await publishCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("previewCommand", async () => {
    it("success", async () => {
      sandbox.stub(localTelemetryReporter, "runWithTelemetryGeneric").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...previewCommand, fullName: "teamsfx" },
        optionValues: { env: "local" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await previewCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("error", async () => {
      sandbox
        .stub(localTelemetryReporter, "runWithTelemetryGeneric")
        .resolves(err(new UserCancelError()));
      const ctx: CLIContext = {
        command: { ...previewCommand, fullName: "teamsfx" },
        optionValues: { env: "local" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await previewCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });
  describe("updateAadAppCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "deployAadManifest").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...updateAadAppCommand, fullName: "teamsfx" },
        optionValues: {
          env: "local",
          projectPath: "./",
          "manifest-file-path": "./aad.manifest.json",
        },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await updateAadAppCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("updateTeamsAppCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "deployTeamsManifest").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...updateTeamsAppCommand, fullName: "teamsfx" },
        optionValues: { env: "local" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await updateTeamsAppCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });

    it("MissingRequiredOptionError", async () => {
      sandbox.stub(FxCore.prototype, "deployTeamsManifest").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...updateTeamsAppCommand, fullName: "teamsfx" },
        optionValues: { "manifest-path": "fakePath", projectPath: "./" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await updateTeamsAppCommand.handler!(ctx);
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, MissingRequiredOptionError.name);
      }
    });
  });
  describe("upgradeCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "phantomMigrationV3").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...upgradeCommand, fullName: "teamsfx" },
        optionValues: { force: true },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await upgradeCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("validateCommand", async () => {
    it("conflict", async () => {
      sandbox.stub(FxCore.prototype, "validateApplication").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...validateCommand, fullName: "teamsfx" },
        optionValues: { "manifest-path": "aaa", "app-package-file-path": "bbb" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await validateCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
    it("none", async () => {
      sandbox.stub(FxCore.prototype, "validateApplication").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...validateCommand, fullName: "teamsfx" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await validateCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
    it("manifest", async () => {
      sandbox.stub(FxCore.prototype, "validateApplication").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...validateCommand, fullName: "teamsfx" },
        optionValues: { "manifest-path": "aaa", env: "dev" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await validateCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("manifest missing env", async () => {
      sandbox.stub(FxCore.prototype, "validateApplication").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...validateCommand, fullName: "teamsfx" },
        optionValues: { "manifest-path": "aaa" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await validateCommand.handler!(ctx);
      assert.isTrue(res.isErr() && res.error instanceof MissingRequiredOptionError);
    });
    it("package", async () => {
      sandbox.stub(FxCore.prototype, "validateApplication").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...validateCommand, fullName: "teamsfx" },
        optionValues: { "app-package-file-path": "bbb" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await validateCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });

  describe("m365LaunchInfoCommand", async () => {
    beforeEach(() => {
      sandbox.stub(logger, "warning");
    });
    it("success retrieveTitleId", async () => {
      sandbox.stub(m365, "getTokenAndUpn").resolves(["token", "upn"]);
      sandbox.stub(PackageService.prototype, "retrieveTitleId").resolves("id");
      sandbox.stub(PackageService.prototype, "getLaunchInfoByTitleId").resolves("id");
      const ctx: CLIContext = {
        command: { ...m365LaunchInfoCommand, fullName: "teamsfx" },
        optionValues: { "manifest-id": "aaa" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365LaunchInfoCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("success", async () => {
      sandbox.stub(m365, "getTokenAndUpn").resolves(["token", "upn"]);
      sandbox.stub(PackageService.prototype, "getLaunchInfoByTitleId").resolves("id");
      const ctx: CLIContext = {
        command: { ...m365LaunchInfoCommand, fullName: "teamsfx" },
        optionValues: { "title-id": "aaa" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365LaunchInfoCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("MissingRequiredOptionError", async () => {
      sandbox.stub(m365, "getTokenAndUpn").resolves(["token", "upn"]);
      const ctx: CLIContext = {
        command: { ...m365LaunchInfoCommand, fullName: "teamsfx" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365LaunchInfoCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });

  describe("m365SideloadingCommand", async () => {
    beforeEach(() => {
      sandbox.stub(logger, "warning");
    });
    it("should success with zip package", async () => {
      sandbox.stub(m365, "getTokenAndUpn").resolves(["token", "upn"]);
      sandbox.stub(PackageService.prototype, "sideLoading").resolves();
      const ctx: CLIContext = {
        command: { ...m365SideloadingCommand, fullName: "teamsfx" },
        optionValues: { "manifest-id": "aaa", "file-path": "./" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365SideloadingCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("should success with xml", async () => {
      sandbox.stub(m365, "getTokenAndUpn").resolves(["token", "upn"]);
      sandbox.stub(PackageService.prototype, "sideLoadXmlManifest").resolves();
      const ctx: CLIContext = {
        command: { ...m365SideloadingCommand, fullName: "teamsfx" },
        optionValues: { "manifest-id": "aaa", "xml-path": "./" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365SideloadingCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("should fail if both zip and xml are provided", async () => {
      const ctx: CLIContext = {
        command: { ...m365SideloadingCommand, fullName: "teamsfx" },
        optionValues: { "manifest-id": "aaa", "xml-path": "./", "file-path": "./" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365SideloadingCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
    it("should fail if non of zip and xml are provided", async () => {
      const ctx: CLIContext = {
        command: { ...m365SideloadingCommand, fullName: "teamsfx" },
        optionValues: { "manifest-id": "aaa" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365SideloadingCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });

  describe("m365UnacquireCommand", async () => {
    beforeEach(() => {
      sandbox.stub(logger, "warning");
    });
    it("MissingRequiredOptionError", async () => {
      sandbox.stub(m365, "getTokenAndUpn").resolves(["token", "upn"]);
      const ctx: CLIContext = {
        command: { ...m365UnacquireCommand, fullName: "teamsfx" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365UnacquireCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
    it("success retrieveTitleId", async () => {
      sandbox.stub(m365, "getTokenAndUpn").resolves(["token", "upn"]);
      sandbox.stub(PackageService.prototype, "retrieveTitleId").resolves("id");
      sandbox.stub(PackageService.prototype, "unacquire").resolves();
      const ctx: CLIContext = {
        command: { ...m365UnacquireCommand, fullName: "teamsfx" },
        optionValues: { "manifest-id": "aaa" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365UnacquireCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("success", async () => {
      sandbox.stub(m365, "getTokenAndUpn").resolves(["token", "upn"]);
      sandbox.stub(PackageService.prototype, "unacquire").resolves();
      const ctx: CLIContext = {
        command: { ...m365UnacquireCommand, fullName: "teamsfx" },
        optionValues: { "title-id": "aaa" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365UnacquireCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });

  describe("v3 commands", async () => {
    beforeEach(() => {
      sandbox.stub(logger, "warning");
    });
    afterEach(() => {
      sandbox.restore();
    });
    it("update", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "updateTeamsAppCLIV3").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...teamsappUpdateCommand, fullName: "teamsapp update" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await teamsappUpdateCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("update conflict", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "updateTeamsAppCLIV3").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...teamsappUpdateCommand, fullName: "teamsapp update" },
        optionValues: { "manifest-file": "manifest.json", "package-file": "package.zip" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await teamsappUpdateCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
    it("package", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "packageTeamsAppCLIV3").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...teamsappPackageCommand, fullName: "teamsapp package" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await teamsappPackageCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("validate", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "validateTeamsAppCLIV3").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...teamsappValidateCommand, fullName: "teamsapp validate" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await teamsappValidateCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("validate conflict", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "validateTeamsAppCLIV3").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...teamsappValidateCommand, fullName: "teamsapp validate" },
        optionValues: { "manifest-file": "manifest.json", "package-file": "package.zip" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await teamsappValidateCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
    it("publish", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "publishTeamsAppCLIV3").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...teamsappPublishCommand, fullName: "teamsapp publish" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await teamsappPublishCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("publish conflict", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "publishTeamsAppCLIV3").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...teamsappPublishCommand, fullName: "teamsapp publish" },
        optionValues: { "manifest-file": "manifest.json", "package-file": "package.zip" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await teamsappPublishCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });
});

describe("CLI read-only commands", () => {
  const sandbox = sinon.createSandbox();

  let messages: string[] = [];

  beforeEach(() => {
    sandbox.stub(logger, "info").callsFake(async (message: string) => {
      messages.push(message);
      return true;
    });
    sandbox.stub(logger, "error").callsFake(async (message: string) => {
      messages.push(message);
      return true;
    });
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("accountShowCommand", async () => {
    it("both signedOut", async () => {
      sandbox.stub(M365TokenProvider, "getStatus").resolves(ok({ status: signedOut }));
      sandbox.stub(AzureTokenProvider, "getStatus").resolves({ status: signedOut });
      messages = [];
      const ctx: CLIContext = {
        command: { ...accountShowCommand, fullName: "teamsapp auth show" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountShowCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("both signedIn and checkIsOnline = true", async () => {
      sandbox.stub(M365TokenProvider, "getStatus").resolves(ok({ status: signedIn }));
      sandbox.stub(AzureTokenProvider, "getStatus").resolves({ status: signedIn });
      sandbox.stub(codeFlowLogin, "checkIsOnline").resolves(true);
      const outputM365Info = sandbox.stub(accountUtils, "outputM365Info").resolves();
      const outputAzureInfo = sandbox.stub(accountUtils, "outputAzureInfo").resolves();
      messages = [];
      const ctx: CLIContext = {
        command: { ...accountShowCommand, fullName: "teamsapp auth show" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountShowCommand.handler!(ctx);
      assert.isTrue(res.isOk());
      assert.isTrue(outputM365Info.calledOnce);
      assert.isTrue(outputAzureInfo.calledOnce);
    });
    it("both signedIn and checkIsOnline = false", async () => {
      sandbox
        .stub(M365TokenProvider, "getStatus")
        .resolves(ok({ status: signedIn, accountInfo: { upn: "xxx" } }));
      sandbox
        .stub(AzureTokenProvider, "getStatus")
        .resolves({ status: signedIn, accountInfo: { upn: "xxx" } });
      sandbox.stub(codeFlowLogin, "checkIsOnline").resolves(false);
      const outputAccountInfoOffline = sandbox.stub(accountUtils, "outputAccountInfoOffline");
      messages = [];
      const ctx: CLIContext = {
        command: { ...accountShowCommand, fullName: "teamsapp auth show" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountShowCommand.handler!(ctx);
      assert.isTrue(res.isOk());
      assert.isTrue(outputAccountInfoOffline.calledTwice);
    });
    it("M365TokenProvider.getStatus() returns error", async () => {
      sandbox.stub(M365TokenProvider, "getStatus").resolves(err(new UserCancelError()));
      messages = [];
      const ctx: CLIContext = {
        command: { ...accountShowCommand, fullName: "teamsapp auth show" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountShowCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });

  describe("accountLogoutCommand", async () => {
    it("azure success", async () => {
      sandbox.stub(AzureTokenProvider, "signout").resolves(true);
      const ctx: CLIContext = {
        command: { ...accountLogoutCommand, fullName: "teamsapp auth logout" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: ["azure"],
        telemetryProperties: {},
      };
      messages = [];
      const res = await accountLogoutCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("azure fail", async () => {
      sandbox.stub(AzureTokenProvider, "signout").resolves(false);
      const ctx: CLIContext = {
        command: { ...accountLogoutCommand, fullName: "teamsapp auth logout" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: ["azure"],
        telemetryProperties: {},
      };
      messages = [];
      const res = await accountLogoutCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("m365 success", async () => {
      sandbox.stub(M365TokenProvider, "signout").resolves(true);
      const ctx: CLIContext = {
        command: { ...accountLogoutCommand, fullName: "teamsapp auth logout" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: ["m365"],
        telemetryProperties: {},
      };
      const res = await accountLogoutCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("m365 fail", async () => {
      sandbox.stub(M365TokenProvider, "signout").resolves(false);
      const ctx: CLIContext = {
        command: { ...accountLogoutCommand, fullName: "teamsapp auth logout" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: ["m365"],
        telemetryProperties: {},
      };
      messages = [];
      const res = await accountLogoutCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });

  describe("configGetCommand", async () => {
    it("printGlobalConfig all", async () => {
      sandbox.stub(UserSettings, "getConfigSync").returns(ok({ a: 1, b: 2 }));
      const res = await printGlobalConfig();
      assert.isTrue(res.isOk());
      assert.isTrue(messages.includes(JSON.stringify({ a: 1, b: 2 }, null, 2)));
    });
    it("printGlobalConfig some key", async () => {
      sandbox.stub(UserSettings, "getConfigSync").returns(ok({ a: { c: 3 }, b: 2 }));
      const res = await printGlobalConfig("a");
      assert.isTrue(res.isOk());
      assert.isTrue(messages.includes(JSON.stringify({ c: 3 }, null, 2)));
    });
    it("printGlobalConfig error", async () => {
      sandbox.stub(UserSettings, "getConfigSync").returns(err(new UserCancelError()));
      const res = await printGlobalConfig();
      assert.isTrue(res.isErr());
    });
    it("configGetCommand", async () => {
      sandbox.stub(UserSettings, "getConfigSync").returns(ok({ a: 1, b: 2 }));
      const ctx: CLIContext = {
        command: { ...configGetCommand, fullName: "teamsfx ..." },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: ["a"],
        telemetryProperties: {},
      };
      messages = [];
      const res = await configGetCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });

  describe("configSetCommand", async () => {
    it("configSetCommand", async () => {
      sandbox.stub(UserSettings, "setConfigSync").returns(ok(undefined));
      const ctx: CLIContext = {
        command: { ...configGetCommand, fullName: "teamsfx ..." },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: ["key", "value"],
        telemetryProperties: {},
      };
      const res = await configSetCommand.handler!(ctx);
      assert.isTrue(res.isOk());
      assert.isTrue(messages.includes(`Successfully set user configuration key.`));
    });
    it("configSetCommand error", async () => {
      sandbox.stub(UserSettings, "setConfigSync").returns(err(new UserCancelError()));
      const ctx: CLIContext = {
        command: { ...configGetCommand, fullName: "teamsfx ..." },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: ["key", "value"],
        telemetryProperties: {},
      };
      const res = await configSetCommand.handler!(ctx);
      assert.isTrue(res.isErr());
      assert.isTrue(messages.includes("Set user configuration failed."));
    });
  });

  describe("listTemplatesCommand", async () => {
    let mockedEnvRestore: RestoreFn;
    afterEach(() => {
      if (mockedEnvRestore) {
        mockedEnvRestore();
      }
    });
    it("happy path", async () => {
      const ctx: CLIContext = {
        command: { ...listTemplatesCommand, fullName: "teamsfx list" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await listTemplatesCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("json", async () => {
      mockedEnvRestore = mockedEnv({
        COPILOT_PLUGIN: "false",
      });
      const ctx: CLIContext = {
        command: { ...listTemplatesCommand, fullName: "teamsfx ..." },
        optionValues: { format: "json" },
        globalOptionValues: {},
        argumentValues: ["key", "value"],
        telemetryProperties: {},
      };
      const res = await listTemplatesCommand.handler!(ctx);
      assert.isTrue(res.isOk());
      assert.isFalse(!!messages.find((msg) => msg.includes("copilot-plugin-existing-api")));
    });
    it("table with description", async () => {
      const ctx: CLIContext = {
        command: { ...listTemplatesCommand, fullName: "teamsfx ..." },
        optionValues: { format: "table", description: true },
        globalOptionValues: {},
        argumentValues: ["key", "value"],
        telemetryProperties: {},
      };
      const res = await listTemplatesCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("table without description", async () => {
      const ctx: CLIContext = {
        command: { ...listTemplatesCommand, fullName: "teamsfx ..." },
        optionValues: { format: "table", description: false },
        globalOptionValues: {},
        argumentValues: ["key", "value"],
        telemetryProperties: {},
      };
      const res = await listTemplatesCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });

    it("json: bot Copilot plugin enabled only", async () => {
      mockedEnvRestore = mockedEnv({
        COPILOT_PLUGIN: "true",
        API_COPILOT_PLUGIN: "false",
      });
      const ctx: CLIContext = {
        command: { ...listTemplatesCommand, fullName: "teamsfx ..." },
        optionValues: { format: "json" },
        globalOptionValues: {},
        argumentValues: ["key", "value"],
        telemetryProperties: {},
      };
      const res = await listTemplatesCommand.handler!(ctx);
      assert.isTrue(res.isOk());
      assert.isFalse(!!messages.find((msg) => msg.includes("copilot-plugin-existing-api")));
    });

    it("json: API Copilot plugin feature flag enabled", async () => {
      mockedEnvRestore = mockedEnv({
        COPILOT_PLUGIN: "true",
        API_COPILOT_PLUGIN: "true",
      });
      const ctx: CLIContext = {
        command: { ...listTemplatesCommand, fullName: "teamsfx ..." },
        optionValues: { format: "json" },
        globalOptionValues: {},
        argumentValues: ["key", "value"],
        telemetryProperties: {},
      };
      const res = await listTemplatesCommand.handler!(ctx);
      assert.isTrue(res.isOk());
      assert.isTrue(!!messages.find((msg) => msg.includes("copilot-plugin-existing-api")));
    });
  });
  describe("listSamplesCommand", async () => {
    it("json", async () => {
      sandbox.stub(utils, "getTemplates").resolves([]);
      const ctx: CLIContext = {
        command: { ...listSamplesCommand, fullName: "teamsfx ..." },
        optionValues: { format: "json" },
        globalOptionValues: {},
        argumentValues: ["key", "value"],
        telemetryProperties: {},
      };
      const res = await listSamplesCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("table with filter + description", async () => {
      sandbox.stub(utils, "getTemplates").resolves([]);
      const ctx: CLIContext = {
        command: { ...listSamplesCommand, fullName: "teamsfx ..." },
        optionValues: { tag: "tab", format: "table", description: true },
        globalOptionValues: {},
        argumentValues: ["key", "value"],
        telemetryProperties: {},
      };
      const res = await listSamplesCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("table without description", async () => {
      sandbox.stub(utils, "getTemplates").resolves([]);
      const ctx: CLIContext = {
        command: { ...listSamplesCommand, fullName: "teamsfx ..." },
        optionValues: { format: "table", description: false },
        globalOptionValues: {},
        argumentValues: ["key", "value"],
        telemetryProperties: {},
      };
      const res = await listSamplesCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("helpCommand", async () => {
    it("happy", async () => {
      const ctx: CLIContext = {
        command: { ...helpCommand, fullName: "teamsfx ..." },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await helpCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });

  describe("doctor", async () => {
    describe("checkAccount", async () => {
      it("checkAccount error", async () => {
        sandbox
          .stub(DoctorChecker.prototype, "checkM365Account")
          .resolves(err(new UserCancelError()));
        const checker = new DoctorChecker();
        await checker.checkAccount();
      });
      it("checkAccount success", async () => {
        sandbox.stub(DoctorChecker.prototype, "checkM365Account").resolves(ok("success"));
        const checker = new DoctorChecker();
        await checker.checkAccount();
      });
    });
    describe("checkM365Account", async () => {
      it("checkM365Account - signin", async () => {
        const token = "test-token";
        const tenantId = "test-tenant-id";
        const upn = "test-user";
        sandbox.stub(M365TokenInstance, "getStatus").returns(
          Promise.resolve(
            ok({
              status: signedIn,
              token: token,
              accountInfo: {
                tid: tenantId,
                upn: upn,
              },
            })
          )
        );
        sandbox.stub(tools, "getSideloadingStatus").resolves(true);
        const checker = new DoctorChecker();
        const accountRes = await checker.checkM365Account();
        assert.isTrue(accountRes.isOk());
        const account = (accountRes as any).value;
        assert.include(account, "is logged in and sideloading permission enabled");
      });

      it("checkM365Account - signout", async () => {
        const token = "test-token";
        const tenantId = "test-tenant-id";
        const upn = "test-user";
        const getStatusStub = sandbox.stub(M365TokenInstance, "getStatus");
        getStatusStub.onCall(0).resolves(
          ok({
            status: signedOut,
          })
        );
        getStatusStub.onCall(1).resolves(
          ok({
            status: signedIn,
            token: token,
            accountInfo: {
              tid: tenantId,
              upn: upn,
            },
          })
        );
        sandbox.stub(M365TokenInstance, "getAccessToken").resolves(ok(token));
        sandbox.stub(tools, "getSideloadingStatus").resolves(true);
        const checker = new DoctorChecker();
        const accountRes = await checker.checkM365Account();
        assert.isTrue(accountRes.isOk());
        const account = (accountRes as any).value;
        assert.include(account, "is logged in and sideloading permission enabled");
      });

      it("checkM365Account - no sideloading permission", async () => {
        const token = "test-token";
        const tenantId = "test-tenant-id";
        const upn = "test-user";
        sandbox.stub(M365TokenInstance, "getStatus").returns(
          Promise.resolve(
            ok({
              status: signedIn,
              token: token,
              accountInfo: {
                tid: tenantId,
                upn: upn,
              },
            })
          )
        );
        sandbox.stub(tools, "getSideloadingStatus").resolves(false);
        const checker = new DoctorChecker();
        const accountRes = await checker.checkM365Account();
        assert.isTrue(accountRes.isOk());
        const value = (accountRes as any).value;
        assert.include(
          value,
          "Your Microsoft 365 tenant admin hasn't enabled sideloading permission for your account"
        );
      });
    });

    describe("checkNodejs", async () => {
      it("installed", async () => {
        sandbox
          .stub(LtsNodeChecker.prototype, "getInstallationInfo")
          .resolves({ isInstalled: true } as any);
        const checker = new DoctorChecker();
        await checker.checkNodejs();
      });
      it("error", async () => {
        sandbox
          .stub(LtsNodeChecker.prototype, "getInstallationInfo")
          .resolves({ isInstalled: true, error: new UserCancelError() } as any);
        const checker = new DoctorChecker();
        await checker.checkNodejs();
      });
      it("not installed", async () => {
        sandbox
          .stub(LtsNodeChecker.prototype, "getInstallationInfo")
          .resolves({ isInstalled: false } as any);
        const checker = new DoctorChecker();
        await checker.checkNodejs();
      });
    });
    describe("checkFuncCoreTool", async () => {
      it("installed", async () => {
        sandbox
          .stub(FuncToolChecker.prototype, "queryFuncVersion")
          .resolves({ versionStr: "3.0" } as any);
        const checker = new DoctorChecker();
        await checker.checkFuncCoreTool();
      });
      it("not installed", async () => {
        sandbox.stub(FuncToolChecker.prototype, "queryFuncVersion").rejects(new Error());
        const checker = new DoctorChecker();
        await checker.checkFuncCoreTool();
      });
    });
    describe("checkCert", async () => {
      it("not found", async () => {
        sandbox
          .stub(LocalCertificateManager.prototype, "setupCertificate")
          .resolves({ found: false } as any);
        const checker = new DoctorChecker();
        await checker.checkCert();
      });
      it("found trusted", async () => {
        sandbox
          .stub(LocalCertificateManager.prototype, "setupCertificate")
          .resolves({ found: true, alreadyTrusted: true } as any);
        const checker = new DoctorChecker();
        await checker.checkCert();
      });
      it("found not trusted", async () => {
        sandbox
          .stub(LocalCertificateManager.prototype, "setupCertificate")
          .resolves({ found: true, alreadyTrusted: false } as any);
        const checker = new DoctorChecker();
        await checker.checkCert();
      });
    });
    it("happy", async () => {
      sandbox.stub(DoctorChecker.prototype, "checkAccount").resolves();
      sandbox.stub(DoctorChecker.prototype, "checkNodejs").resolves();
      sandbox.stub(DoctorChecker.prototype, "checkFuncCoreTool").resolves();
      sandbox.stub(DoctorChecker.prototype, "checkCert").resolves();
      const ctx: CLIContext = {
        command: { ...teamsappDoctorCommand, fullName: "teamsapp doctor" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await teamsappDoctorCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
});
