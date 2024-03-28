// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import * as sinon from "sinon";
import chai from "chai";
import fs from "fs-extra";
import {
  ManifestUtil,
  SystemError,
  err,
  ok,
  Platform,
  TeamsAppManifest,
} from "@microsoft/teamsfx-api";
import * as commonTools from "../../../../src/common/tools";
import { ValidateManifestDriver } from "../../../../src/component/driver/teamsApp/validate";
import { ValidateManifestArgs } from "../../../../src/component/driver/teamsApp/interfaces/ValidateManifestArgs";
import { IAppValidationNote } from "../../../../src/component/driver/teamsApp/interfaces/appdefinitions/IValidationResult";
import { AsyncAppValidationResultsResponse } from "../../../../src/component/driver/teamsApp/interfaces/AsyncAppValidationResultsResponse";
import {
  AsyncAppValidationResponse,
  AsyncAppValidationStatus,
} from "../../../../src/component/driver/teamsApp/interfaces/AsyncAppValidationResponse";
import { ValidateAppPackageDriver } from "../../../../src/component/driver/teamsApp/validateAppPackage";
import { ValidateAppPackageArgs } from "../../../../src/component/driver/teamsApp/interfaces/ValidateAppPackageArgs";
import { ValidateWithTestCasesDriver } from "../../../../src/component/driver/teamsApp/validateTestCases";
import { ValidateWithTestCasesArgs } from "../../../../src/component/driver/teamsApp/interfaces/ValidateWithTestCasesArgs";
import { AppStudioError } from "../../../../src/component/driver/teamsApp/errors";
import { AppStudioClient } from "../../../../src/component/driver/teamsApp/clients/appStudioClient";
import {
  MockedLogProvider,
  MockedM365Provider,
  MockedUserInteraction,
} from "../../../plugins/solution/util";
import AdmZip from "adm-zip";
import { Constants } from "../../../../src/component/driver/teamsApp/constants";
import { metadataUtil } from "../../../../src/component/utils/metadataUtil";
import { InvalidActionInputError, UserCancelError } from "../../../../src/error/common";
import { teamsappMgr } from "../../../../src/component/driver/teamsApp/teamsappMgr";
import { setTools } from "../../../../src/core/globalVars";
import { MockTools } from "../../../core/utils";

describe("teamsApp/validateManifest", async () => {
  const teamsAppDriver = new ValidateManifestDriver();
  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    logProvider: new MockedLogProvider(),
    ui: new MockedUserInteraction(),
    projectPath: "./",
  };

  afterEach(() => {
    sinon.restore();
  });

  it("file not found - manifest", async () => {
    const args: ValidateManifestArgs = {
      manifestPath: "fakepath",
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(AppStudioError.FileNotFoundError.name, result.error.name);
    }
  });

  it("invalid param error", async () => {
    const args: ValidateManifestArgs = {
      manifestPath: "",
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof InvalidActionInputError);
    }
  });

  it("happy path - validate against schema", async () => {
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
    };

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("execute", async () => {
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
    };

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = await teamsAppDriver.execute(args, mockedDriverContext);
    chai.assert(result.result.isOk());
  });

  it("happy path - VS", async () => {
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
    };

    mockedDriverContext.platform = Platform.VS;

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("happy path- CLI", async () => {
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      showMessage: true,
    };

    mockedDriverContext.platform = Platform.CLI;

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("happy path- VSC", async () => {
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      showMessage: true,
    };

    mockedDriverContext.platform = Platform.VSCode;

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("validation error - no schema", async () => {
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.noSchema.manifest.json",
    };

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert(result.error.name, AppStudioError.ValidationFailedError.name);
    }
  });

  it("validation error - invalid", async () => {
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.invalid.manifest.json",
    };

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert(result.error.name, AppStudioError.ValidationFailedError.name);
    }
  });

  it("validation error - cli", async () => {
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.invalid.manifest.json",
    };

    const mockedCliDriverContext = {
      ...mockedDriverContext,
      platform: Platform.CLI,
    };

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = (await teamsAppDriver.execute(args, mockedCliDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert(result.error.name, AppStudioError.ValidationFailedError.name);
    }
  });

  it("validation error - download failed", async () => {
    sinon
      .stub(ManifestUtil, "validateManifest")
      .throws(new Error(`Failed to get manifest at url due to: unknown error`));
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
    };

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert(result.error.name, AppStudioError.ValidationFailedError.name);
    }
  });
});

describe("teamsApp/validateAppPackage", async () => {
  const teamsAppDriver = new ValidateAppPackageDriver();
  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    logProvider: new MockedLogProvider(),
    ui: new MockedUserInteraction(),
    projectPath: "./",
  };
  const contextWithoutUI: any = {
    m365TokenProvider: new MockedM365Provider(),
    logProvider: new MockedLogProvider(),
    projectPath: "./",
  };

  afterEach(() => {
    sinon.restore();
  });

  it("file not found - app package", async () => {
    const args: ValidateAppPackageArgs = {
      appPackagePath: "fakepath",
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(AppStudioError.FileNotFoundError.name, result.error.name);
    }
  });

  it("validate app package - error", async () => {
    sinon.stub(AppStudioClient, "partnerCenterAppPackageValidation").resolves({
      errors: [
        {
          id: "fakeId",
          content: "Reserved Tab Name property should not be specified.",
          filePath: "",
          helpUrl: "https://docs.microsoft.com",
          shortCodeNumber: 123,
          validationCategory: "tab",
          title: "tab name",
        },
        {
          id: "fakeId",
          content: "Reserved Tab Name property should not be specified.",
          filePath: "",
          shortCodeNumber: 123,
          validationCategory: "tab",
          title: "tab name",
        },
      ],
      status: "Rejected",
      warnings: [
        {
          id: "fakeId",
          content: "Valid domains cannot contain a hosting site with a wildcard.",
          filePath: "",
          shortCodeNumber: 123,
          validationCategory: "domain",
          title: "valid domain",
        },
        {
          id: "fakeId",
          content: "Valid domains cannot contain a hosting site with a wildcard.",
          filePath: "",
          helpUrl: "https://docs.microsoft.com",
          shortCodeNumber: 123,
          validationCategory: "domain",
          title: "valid domain",
        },
      ],
      notes: [
        {
          id: "fakeId",
          content: "Schema URL is present.",
          title: "schema",
        },
      ],
      addInDetails: {
        displayName: "fake name",
        developerName: "fake name",
        version: "1.14.1",
        manifestVersion: "1.14.1",
      },
    });
    sinon.stub(fs, "pathExists").resolves(true);
    // sinon.stub(fs, "readFile").resolves(Buffer.from(""));
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outlie.png", new Buffer(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    const args: ValidateAppPackageArgs = {
      appPackagePath: "fakePath",
      showMessage: true,
    };
    let result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());

    result = (await teamsAppDriver.execute(args, contextWithoutUI)).result;
    chai.assert(result.isErr());
  });

  it("validate app package - no error", async () => {
    sinon.stub(AppStudioClient, "partnerCenterAppPackageValidation").resolves({
      errors: [],
      status: "Accepted",
      warnings: [],
      notes: [
        {
          id: "fakeId",
          content: "Schema URL is present.",
          title: "schema",
        },
        {
          id: "632652a7-0cf8-43c7-a65d-6a19e5822467",
          title: "Manifest Version is valid",
          code: "The app is using manifest version '1.16'",
        } as any as IAppValidationNote,
      ],
      addInDetails: {
        displayName: "fake name",
        developerName: "fake name",
        version: "1.14.1",
        manifestVersion: "1.14.1",
      },
    });
    sinon.stub(fs, "pathExists").resolves(true);
    // sinon.stub(fs, "readFile").resolves(Buffer.from(""));
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outlie.png", new Buffer(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    const args: ValidateAppPackageArgs = {
      appPackagePath: "fakePath",
      showMessage: true,
    };
    let result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());

    result = (await teamsAppDriver.execute(args, contextWithoutUI)).result;
    chai.assert(result.isOk());
  });

  it("validate app package - stop-on-error", async () => {
    sinon.stub(AppStudioClient, "partnerCenterAppPackageValidation").resolves({
      errors: [
        {
          id: "fakeId",
          content: "Reserved Tab Name property should not be specified.",
          filePath: "",
          helpUrl: "https://docs.microsoft.com",
          shortCodeNumber: 123,
          validationCategory: "tab",
          title: "tab name",
        },
      ],
      status: "Rejected",
      warnings: [],
      notes: [],
      addInDetails: {
        displayName: "fake name",
        developerName: "fake name",
        version: "1.14.1",
        manifestVersion: "1.14.1",
      },
    });
    sinon.stub(fs, "pathExists").resolves(true);
    // sinon.stub(fs, "readFile").resolves(Buffer.from(""));
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outlie.png", new Buffer(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    const args: ValidateAppPackageArgs = {
      appPackagePath: "fakePath",
      showMessage: false,
    };
    let result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());

    result = (await teamsAppDriver.execute(args, contextWithoutUI)).result;
    chai.assert(result.isErr());
  });

  it("errors - cli", async () => {
    sinon.stub(AppStudioClient, "partnerCenterAppPackageValidation").resolves({
      errors: [
        {
          id: "fakeId",
          content: "Reserved Tab Name property should not be specified.",
          filePath: "",
          helpUrl: "https://docs.microsoft.com",
          shortCodeNumber: 123,
          validationCategory: "tab",
          title: "tab name",
        },
        {
          id: "fakeId",
          content: "Reserved Tab Name property should not be specified.",
          filePath: "",
          shortCodeNumber: 123,
          validationCategory: "tab",
          title: "tab name",
        },
      ],
      status: "Rejected",
      warnings: [
        {
          id: "fakeId",
          content: "Valid domains cannot contain a hosting site with a wildcard.",
          filePath: "",
          helpUrl: "https://docs.microsoft.com",
          shortCodeNumber: 123,
          validationCategory: "domain",
          title: "valid domain",
        },
        {
          id: "fakeId",
          content: "Valid domains cannot contain a hosting site with a wildcard.",
          filePath: "",
          shortCodeNumber: 123,
          validationCategory: "domain",
          title: "valid domain",
        },
      ],
      notes: [
        {
          id: "fakeId",
          content: "Schema URL is present.",
          title: "schema",
        },
      ],
      addInDetails: {
        displayName: "fake name",
        developerName: "fake name",
        version: "1.14.1",
        manifestVersion: "1.14.1",
      },
    });
    sinon.stub(fs, "pathExists").resolves(true);
    // sinon.stub(fs, "readFile").resolves(Buffer.from(""));
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outlie.png", new Buffer(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    const args: ValidateAppPackageArgs = {
      appPackagePath: "fakePath",
    };

    const mockedCliDriverContext = {
      ...mockedDriverContext,
      platform: Platform.CLI,
    };

    const result = (await teamsAppDriver.execute(args, mockedCliDriverContext)).result;
    chai.assert(result.isErr());
  });

  it("validation with only errors - cli", async () => {
    sinon.stub(AppStudioClient, "partnerCenterAppPackageValidation").resolves({
      errors: [
        {
          id: "fakeId",
          content: "Reserved Tab Name property should not be specified.",
          filePath: "",
          helpUrl: "https://docs.microsoft.com",
          shortCodeNumber: 123,
          validationCategory: "tab",
          title: "tab name",
        },
      ],
      status: "Rejected",
      warnings: [],
      notes: [
        {
          id: "fakeId",
          content: "Schema URL is present.",
          title: "schema",
        },
      ],
      addInDetails: {
        displayName: "fake name",
        developerName: "fake name",
        version: "1.14.1",
        manifestVersion: "1.14.1",
      },
    });
    sinon.stub(fs, "pathExists").resolves(true);
    // sinon.stub(fs, "readFile").resolves(Buffer.from(""));
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outlie.png", new Buffer(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    const args: ValidateAppPackageArgs = {
      appPackagePath: "fakePath",
    };

    const mockedCliDriverContext = {
      ...mockedDriverContext,
      platform: Platform.CLI,
    };

    const result = (await teamsAppDriver.execute(args, mockedCliDriverContext)).result;
    chai.assert(result.isErr());
  });

  it("validation with warnings - cli", async () => {
    sinon.stub(AppStudioClient, "partnerCenterAppPackageValidation").resolves({
      errors: [],
      status: "Rejected",
      warnings: [
        {
          id: "fakeId",
          content: "Valid domains cannot contain a hosting site with a wildcard.",
          filePath: "",
          helpUrl: "https://docs.microsoft.com",
          shortCodeNumber: 123,
          validationCategory: "domain",
          title: "valid domain",
        },
        {
          id: "fakeId",
          content: "Valid domains cannot contain a hosting site with a wildcard.",
          filePath: "",
          shortCodeNumber: 123,
          validationCategory: "domain",
          title: "valid domain",
        },
      ],
      notes: [
        {
          id: "fakeId",
          content: "Schema URL is present.",
          title: "schema",
        },
      ],
      addInDetails: {
        displayName: "fake name",
        developerName: "fake name",
        version: "1.14.1",
        manifestVersion: "1.14.1",
      },
    });
    sinon.stub(fs, "pathExists").resolves(true);
    // sinon.stub(fs, "readFile").resolves(Buffer.from(""));
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outlie.png", new Buffer(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    const args: ValidateAppPackageArgs = {
      appPackagePath: "fakePath",
    };

    const mockedCliDriverContext = {
      ...mockedDriverContext,
      platform: Platform.CLI,
    };

    const result = (await teamsAppDriver.execute(args, mockedCliDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("happy path - cli", async () => {
    sinon.stub(AppStudioClient, "partnerCenterAppPackageValidation").resolves({
      errors: [],
      status: "Rejected",
      warnings: [],
      notes: [
        {
          id: "fakeId",
          content: "Schema URL is present.",
          title: "schema",
        },
        {
          id: "632652a7-0cf8-43c7-a65d-6a19e5822467",
          title: "Manifest Version is valid",
          code: "The app is using manifest version '1.16'",
        } as any as IAppValidationNote,
      ],
      addInDetails: {
        displayName: "fake name",
        developerName: "fake name",
        version: "1.14.1",
        manifestVersion: "1.14.1",
      },
    });
    sinon.stub(fs, "pathExists").resolves(true);
    // sinon.stub(fs, "readFile").resolves(Buffer.from(""));
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outlie.png", new Buffer(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    const args: ValidateAppPackageArgs = {
      appPackagePath: "fakePath",
    };

    const mockedCliDriverContext = {
      ...mockedDriverContext,
      platform: Platform.CLI,
    };

    const result = (await teamsAppDriver.execute(args, mockedCliDriverContext)).result;
    chai.assert(result.isOk());
  });
});

describe("teamsApp/validateWithTestCases", async () => {
  const tools = new MockTools();
  setTools(tools);

  const teamsAppDriver = new ValidateWithTestCasesDriver();

  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    logProvider: new MockedLogProvider(),
    ui: new MockedUserInteraction(),
    projectPath: "./",
  };

  beforeEach(() => {
    sinon.stub(commonTools, "waitSeconds").resolves();
  });

  afterEach(() => {
    sinon.restore();
  });

  it("file not found - app package", async () => {
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(AppStudioError.FileNotFoundError.name, result.error.name);
    }
  });

  it("file not found - manifest.json", async () => {
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(AppStudioError.FileNotFoundError.name, result.error.name);
    }
  });

  it("invalid param error", async () => {
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "",
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof InvalidActionInputError);
    }
  });

  it("Failed to get token", async () => {
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakePath",
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");
    sinon
      .stub(mockedDriverContext.m365TokenProvider, "getAccessToken")
      .resolves(err(new SystemError({})));

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
  });

  it("Invalid Validation Result Response - Null Details", async () => {
    const mockSubmitValidationResponse: AsyncAppValidationResponse = {
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    };
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const invalidValidationResultResponseJson: any = {
      appValidationId: "appValidationId123",
      appId: "appId123",
      status: "Completed",
      appVersion: "1.0.0",
      manifestVersion: "1.0.0",
      createdAt: "2024-03-27T12:00:00.000Z",
      updatedAt: "2024-03-27T12:00:00.000Z",
      validationResults: {
        successes: null,
        warnings: null,
        failures: null,
        skipped: null,
      },
    };
    const invalidValidationResultResponse: AsyncAppValidationResultsResponse = <
      AsyncAppValidationResultsResponse
    >invalidValidationResultResponseJson;
    sinon.stub(AppStudioClient, "getAppValidationById").resolves(invalidValidationResultResponse);
    await teamsAppDriver.runningBackgroundJob(
      args,
      mockedDriverContext,
      "test_token",
      mockSubmitValidationResponse,
      "test_id"
    );
    chai.assert(
      mockedDriverContext.logProvider.msg.includes("Validation request completed, status:")
    );
  });

  it("Invalid Validation Result Response - Null Validation Results", async () => {
    const mockSubmitValidationResponse: AsyncAppValidationResponse = {
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    };
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const invalidValidationResultResponseJson: any = {
      appValidationId: "appValidationId123",
      appId: "appId123",
      status: "Completed",
      appVersion: "1.0.0",
      manifestVersion: "1.0.0",
      createdAt: "2024-03-27T12:00:00.000Z",
      updatedAt: "2024-03-27T12:00:00.000Z",
      validationResults: null,
    };
    const invalidValidationResultResponse: AsyncAppValidationResultsResponse = <
      AsyncAppValidationResultsResponse
    >invalidValidationResultResponseJson;
    sinon.stub(AppStudioClient, "getAppValidationById").resolves(invalidValidationResultResponse);
    await teamsAppDriver.runningBackgroundJob(
      args,
      mockedDriverContext,
      "test_token",
      mockSubmitValidationResponse,
      "test_id"
    );
    chai.assert(
      mockedDriverContext.logProvider.msg.includes("Validation request completed, status:")
    );
  });

  it("Valid Validation Result Response", async () => {
    const mockSubmitValidationResponse: AsyncAppValidationResponse = {
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    };
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const invalidValidationResultResponseJson: any = {
      status: AsyncAppValidationStatus.Completed,
      appValidationId: "fakeId",
      appId: "fakeAppId",
      appVersion: "1.0.0",
      manifestVersion: "1.16",
      validationResults: {
        successes: [
          {
            title: "Validation_Success_Example",
            message: "Success validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        warnings: [
          {
            title: "Validation_Warning_Example",
            message: "Warning validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        failures: [
          {
            title: "Validation_Failure_Example",
            message: "Failure validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        skipped: [
          {
            title: "Validation_Skipped_Example",
            message: "Skipped validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
      },
      createdAt: new Date(),
      updatedAt: new Date(),
    };
    const invalidValidationResultResponse: AsyncAppValidationResultsResponse = <
      AsyncAppValidationResultsResponse
    >invalidValidationResultResponseJson;
    sinon.stub(AppStudioClient, "getAppValidationById").resolves(invalidValidationResultResponse);
    await teamsAppDriver.runningBackgroundJob(
      args,
      mockedDriverContext,
      "test_token",
      mockSubmitValidationResponse,
      "test_id"
    );
    chai.assert(
      mockedDriverContext.logProvider.msg.includes("Validation request completed, status:")
    );
    chai.assert(
      mockedDriverContext.logProvider.msg.includes("1 failed, 1 warning, 1 skipped, 1 passed")
    );
  });

  it("Happy path", async () => {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    sinon.stub(AppStudioClient, "submitAppValidationRequest").resolves({
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    });

    sinon.stub(AppStudioClient, "getAppValidationById").resolves({
      status: AsyncAppValidationStatus.Completed,
      appValidationId: "fakeId",
      appId: "fakeAppId",
      appVersion: "1.0.0",
      manifestVersion: "1.16",
      validationResults: {
        successes: [
          {
            title: "Validation_Success_Example",
            message: "Success validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        warnings: [
          {
            title: "Validation_Warning_Example",
            message: "Warning validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        failures: [
          {
            title: "Validation_Failure_Example",
            message: "Failure validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        skipped: [
          {
            title: "Validation_Skipped_Example",
            message: "Skipped validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
      },
      createdAt: new Date(),
      updatedAt: new Date(),
    });

    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("Aborted", async () => {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    sinon.stub(AppStudioClient, "submitAppValidationRequest").resolves({
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    });

    sinon.stub(AppStudioClient, "getAppValidationById").resolves({
      status: AsyncAppValidationStatus.Aborted,
      appValidationId: "fakeId",
      appId: "fakeAppId",
      appVersion: "1.0.0",
      manifestVersion: "1.16",
      validationResults: {
        failures: [],
        warnings: [],
        successes: [],
        skipped: [],
      },
      createdAt: new Date(),
      updatedAt: new Date(),
    });

    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: false,
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("Happy path - CLI", async () => {
    const mockedCliDriverContext = {
      ...mockedDriverContext,
      platform: Platform.CLI,
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    sinon.stub(AppStudioClient, "submitAppValidationRequest").resolves({
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    });

    sinon.stub(AppStudioClient, "getAppValidationById").resolves({
      status: AsyncAppValidationStatus.Completed,
      appValidationId: "fakeId",
      appId: "fakeAppId",
      appVersion: "1.0.0",
      manifestVersion: "1.16",
      validationResults: {
        failures: [],
        warnings: [],
        successes: [],
        skipped: [],
      },
      createdAt: new Date(),
      updatedAt: new Date(),
    });

    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const result = (await teamsAppDriver.execute(args, mockedCliDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("CLI - succeed", async () => {
    sinon.stub(ValidateWithTestCasesDriver.prototype, "validate").resolves(ok(new Map()));
    const result = await teamsappMgr.validateTeamsApp({
      projectPath: "xxx",
      platform: Platform.CLI,
      "package-file": "xxx",
      "validate-method": "test-cases",
    });
    chai.assert(result.isOk());
  });

  it("CLI - failed", async () => {
    sinon
      .stub(ValidateWithTestCasesDriver.prototype, "validate")
      .resolves(err(new UserCancelError()));
    const result = await teamsappMgr.validateTeamsApp({
      projectPath: "xxx",
      platform: Platform.CLI,
      "package-file": "xxx",
      "validate-method": "test-cases",
    });
    chai.assert(result.isErr());
  });
});
