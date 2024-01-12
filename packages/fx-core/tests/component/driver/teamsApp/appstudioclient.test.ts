// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import * as chai from "chai";
import * as sinon from "sinon";
import axios from "axios";
import { v4 as uuid } from "uuid";
import { TeamsAppManifest, ok, err } from "@microsoft/teamsfx-api";
import { AppStudioClient } from "../../../../src/component/driver/teamsApp/clients/appStudioClient";
import { AppDefinition } from "../../../../src/component/driver/teamsApp/interfaces/appdefinitions/appDefinition";
import { AppUser } from "../../../../src/component/driver/teamsApp/interfaces/appdefinitions/appUser";
import { AppStudioError } from "../../../../src/component/driver/teamsApp/errors";
import { TelemetryUtils } from "../../../../src/component/driver/teamsApp/utils/telemetry";
import { RetryHandler } from "../../../../src/component/driver/teamsApp/utils/utils";
import { PublishingState } from "../../../../src/component/driver/teamsApp/interfaces/appdefinitions/IPublishingAppDefinition";
import { manifestUtils } from "../../../../src/component/driver/teamsApp/utils/ManifestUtils";
import { AppStudioResultFactory } from "../../../../src/component/driver/teamsApp/results";
import { Constants } from "../../../../src/component/driver/teamsApp/constants";
import { MockedLogProvider } from "../../../plugins/solution/util";
import { DeveloperPortalAPIFailedError } from "../../../../src/error/teamsApp";
import {
  ApiSecretRegistration,
  ApiSecretRegistrationAppType,
} from "../../../../src/component/driver/teamsApp/interfaces/ApiSecretRegistration";

function newEnvInfo() {
  return {
    envName: "default",
    config: {},
    state: new Map(),
  };
}

describe("App Studio API Test", () => {
  const appStudioToken = "appStudioToken";
  const logProvider = new MockedLogProvider();

  const appDef: AppDefinition = {
    appName: "fake",
    teamsAppId: uuid(),
    userList: [],
  };

  const appApiRegistration: ApiSecretRegistration = {
    id: "fakeId",
    description: "An Api Key registration for auth",
    clientSecrets: [
      {
        id: uuid(),
        value: "fakeValue",
        isValueRedacted: false,
      },
    ],
    applicableToApps: ApiSecretRegistrationAppType.AnyApp,
    targetUrlsShouldStartWith: ["https://www.example.com"],
  };

  beforeEach(() => {
    sinon.stub(RetryHandler, "RETRIES").value(1);
  });

  afterEach(() => {
    sinon.restore();
  });

  describe("publish Teams app", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);
      const response = {
        data: {
          id: "fakeId",
        },
      };
      sinon.stub(fakeAxiosInstance, "post").resolves(response);

      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      const res = await AppStudioClient.publishTeamsApp(
        appStudioToken,
        Buffer.from(""),
        appStudioToken
      );
      chai.assert.equal(res, response.data.id);
    });

    it("API Failure", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "error",
        message: "fake message",
      };
      sinon.stub(fakeAxiosInstance, "post").throws(error);

      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      sinon.stub(TelemetryUtils, "sendErrorEvent").callsFake(() => {});

      try {
        await AppStudioClient.publishTeamsApp(appStudioToken, Buffer.from(""), appStudioToken);
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedError.name);
      }
    });

    it("should contain x-correlation-id on BadeRequest with 2xx status code", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const xCorrelationId = "fakeCorrelationId";
      const response = {
        data: {
          error: "BadRequest",
        },
        message: "fake message",
        headers: {
          "x-correlation-id": xCorrelationId,
        },
      };
      sinon.stub(fakeAxiosInstance, "post").resolves(response);

      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      sinon.stub(TelemetryUtils, "sendErrorEvent").callsFake(() => {});

      try {
        await AppStudioClient.publishTeamsApp(appStudioToken, Buffer.from(""), appStudioToken);
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedError.name);
        chai.assert.include(error.message, xCorrelationId);
      }
    });

    it("Bad gateway", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const postResponse = {
        data: {
          error: {
            code: "BadGateway",
            message: "fakeMessage",
          },
        },
      };
      sinon.stub(fakeAxiosInstance, "post").resolves(postResponse);

      const getResponse = {
        data: {
          value: [
            {
              appDefinitions: [
                {
                  lastModifiedDateTime: new Date(),
                  publishingState: PublishingState.submitted,
                  teamsAppId: uuid(),
                  displayName: "fakeApp",
                },
              ],
            },
          ],
        },
      };
      sinon.stub(fakeAxiosInstance, "get").resolves(getResponse);

      const res = await AppStudioClient.publishTeamsApp(
        appStudioToken,
        Buffer.from(""),
        appStudioToken
      );
      chai.assert.equal(res, getResponse.data.value[0].appDefinitions[0].teamsAppId);
    });

    it("AppdefinitionsAlreadyExists - update", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const publishResponse = {
        data: {
          error: {
            code: "Conflict",
            message: "Conflict",
            innerError: {
              code: "AppDefinitionAlreadyExists",
            },
          },
        },
      };

      const updateResponse = {
        data: {
          teamsAppId: "fakeId",
        },
      };
      sinon
        .stub(fakeAxiosInstance, "post")
        .onFirstCall()
        .resolves(publishResponse)
        .onSecondCall()
        .resolves(updateResponse);
      sinon.stub(AppStudioClient, "publishTeamsAppUpdate").resolves("fakeId");

      const getResponse = {
        data: {
          value: [
            {
              appDefinitions: [
                {
                  lastModifiedDateTime: new Date(),
                  publishingState: PublishingState.submitted,
                  teamsAppId: uuid(),
                  displayName: "fakeApp",
                },
              ],
            },
          ],
        },
      };
      sinon.stub(fakeAxiosInstance, "get").resolves(getResponse);

      const res = await AppStudioClient.publishTeamsApp(
        appStudioToken,
        Buffer.from(""),
        appStudioToken
      );
      chai.assert.equal(res, "fakeId");
    });

    it("AppdefinitionsAlreadyExists - failed", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const postResponse = {
        data: {
          error: {
            code: "Conflict",
            message: "Conflict",
            innerError: {
              code: "AppDefinitionAlreadyExists",
            },
          },
        },
      };
      sinon.stub(fakeAxiosInstance, "post").resolves(postResponse);

      try {
        await AppStudioClient.publishTeamsApp(appStudioToken, Buffer.from(""), appStudioToken);
      } catch (error) {
        chai.assert.equal(error.name, AppStudioError.TeamsAppPublishConflictError.name);
      }
    });
  });

  describe("import Teams app", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: appDef,
      };
      sinon.stub(fakeAxiosInstance, "post").resolves(response);

      AppStudioClient.setRegion("https://dev.teams.microsoft.com/amer");

      const res = await AppStudioClient.importApp(Buffer.from(""), appStudioToken, logProvider);
      chai.assert.equal(res, appDef);
    });

    it("Happy path - with wrong region", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: appDef,
      };
      sinon.stub(fakeAxiosInstance, "post").resolves(response);
      AppStudioClient.setRegion("https://dev.teams.microsoft.com");

      const res = await AppStudioClient.importApp(Buffer.from(""), appStudioToken, logProvider);
      chai.assert.equal(res, appDef);
    });

    it("409 conflict", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        response: {
          status: 409,
        },
      };
      sinon.stub(fakeAxiosInstance, "post").throws(error);

      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      sinon.stub(TelemetryUtils, "sendErrorEvent").callsFake(() => {});

      try {
        await AppStudioClient.importApp(Buffer.from(""), appStudioToken, logProvider);
      } catch (error) {
        chai.assert.equal(error.name, AppStudioError.TeamsAppCreateConflictError.name);
      }
    });

    it("422 conflict", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        response: {
          status: 422,
          data: "Unable import, App already exists and published. publishStatus: 'LobStore'",
        },
      };
      sinon.stub(fakeAxiosInstance, "post").throws(error);
      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      sinon.stub(TelemetryUtils, "sendErrorEvent").callsFake(() => {});

      try {
        await AppStudioClient.importApp(Buffer.from(""), appStudioToken, logProvider);
      } catch (error) {
        chai.assert.equal(
          error.name,
          AppStudioError.TeamsAppCreateConflictWithPublishedAppError.name
        );
      }
    });

    it("422 other error", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        response: {
          status: 422,
          data: "fake error message",
          headers: {
            "x-correlation-id": uuid(),
          },
        },
      };
      sinon.stub(fakeAxiosInstance, "post").throws(error);
      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      sinon.stub(TelemetryUtils, "sendErrorEvent").callsFake(() => {});

      try {
        await AppStudioClient.importApp(Buffer.from(""), appStudioToken, logProvider);
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedError.name);
      }
    });

    it("invalid Teams app id", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);
      sinon
        .stub(manifestUtils, "extractManifestFromArchivedFile")
        .returns(ok(new TeamsAppManifest()));

      const error = {
        response: {
          status: 400,
          data: "App Id must be a GUID",
        },
      };
      sinon.stub(fakeAxiosInstance, "post").throws(error);

      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      sinon.stub(TelemetryUtils, "sendErrorEvent").callsFake(() => {});

      try {
        await AppStudioClient.importApp(Buffer.from(""), appStudioToken, logProvider);
      } catch (error) {
        chai.assert.equal(error.name, AppStudioError.InvalidTeamsAppIdError.name);
      }
    });

    it("extract manifet failed", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);
      const fileNotFoundError = AppStudioResultFactory.UserError(
        AppStudioError.FileNotFoundError.name,
        AppStudioError.FileNotFoundError.message(Constants.MANIFEST_FILE)
      );
      sinon.stub(manifestUtils, "extractManifestFromArchivedFile").returns(err(fileNotFoundError));

      const error = {
        response: {
          status: 400,
          data: "App Id must be a GUID",
        },
      };
      sinon.stub(fakeAxiosInstance, "post").throws(error);

      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      sinon.stub(TelemetryUtils, "sendErrorEvent").callsFake(() => {});

      try {
        await AppStudioClient.importApp(Buffer.from(""), appStudioToken, logProvider);
      } catch (error) {
        chai.assert.equal(error.name, AppStudioError.FileNotFoundError.name);
      }
    });

    it("400 bad reqeust", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        response: {
          staus: 400,
          data: "BadRequest",
          headers: {
            "x-correlation-id": uuid(),
          },
        },
        message: "fake message",
      };
      sinon.stub(fakeAxiosInstance, "post").throws(error);
      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      sinon.stub(TelemetryUtils, "sendErrorEvent").callsFake(() => {});

      try {
        await AppStudioClient.importApp(Buffer.from(""), appStudioToken, logProvider);
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedError.name);
      }
    });
  });

  describe("get Teams app", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: appDef,
      };
      sinon.stub(fakeAxiosInstance, "get").resolves(response);

      const res = await AppStudioClient.getApp(appDef.teamsAppId!, appStudioToken, logProvider);
      chai.assert.equal(res, appDef);
    });

    it("404 not found", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "404",
        message: "fake message",
      };
      sinon.stub(fakeAxiosInstance, "get").throws(error);

      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      sinon.stub(TelemetryUtils, "sendErrorEvent").callsFake(() => {});

      try {
        await AppStudioClient.getApp(appDef.teamsAppId!, appStudioToken, logProvider);
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedError.name);
      }
    });

    it("region - 404", async () => {
      AppStudioClient.setRegion("https://dev.teams.microsoft.com/amer");
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        response: {
          status: 404,
          headers: {
            "x-correlation-id": "fakeCorrelationId",
          },
        },
      };
      sinon.stub(fakeAxiosInstance, "get").throws(error);

      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      sinon.stub(TelemetryUtils, "sendErrorEvent").callsFake(() => {});

      try {
        await AppStudioClient.getApp(appDef.teamsAppId!, appStudioToken, logProvider);
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedError.name);
      } finally {
        AppStudioClient.setRegion(undefined as unknown as string);
      }
    });
  });

  describe("get app package", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: "fakeData",
      };
      sinon.stub(fakeAxiosInstance, "get").resolves(response);

      const res = await AppStudioClient.getAppPackage(
        appDef.teamsAppId!,
        appStudioToken,
        logProvider
      );
      chai.assert.equal(res, "fakeData");
    });

    it("404 not found", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "404",
        message: "fake message",
      };
      sinon.stub(fakeAxiosInstance, "get").throws(error);

      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      sinon.stub(TelemetryUtils, "sendErrorEvent").callsFake(() => {});

      try {
        await AppStudioClient.getAppPackage(appDef.teamsAppId!, appStudioToken, logProvider);
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedError.name);
      }
    });
  });

  describe("partner center app validation", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: {
          status: "Accepted",
          errors: [],
          warnings: [],
          notes: [],
          addInDetails: {
            displayName: "fakeApp",
            developerName: "Teams",
            version: "0.0.1",
            manifestVersion: "1.16",
          },
        },
      };
      sinon.stub(fakeAxiosInstance, "post").resolves(response);

      const res = await AppStudioClient.partnerCenterAppPackageValidation(
        Buffer.from(""),
        appStudioToken
      );
      chai.assert.equal(res, response.data);
    });

    it("422", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "422",
        message: "Invalid zip",
      };
      sinon.stub(fakeAxiosInstance, "post").throws(error);

      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      sinon.stub(TelemetryUtils, "sendErrorEvent").callsFake(() => {});

      try {
        await AppStudioClient.partnerCenterAppPackageValidation(Buffer.from(""), appStudioToken);
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedError.name);
      }
    });
  });

  describe("Check exists in tenant", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: true,
      };
      sinon.stub(fakeAxiosInstance, "get").resolves(response);

      const res = await AppStudioClient.checkExistsInTenant(appDef.teamsAppId!, appStudioToken);
      chai.assert.isTrue(res);
    });

    it("404 not found", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "404",
        message: "fake message",
      };
      sinon.stub(fakeAxiosInstance, "get").throws(error);

      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      sinon.stub(TelemetryUtils, "sendErrorEvent").callsFake(() => {});

      try {
        await AppStudioClient.checkExistsInTenant(appDef.teamsAppId!, appStudioToken);
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedError.name);
      }
    });
  });

  describe("publishTeamsAppUpdate", () => {
    it("should contain x-correlation-id on BadeRequest with 2xx status code", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const xCorrelationId = "fakeCorrelationId";
      const postResponse = {
        data: {
          error: "BadRequest",
        },
        message: "fake message",
        headers: {
          "x-correlation-id": xCorrelationId,
        },
      };

      sinon.stub(fakeAxiosInstance, "post").resolves(postResponse);

      const getResponse = {
        data: {
          value: [
            {
              appDefinitions: [
                {
                  publishingState: PublishingState.submitted,
                  teamsAppId: "xx",
                  displayName: "xx",
                  lastModifiedDateTime: null,
                },
              ],
            },
          ],
        },
      };
      sinon.stub(fakeAxiosInstance, "get").resolves(getResponse);

      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      sinon.stub(TelemetryUtils, "sendErrorEvent").callsFake(() => {});

      try {
        await AppStudioClient.publishTeamsAppUpdate("", Buffer.from(""), appStudioToken);
      } catch (error) {
        chai.assert.include(error.message, xCorrelationId);
      }
    });

    it("API Failure", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "error",
        message: "fake message",
      };
      sinon.stub(fakeAxiosInstance, "post").throws(error);

      const getResponse = {
        data: {
          value: [
            {
              appDefinitions: [
                {
                  publishingState: PublishingState.submitted,
                  teamsAppId: "xx",
                  displayName: "xx",
                  lastModifiedDateTime: null,
                },
              ],
            },
          ],
        },
      };
      sinon.stub(fakeAxiosInstance, "get").resolves(getResponse);

      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      sinon.stub(TelemetryUtils, "sendErrorEvent").callsFake(() => {});

      try {
        await AppStudioClient.publishTeamsAppUpdate(
          appStudioToken,
          Buffer.from(""),
          appStudioToken
        );
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedError.name);
      }
    });
  });

  describe("grantPermission", () => {
    it("API Failure", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "error",
        message: "fake message",
      };
      sinon.stub(fakeAxiosInstance, "post").throws(error);
      sinon.stub(fakeAxiosInstance, "get").resolves({ data: appDef });

      const appUser: AppUser = {
        tenantId: uuid(),
        aadId: uuid(),
        displayName: "fake",
        userPrincipalName: "fake",
        isAdministrator: false,
      };

      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      sinon.stub(TelemetryUtils, "sendErrorEvent").callsFake(() => {});

      try {
        await AppStudioClient.grantPermission(
          appDef.teamsAppId!,
          appStudioToken,
          appUser,
          logProvider
        );
      } catch (e) {
        chai.assert.equal(e.name, error.name);
      }
    });

    it("happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const newAppUser: AppUser = {
        tenantId: "new-tenant-id",
        aadId: "new-aad-id",
        displayName: "fake",
        userPrincipalName: "fake",
        isAdministrator: false,
      };
      const teamsAppId = appDef.teamsAppId!;
      const appDefWithUser: AppDefinition = {
        appName: "fake",
        teamsAppId: teamsAppId,
        userList: [
          {
            tenantId: "fake-tenant-id",
            aadId: "fake-aad-id",
            displayName: "fake",
            userPrincipalName: "fake",
            isAdministrator: false,
          },
        ],
      };
      const appDefWithUserAdded: AppDefinition = {
        appName: "fake",
        teamsAppId: teamsAppId,
        userList: [
          {
            tenantId: "fake-tenant-id",
            aadId: "fake-aad-id",
            displayName: "fake",
            userPrincipalName: "fake",
            isAdministrator: false,
          },
          newAppUser,
        ],
      };
      sinon.stub(fakeAxiosInstance, "get").resolves({
        data: appDefWithUser,
      });
      sinon.stub(fakeAxiosInstance, "post").resolves({
        data: appDefWithUserAdded,
      });

      const res = await AppStudioClient.grantPermission(
        appDef.teamsAppId!,
        appStudioToken,
        newAppUser,
        logProvider
      );
    });

    it("happy path with region", async () => {
      AppStudioClient.setRegion("https://dev.teams.microsoft.com/amer");

      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const newAppUser: AppUser = {
        tenantId: "new-tenant-id",
        aadId: "new-aad-id",
        displayName: "fake",
        userPrincipalName: "fake",
        isAdministrator: false,
      };
      const teamsAppId = appDef.teamsAppId!;
      const appDefWithUser: AppDefinition = {
        appName: "fake",
        teamsAppId: teamsAppId,
        userList: [
          {
            tenantId: "fake-tenant-id",
            aadId: "fake-aad-id",
            displayName: "fake",
            userPrincipalName: "fake",
            isAdministrator: false,
          },
        ],
      };
      const appDefWithUserAdded: AppDefinition = {
        appName: "fake",
        teamsAppId: teamsAppId,
        userList: [
          {
            tenantId: "fake-tenant-id",
            aadId: "fake-aad-id",
            displayName: "fake",
            userPrincipalName: "fake",
            isAdministrator: false,
          },
          newAppUser,
        ],
      };
      sinon.stub(fakeAxiosInstance, "get").resolves({
        data: appDefWithUser,
      });
      sinon.stub(fakeAxiosInstance, "post").resolves({
        data: appDefWithUserAdded,
      });

      const res = await AppStudioClient.grantPermission(
        appDef.teamsAppId!,
        appStudioToken,
        newAppUser,
        logProvider
      );
    });
  });

  describe("getUserList", () => {
    it("happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: appDef,
      };
      sinon.stub(fakeAxiosInstance, "get").resolves(response);

      const res = await AppStudioClient.getUserList(
        appDef.teamsAppId!,
        appStudioToken,
        logProvider
      );
    });
  });

  describe("checkPermission", () => {
    it("No permission", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: appDef,
      };
      sinon.stub(fakeAxiosInstance, "get").resolves(response);

      const res = await AppStudioClient.checkPermission(
        appDef.teamsAppId!,
        appStudioToken,
        "fakeUesrId",
        logProvider
      );
      chai.assert.equal(res, Constants.PERMISSIONS.noPermission);
    });
  });

  describe("getApiKeyRegistration", () => {
    it("404 not found", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "404",
        message: "fake message",
      };
      sinon.stub(fakeAxiosInstance, "get").throws(error);

      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      sinon.stub(TelemetryUtils, "sendErrorEvent").callsFake(() => {});

      try {
        await AppStudioClient.getApiKeyRegistrationById(appStudioToken, "fakeId");
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedError.name);
      }
    });

    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: appApiRegistration,
      };
      sinon.stub(fakeAxiosInstance, "get").resolves(response);

      const res = await AppStudioClient.getApiKeyRegistrationById(appStudioToken, "fakeId");
      chai.assert.equal(res, appApiRegistration);
    });
  });

  describe("createApiKeyRegistration", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: appApiRegistration,
      };
      sinon.stub(fakeAxiosInstance, "post").resolves(response);

      const res = await AppStudioClient.createApiKeyRegistration(
        appStudioToken,
        appApiRegistration
      );
      chai.assert.equal(res, appApiRegistration);
    });

    it("Graph API failure", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        response: {
          staus: 400,
          data: {
            statusCode: 400,
            errorMessage:
              "Unsuccessful response received from Teams Graph Service. Error Message: System.Net.Http.HttpConnectionResponseContent",
          },
          headers: {
            "x-correlation-id": uuid(),
          },
        },
      };
      sinon.stub(fakeAxiosInstance, "post").throws(error);

      const ctx = {
        envInfo: newEnvInfo(),
        root: "fakeRoot",
      } as any;
      TelemetryUtils.init(ctx);
      sinon.stub(TelemetryUtils, "sendErrorEvent").callsFake(() => {});

      try {
        await AppStudioClient.createApiKeyRegistration(appStudioToken, appApiRegistration);
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedError.name);
      }
    });
  });

  describe("list Teams app", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: [appDef],
      };
      sinon.stub(fakeAxiosInstance, "get").resolves(response);
      AppStudioClient.setRegion("https://dev.teams.microsoft.com/amer");
      const res = await AppStudioClient.listApps(appStudioToken, logProvider);
      chai.assert.deepEqual(res, [appDef]);
    });
    it("Error - no region", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: [appDef],
      };
      sinon.stub(fakeAxiosInstance, "get").resolves(response);
      AppStudioClient.setRegion("");
      try {
        await AppStudioClient.listApps(appStudioToken, logProvider);
        chai.assert.fail("should throw error");
      } catch (e) {
        chai.assert.isTrue(e instanceof DeveloperPortalAPIFailedError);
      }
    });

    it("Error - no data", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: undefined,
      };
      sinon.stub(fakeAxiosInstance, "get").resolves(response);
      AppStudioClient.setRegion("https://dev.teams.microsoft.com/amer");
      try {
        await AppStudioClient.listApps(appStudioToken, logProvider);
        chai.assert.fail("should throw error");
      } catch (e) {
        chai.assert.equal(e.message, "Cannot get the app definitions");
      }
    });
  });

  describe("delete Teams app", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);
      const response = {
        data: true,
      };
      sinon.stub(fakeAxiosInstance, "delete").resolves(response);
      AppStudioClient.setRegion("https://dev.teams.microsoft.com/amer");
      const res = await AppStudioClient.deleteApp("testid", appStudioToken, logProvider);
      chai.assert.isTrue(res);
    });
    it("Error - no region", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: [appDef],
      };
      sinon.stub(fakeAxiosInstance, "delete").resolves(response);
      AppStudioClient.setRegion("");
      try {
        await AppStudioClient.deleteApp("testid", appStudioToken, logProvider);
        chai.assert.fail("should throw error");
      } catch (e) {
        chai.assert.isTrue(e instanceof DeveloperPortalAPIFailedError);
      }
    });

    it("Error - no data", async () => {
      const fakeAxiosInstance = axios.create();
      sinon.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: undefined,
      };
      sinon.stub(fakeAxiosInstance, "delete").resolves(response);
      AppStudioClient.setRegion("https://dev.teams.microsoft.com/amer");
      try {
        await AppStudioClient.deleteApp("testid", appStudioToken, logProvider);
        chai.assert.fail("should throw error");
      } catch (e) {
        chai.assert.equal(e.message, "Cannot delete the app: " + "testid");
      }
    });
  });
});
