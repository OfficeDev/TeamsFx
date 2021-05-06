// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { assert, expect, use as chaiUse } from "chai";
import chaiPromises from "chai-as-promised";
import { AccessToken, GetTokenOptions } from "@azure/core-auth";
import sinon from "sinon";
import { loadConfiguration, TeamsUserCredential, ErrorWithCode, ErrorCode } from "../../../src";
import { getSSOToken, AADJwtPayLoad } from "../helper.browser";
import jwtDecode from "jwt-decode";

chaiUse(chaiPromises);
const env = (window as any).__env__;

const FAKE_LOGIN_ENDPOINT = "FakeLoginEndpoint";
const UIREQUIREDERROR = "UiRequiredError";

describe("TeamsUserCredential Tests - Browser", () => {
    const TEST_USER_OBJECT_ID = env.SDK_INTEGRATION_TEST_USER_OBJECT_ID;
    let ssoToken: string;
    let expire_time: number;
    before(async () => {
        [ssoToken, expire_time] = await getSSOToken();
        loadConfiguration({
            authentication: {
                initiateLoginEndpoint: FAKE_LOGIN_ENDPOINT,
                simpleAuthEndpoint: "http://localhost:5000",
                clientId: env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID
            }
        });
        sinon.stub(TeamsUserCredential.prototype, <any>"getSSOToken").callsFake(
            (scopes: string | string[], options?: GetTokenOptions): Promise<AccessToken | null> => {
                return new Promise((resolve, reject) => {
                    resolve({
                        token: ssoToken!,
                        expiresOnTimestamp: expire_time!
                    });
                });
            }
        );
    });
    after(() => {
        sinon.restore();
    });

    it("GetUserInfo should success with SSOToken", async function () {
        const credential: TeamsUserCredential = new TeamsUserCredential();
        const info = await credential.getUserInfo();
        assert.strictEqual(info.preferredUserName, env.SDK_INTEGRATION_TEST_ACCOUNT_NAME);
        assert.strictEqual(info.displayName, env.SDK_INTEGRATION_TEST_ACCOUNT_NAME.split("@", 1)[0]);
        assert.strictEqual(info.objectId, TEST_USER_OBJECT_ID);
    });

    it("GetToken should success with consent scope", async function () {
        const credential: TeamsUserCredential = new TeamsUserCredential();
        // await expect(credential.getToken(["User.Read"])).to.be.eventually.have.property("token");
        const accessToken = await credential.getToken(["User.Read"]);
        const decodedToken = jwtDecode<AADJwtPayLoad>(accessToken!.token);
        assert.strictEqual(decodedToken.aud, "00000003-0000-0000-c000-000000000000");
        assert.isTrue(decodedToken.scp!.startsWith("User.Read"));
    });

    it("GetToken should throw UiRequiredError with unconsent scope", async function () {
        const credential: TeamsUserCredential = new TeamsUserCredential();
        await expect(credential.getToken(["Calendars.Read"])).to.eventually.be.rejectedWith(ErrorWithCode).and.property("code", UIREQUIREDERROR);
    });
});
