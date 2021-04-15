// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as utils from "./utils/common";
import { AxiosInstance, default as axios } from "axios";

import { AADRegistrationConstants, Retry } from "./constants";
import { IAADApplication } from "./appStudio/interfaces/IAADApplication";
import * as AppStudio from "./appStudio/appStudio";
import { ProvisionException } from "./exceptions";
import { CommonStrings } from "./resources/strings";
import { Logger } from "./logger";
import { BotAuthCredential } from "./botAuthCredential";

export async function registerAADAppAndGetSecretByGraph(graphToken: string, displayName: string): Promise<BotAuthCredential> {
    const axiosInstance: AxiosInstance = axios.create({
        headers: {
            post: {
                Authorization: `Bearer ${graphToken}`,
            },
        },
    });

    const result = new BotAuthCredential();

    // 1. Register a new AAD App.
    let regResponse = undefined;
    try {
        regResponse = await axiosInstance.post(`${AADRegistrationConstants.GRAPH_REST_BASE_URL}/applications`, {
            displayName: displayName,
            signInAudience: AADRegistrationConstants.AZURE_AD_MULTIPLE_ORGS
        });
    } catch (e) {
        throw new ProvisionException(CommonStrings.AAD_APP, e);
    }

    if (!regResponse || !utils.isHttpCodeOkOrCreated(regResponse.status)) {
        throw new ProvisionException(CommonStrings.AAD_APP);
    }

    result.clientId = regResponse.data.appId;
    result.objectId = regResponse.data.id;

    // 2. Generate client secret.

    let genResponse = undefined;
    try {
        genResponse = await axiosInstance.post(
            `${AADRegistrationConstants.GRAPH_REST_BASE_URL}/applications/${result.objectId}/addPassword`,
            {
                passwordCredential: {
                    displayName: "default",
                },
            },
        );
    } catch (e) {
        throw new ProvisionException(CommonStrings.AAD_CLIENT_SECRET, e);
    }

    if (!genResponse || !genResponse.data) {
        throw new ProvisionException(CommonStrings.AAD_CLIENT_SECRET);
    }

    result.clientSecret = genResponse.data.secretText;
    return result;
}

export async function registerAADAppAndGetSecretByAppStudio(appStudioToken: string, displayName: string): Promise<BotAuthCredential> {

    const result = new BotAuthCredential();

    await AppStudio.init(appStudioToken);

    const appConfig: IAADApplication = {
        displayName: displayName
    };

    const app = await AppStudio.createAADApp(appConfig);
    result.clientId = app.id;
    result.objectId = app.objectId;

    const password = await AppStudio.createAADAppPassword(app.objectId);

    if (!password || !password.value) {
        throw new ProvisionException(CommonStrings.AAD_CLIENT_SECRET);
    }

    result.clientSecret = password.value;

    return result;
}