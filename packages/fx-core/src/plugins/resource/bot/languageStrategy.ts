// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as utils from "./utils/common";
import { ProgrammingLanguage } from "./enums/programmingLanguage";
import { TemplateManifest } from "./utils/templateManifest";
import { DeployConfigs, TemplateProjectsConstants } from "./constants";
import { Commands } from "./resources/strings";

import * as appService from "@azure/arm-appservice";
import { NameValuePair } from "@azure/arm-appservice/esm/models";
import AdmZip from "adm-zip";
import { CommandExecutionException, SomethingMissingException } from "./exceptions";
import { downloadByUrl } from "./utils/downloadByUrl";
import * as path from "path";
import * as fs from "fs-extra";

export class LanguageStrategy {
    public static async getTemplateProjectZip(programmingLanguage: ProgrammingLanguage, groupName: string): Promise<AdmZip> {
        try {
            const zipUrl = await LanguageStrategy.getTemplateProjectZipUrl(programmingLanguage, groupName);
            const zipBuffer = await downloadByUrl(zipUrl);
            return new AdmZip(zipBuffer);
        } catch (e) {
            // ToDo: Add log for debug.
            const fallbackFilePath = await LanguageStrategy.generateLocalFallbackFilePath(programmingLanguage, groupName);
            return new AdmZip(fallbackFilePath);
        }
    }

    public static async getTemplateProjectZipUrl(programmingLanguage: ProgrammingLanguage, groupName: string): Promise<string> {
        const manifest: TemplateManifest = await TemplateManifest.fromUrl(
            TemplateProjectsConstants.NEWEST_MANIFEST_URL,
        );

        return manifest.getNewestTemplateUrl(programmingLanguage, groupName);
    }

    public static getSiteEnvelope(
        language: ProgrammingLanguage,
        appServicePlanName: string,
        location: string,
        appSettings?: NameValuePair[],
    ): appService.WebSiteManagementModels.Site {
        const siteEnvelope: appService.WebSiteManagementModels.Site = {
            location: location,
            serverFarmId: appServicePlanName,
            siteConfig: {
                appSettings: [],
            },
        };

        if (!appSettings) {
            appSettings = [];
        }

        appSettings.push({
            name: "SCM_DO_BUILD_DURING_DEPLOYMENT",
            value: "true",
        });

        appSettings.push({
            name: "WEBSITE_NODE_DEFAULT_VERSION",
            value: "12.13.0",
        });

        appSettings.forEach((p: NameValuePair) => {
            siteEnvelope?.siteConfig?.appSettings?.push(p);
        });

        return siteEnvelope;
    }

    public static async buildAndZipPackage(programmingLanguage: ProgrammingLanguage, packDir: string, unPackFlag?: boolean): Promise<Buffer> {
        if (programmingLanguage === ProgrammingLanguage.TypeScript) {
            //Typescript needs tsc build before deploy because of windows app server. other languages don"t need it.
            try {
                await utils.execute("npm install", packDir);
                await utils.execute("npm run build", packDir);
            } catch (e) {
                throw new CommandExecutionException(`${Commands.NPM_INSTALL},${Commands.NPM_BUILD}`, e.message, e);
            }
        }

        if (!unPackFlag && programmingLanguage === ProgrammingLanguage.JavaScript) {
            // Since teamfx package is in private registry, have to npm install to pack the node_modulre folder.
            try {
                await utils.execute("npm install", packDir);
            } catch (e) {
                throw new CommandExecutionException(`${Commands.NPM_INSTALL}`, e.message, e);
            }
        }

        return utils.zipAFolder(packDir, unPackFlag ? DeployConfigs.UN_PACK_DIRS : []);
    }

    private static async generateLocalFallbackFilePath(programmingLanguage: ProgrammingLanguage, groupName: string): Promise<string> {
        const fxCorePath = path.join(__dirname, "..", "..", "..", "..");
        const targetFilePath = path.join(fxCorePath, "templates", "plugins", "resource", "bot", `${groupName}.${programmingLanguage}.${TemplateProjectsConstants.DEFAULT_SCENARIO_NAME}.zip`);

        const targetExisted = await fs.pathExists(targetFilePath);
        if (!targetExisted) {
            throw new SomethingMissingException(targetFilePath);
        }

        return targetFilePath;
    }
}
