// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as path from "path";
import AdmZip from "adm-zip";
import Mustache from "mustache";
import axios from "axios";
import fs from "fs-extra";

import {
  FetchTemplateManifestError,
  FetchTemplatePackageError,
  FrontendPluginError,
  InvalidTemplateManifestError,
  runWithErrorCatchAndThrow,
  TemplateManifestError,
  TemplateZipFallbackError,
  UnknownScaffoldError,
  UnzipTemplateError,
} from "../resources/errors";
import { Constants, FrontendPathInfo as PathInfo } from "../constants";
import { Logger } from "../utils/logger";
import { Messages } from "../resources/messages";
import { PluginContext } from "@microsoft/teamsfx-api";
import { Utils } from "../utils";
import { TemplateInfo, TemplateVariable } from "../resources/templateInfo";
import { selectTag, tagListURL, templateURL } from "../../../../common/templates";
import { TelemetryHelper } from "../utils/telemetry-helper";
import {
  genTemplateRenderReplaceFn,
  removeTemplateExtReplaceFn,
  ScaffoldAction,
  ScaffoldActionName,
  ScaffoldContext,
  scaffoldFromTemplates,
} from "../../../../common/templatesActions";

export type Manifest = {
  [key: string]: {
    [key: string]: {
      [key: string]: {
        version: string;
        url: string;
      }[];
    };
  };
};

export class FrontendScaffold {
  public static async scaffoldFromZipPackage(
    componentPath: string,
    templateInfo: TemplateInfo
  ): Promise<void> {
    await scaffoldFromTemplates({
      group: templateInfo.group,
      lang: templateInfo.language,
      scenario: templateInfo.scenario,
      templatesFolderName: PathInfo.TemplateFolderName,
      dst: componentPath,
      fileNameReplaceFn: removeTemplateExtReplaceFn,
      fileDataReplaceFn: genTemplateRenderReplaceFn(templateInfo.variables),
      onActionError: async (action: ScaffoldAction, context: ScaffoldContext, error: Error) => {
        Logger.info(error.toString());
        switch (action.name) {
          case ScaffoldActionName.FetchTemplatesUrlWithTag:
          case ScaffoldActionName.FetchTemplatesZipFromUrl:
            TelemetryHelper.sendScaffoldFallbackEvent(new TemplateManifestError(error.message));
            Logger.info(Messages.FailedFetchTemplate);
            break;
          case ScaffoldActionName.FetchTemplateZipFromLocal:
            throw new TemplateZipFallbackError();
          case ScaffoldActionName.Unzip:
            throw new UnzipTemplateError();
          default:
            throw new UnknownScaffoldError();
        }
      },
    });
  }

  public static async fetchTemplateTagList(url: string): Promise<string> {
    const result = await runWithErrorCatchAndThrow(
      new FetchTemplateManifestError(),
      async () =>
        await Utils.requestWithRetry(async () => {
          return axios.get(url, {
            timeout: Constants.RequestTimeoutInMS,
          });
        }, Constants.ScaffoldTryCounts)
    );
    if (!result) {
      throw new FetchTemplateManifestError();
    }
    return result.data;
  }

  public static async getTemplateURL(
    manifestUrl: string,
    templateBaseName: string
  ): Promise<string> {
    const tags: string = await this.fetchTemplateTagList(manifestUrl);
    const selectedTag = selectTag(tags.replace(/\r/g, Constants.EmptyString).split("\n"));
    if (!selectedTag) {
      throw new InvalidTemplateManifestError(templateBaseName);
    }
    return templateURL(selectedTag, templateBaseName);
  }

  public static async fetchZipFromUrl(url: string): Promise<AdmZip> {
    const result = await runWithErrorCatchAndThrow(
      new FetchTemplatePackageError(),
      async () =>
        await Utils.requestWithRetry(async () => {
          return axios.get(url, {
            responseType: "arraybuffer",
            timeout: Constants.RequestTimeoutInMS,
          });
        }, Constants.ScaffoldTryCounts)
    );

    if (!result) {
      throw new FetchTemplatePackageError();
    }
    return new AdmZip(result.data);
  }

  public static getTemplateZipFromLocal(templateInfo: TemplateInfo): AdmZip {
    const templatePath = templateInfo.localTemplatePath;
    return new AdmZip(templatePath);
  }

  public static async getTemplateZip(
    ctx: PluginContext,
    templateInfo: TemplateInfo
  ): Promise<AdmZip> {
    try {
      const templateUrl = await FrontendScaffold.getTemplateURL(
        tagListURL,
        templateInfo.localTemplateBaseName
      );
      return await FrontendScaffold.fetchZipFromUrl(templateUrl);
    } catch (e) {
      Logger.debug(e.toString());
      Logger.warning(Messages.FailedFetchTemplate);

      if (e instanceof FrontendPluginError) {
        TelemetryHelper.sendScaffoldFallbackEvent(e);
      } else {
        TelemetryHelper.sendScaffoldFallbackEvent(new UnknownScaffoldError());
      }

      return FrontendScaffold.getTemplateZipFromLocal(templateInfo);
    }
  }

  public static fulfill(
    filePath: string,
    data: Buffer,
    variables: TemplateVariable
  ): string | Buffer {
    if (path.extname(filePath) === PathInfo.TemplateFileExt) {
      return Mustache.render(data.toString(), variables);
    }
    return data;
  }

  public static async scaffoldFromZip(
    zip: AdmZip,
    dstPath: string,
    nameReplaceFn?: (filePath: string, data: Buffer) => string,
    dataReplaceFn?: (filePath: string, data: Buffer) => string | Buffer
  ): Promise<void> {
    await Promise.all(
      zip
        .getEntries()
        .filter((entry) => !entry.isDirectory)
        .map(async (entry) => {
          const data: string | Buffer = dataReplaceFn
            ? dataReplaceFn(entry.name, entry.getData())
            : entry.getData();

          const filePath = path.join(
            dstPath,
            nameReplaceFn ? nameReplaceFn(entry.entryName, entry.getData()) : entry.entryName
          );
          await fs.ensureDir(path.dirname(filePath));
          await fs.writeFile(filePath, data);
        })
    );
  }
}
