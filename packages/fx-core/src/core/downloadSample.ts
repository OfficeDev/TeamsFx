// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { assembleError, err, FxError, Inputs, ok, Platform, Result } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import axios, { AxiosResponse } from "axios";
import * as fs from "fs-extra";
import { glob } from "glob";
import * as path from "path";
import * as uuid from "uuid";
import { TOOLS } from "./globalVars";
import { sampleProvider } from "../common/samples";
import {
  Component,
  sendTelemetryErrorEvent,
  sendTelemetryEvent,
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../common/telemetry";
import { FetchSampleError, InvalidInputError, ProjectFolderInvalidError } from "./error";
import { loadProjectSettings } from "./middleware/projectSettingsLoader";
import { CoreQuestionNames, QuestionRootFolder } from "./question";
import { CoreHookContext } from "./types";

export async function fetchCodeZip(
  url: string,
  sampleId: string
): Promise<Result<AxiosResponse<any> | undefined, FxError>> {
  let retries = 3;
  let result = undefined;
  const error = new FetchSampleError(sampleId);
  while (retries > 0) {
    retries--;
    try {
      result = await axios.get(url, {
        responseType: "arraybuffer",
      });
      if (result.status === 200 || result.status === 201) {
        return ok(result);
      }
    } catch (e) {
      await new Promise<void>((resolve: () => void): NodeJS.Timer => setTimeout(resolve, 10000));
      if (e.response) {
        error.message += `, status code: ${e.response.status}`;
      } else if (e.request) {
        if (e.code === "ENOTFOUND") {
          error.message += ". Network issue, please check your network connectivity";
        } else {
          error.message += `. Request: ${e.request} failed with error message ${e.message}`;
        }
      } else {
        error.message += `. ${e.message}`;
      }
    }
  }
  return err(error);
}

export async function saveFilesRecursively(
  zip: AdmZip,
  appFolder: string,
  dstPath: string
): Promise<void> {
  await Promise.all(
    zip
      .getEntries()
      .filter(
        (entry) =>
          !entry.isDirectory &&
          entry.entryName.includes(appFolder) &&
          entry.entryName.split("/").includes(appFolder)
      )
      .map(async (entry) => {
        const entryPath = entry.entryName.substring(
          entry.entryName.indexOf(appFolder) + appFolder.length
        );
        const filePath = path.join(dstPath, entryPath);
        await fs.ensureDir(path.dirname(filePath));
        await fs.writeFile(filePath, entry.getData());
      })
  );
}

export async function downloadSampleHook(sampleId: string, sampleAppPath: string): Promise<void> {
  // A temporary solution to avoid duplicate componentId
  if (sampleId === "todo-list-SPFx") {
    const originalId = "c314487b-f51c-474d-823e-a2c3ec82b1ff";
    const componentId = uuid.v4();
    glob.glob(`${sampleAppPath}/**/*.json`, { nodir: true, dot: true }, async (err, files) => {
      await Promise.all(
        files.map(async (file) => {
          let content = (await fs.readFile(file)).toString();
          const reg = new RegExp(originalId, "g");
          content = content.replace(reg, componentId);
          await fs.writeFile(file, content);
        })
      );
    });
  }
}

export async function downloadSample(
  inputs: Inputs,
  ctx?: CoreHookContext
): Promise<Result<string, FxError>> {
  let fxError;
  const progress = TOOLS.ui.createProgressBar("Fetch sample app", 3);
  progress.start();
  const telemetryProperties: any = {
    [TelemetryProperty.Success]: TelemetrySuccess.Yes,
    module: "fx-core",
  };
  try {
    const folder = inputs[QuestionRootFolder.name] as string;
    try {
      await fs.ensureDir(folder);
    } catch (e) {
      throw new ProjectFolderInvalidError(folder);
    }
    const sampleId = inputs[CoreQuestionNames.Samples] as string;
    if (!(sampleId && folder)) {
      throw InvalidInputError(`invalid answer for '${CoreQuestionNames.Samples}'`, inputs);
    }
    telemetryProperties[TelemetryProperty.SampleAppName] = sampleId;
    const samples = sampleProvider.SampleCollection.samples.filter(
      (sample) => sample.id.toLowerCase() === sampleId.toLowerCase()
    );
    if (samples.length === 0) {
      throw InvalidInputError(`invalid sample id: '${sampleId}'`, inputs);
    }
    const sample = samples[0];
    const url = sample.link as string;
    let sampleAppPath = path.resolve(folder, sampleId);
    if ((await fs.pathExists(sampleAppPath)) && (await fs.readdir(sampleAppPath)).length > 0) {
      let suffix = 1;
      while (await fs.pathExists(sampleAppPath)) {
        sampleAppPath = `${folder}/${sampleId}_${suffix++}`;
      }
    }
    progress.next(`Downloading from ${url}`);
    const fetchRes = await fetchCodeZip(url, sample.id);
    if (fetchRes.isErr()) {
      throw fetchRes.error;
    } else if (!fetchRes.value) {
      throw new FetchSampleError(sample.id);
    }
    progress.next("Unzipping the sample package");
    await saveFilesRecursively(new AdmZip(fetchRes.value.data), sampleId, sampleAppPath);
    await downloadSampleHook(sampleId, sampleAppPath);
    progress.next("Update project settings");
    const loadInputs: Inputs = {
      ...inputs,
      projectPath: sampleAppPath,
    };
    const projectSettingsRes = await loadProjectSettings(loadInputs, true);
    if (projectSettingsRes.isOk()) {
      const projectSettings = projectSettingsRes.value;
      projectSettings.projectId = inputs.projectId ? inputs.projectId : uuid.v4();
      projectSettings.isFromSample = true;
      inputs.projectId = projectSettings.projectId;
      telemetryProperties[TelemetryProperty.ProjectId] = projectSettings.projectId;
      if (ctx) ctx.projectSettings = projectSettings;
      inputs.projectPath = sampleAppPath;
    } else {
      telemetryProperties[TelemetryProperty.ProjectId] =
        "unknown, failed to set projectId in projectSettings.json";
    }
    progress.end(true);
    sendTelemetryEvent(Component.core, TelemetryEvent.DownloadSample, telemetryProperties);
    return ok(sampleAppPath);
  } catch (e) {
    fxError = assembleError(e);
    progress.end(false);
    telemetryProperties[TelemetryProperty.Success] = TelemetrySuccess.No;
    sendTelemetryErrorEvent(
      Component.core,
      TelemetryEvent.DownloadSample,
      fxError,
      telemetryProperties
    );
    return err(fxError);
  }
}
