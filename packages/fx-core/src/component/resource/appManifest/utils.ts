// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { FxError, ok, Result, TeamsAppManifest } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import * as path from "path";
import "reflect-metadata";
import { getProjectTemplatesFolderPath } from "../../../common/utils";
import { isV3 } from "../../../core/globalVars";
import { convertManifestTemplateToV2, convertManifestTemplateToV3 } from "../../migrate";

export async function readAppManifest(
  projectPath: string
): Promise<Result<TeamsAppManifest, FxError>> {
  const filePath = await getTeamsAppManifestPath(projectPath);
  const content = await fs.readFile(filePath, { encoding: "utf-8" });
  const contentV3 = isV3() ? convertManifestTemplateToV3(content) : content;
  const manifest = JSON.parse(contentV3) as TeamsAppManifest;
  return ok(manifest);
}

export async function writeAppManifest(
  appManifest: TeamsAppManifest,
  projectPath: string
): Promise<Result<undefined, FxError>> {
  const filePath = await getTeamsAppManifestPath(projectPath);
  const content = JSON.stringify(appManifest, undefined, 4);
  const contentV2 = isV3() ? convertManifestTemplateToV2(content) : content;
  await fs.writeFile(filePath, contentV2);
  return ok(undefined);
}

export async function getTeamsAppManifestPath(projectPath: string): Promise<string> {
  const templateFolder = await getProjectTemplatesFolderPath(projectPath);
  const filePath = path.join(templateFolder, "appPackage", "manifest.template.json");
  return filePath;
}
