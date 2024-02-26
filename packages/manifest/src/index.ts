// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TeamsAppManifest, IComposeExtension } from "./manifest";
import fs from "fs-extra";
import Ajv from "ajv-draft-04";
import { JSONSchemaType } from "ajv";
import { DevPreviewSchema } from "./devPreviewManifest";
import { ManifestCommonProperties } from "./ManifestCommonProperties";
import { SharePointAppId } from "./constants";
import fetch from "node-fetch";

export * from "./manifest";
export * as devPreview from "./devPreviewManifest";
export * from "./pluginManifest";
export type TeamsAppManifestJSONSchema = JSONSchemaType<TeamsAppManifest>;
export type DevPreviewManifestJSONSchema = JSONSchemaType<DevPreviewSchema>;

export type Manifest = TeamsAppManifest | DevPreviewSchema;

export class ManifestUtil {
  /**
   * Loads the manifest from the given path without validating its schema.
   *
   * @param path - The path to the manifest file.
   * @throws Will propagate any error thrown by the fs-extra#readJson.
   *
   * @returns The Manifest Object.
   */
  static async loadFromPath<T extends Manifest = TeamsAppManifest>(path: string): Promise<T> {
    return fs.readJson(path);
  }

  /**
   * Writes the manifest object to the given path.
   *
   * @param path - Where to write
   * @param manifest - Manifest object to be saved
   * @throws Will propagate any error thrown by the fs-extra#writeJson.
   *
   */
  static async writeToPath<T extends Manifest = TeamsAppManifest>(
    path: string,
    manifest: T
  ): Promise<void> {
    return fs.writeJson(path, manifest, { spaces: 4 });
  }

  /**
   * Validate manifest against json schema.
   *
   * @param manifest - Manifest object to be validated
   * @param schema - teams-app-manifest schema
   * @returns An empty array if it passes validation, or an array of error string otherwise.
   */
  static validateManifestAgainstSchema<T extends Manifest = TeamsAppManifest>(
    manifest: T,
    schema: JSONSchemaType<T>
  ): Promise<string[]> {
    const ajv = new Ajv({ formats: { uri: true }, allErrors: true, strictTypes: false });
    const validate = ajv.compile(schema);
    const valid = validate(manifest);
    if (!valid && validate.errors) {
      return Promise.resolve(
        validate.errors?.map((error) => `${error.instancePath} ${error.message || ""}`)
      );
    } else {
      return Promise.resolve([]);
    }
  }

  static async fetchSchema<T extends Manifest = TeamsAppManifest>(
    manifest: T
  ): Promise<JSONSchemaType<T>> {
    if (!manifest.$schema) {
      throw new Error("Manifest does not have a $schema property");
    }
    let result: JSONSchemaType<T>;
    try {
      const res = await fetch(manifest.$schema);
      result = (await res.json()) as JSONSchemaType<T>;
    } catch (e: unknown) {
      if (e instanceof Error) {
        throw new Error(`Failed to get manifest at url ${manifest.$schema} due to: ${e.message}`);
      } else {
        throw new Error(`Failed to get manifest at url ${manifest.$schema} due to: unknown error`);
      }
    }
    return result;
  }

  /**
   * Validate manifest against {@link TeamsAppManifest#$schema}.
   *
   * @param manifest - Manifest object to be validated
   * @throws Will throw if {@link TeamsAppManifest#$schema} is undefined, not valid
   *         or there is any network failure when getting the schema.
   *
   * @returns An empty array if schema validation passes, or an array of error string otherwise.
   */
  static async validateManifest<T extends Manifest = TeamsAppManifest>(
    manifest: T
  ): Promise<string[]> {
    const schema = await this.fetchSchema(manifest);
    return ManifestUtil.validateManifestAgainstSchema(manifest, schema);
  }

  /**
   * Parse the manifest and get properties
   * @param manifest
   */
  static parseCommonProperties<T extends Manifest = TeamsAppManifest>(
    manifest: T
  ): ManifestCommonProperties {
    const capabilities: string[] = [];
    if (manifest.staticTabs && manifest.staticTabs.length > 0) {
      capabilities.push("staticTab");
    }
    if (manifest.configurableTabs && manifest.configurableTabs.length > 0) {
      capabilities.push("configurableTab");
    }
    if (manifest.bots && manifest.bots.length > 0) {
      capabilities.push("Bot");
    }
    if (manifest.composeExtensions) {
      capabilities.push("MessageExtension");
    }

    const properties: ManifestCommonProperties = {
      id: manifest.id,
      version: manifest.version,
      capabilities: capabilities,
      manifestVersion: manifest.manifestVersion,
      isApiME: false,
      isSPFx: false,
      isApiPlugin: false,
    };

    // If it's copilot plugin app
    if (
      manifest.composeExtensions &&
      manifest.composeExtensions.length > 0 &&
      (manifest.composeExtensions[0] as IComposeExtension).composeExtensionType == "apiBased"
    ) {
      properties.isApiME = true;
    }

    // If it's SPFx app
    if (
      manifest.webApplicationInfo &&
      manifest.webApplicationInfo.id &&
      manifest.webApplicationInfo.id == SharePointAppId
    ) {
      properties.isSPFx = true;
    }

    if ((manifest as TeamsAppManifest).apiPlugins) {
      const apiPlugins = (manifest as TeamsAppManifest).apiPlugins;
      if (apiPlugins && apiPlugins.length > 0 && apiPlugins[0].pluginFile)
        properties.isApiPlugin = true;
    }

    return properties;
  }

  /**
   * Parse the manifest and get telemetry propreties e.g. appId, capabilities etc.
   * @param manifest
   * @returns Telemetry properties
   */
  static parseCommonTelemetryProperties(manifest: TeamsAppManifest): { [p: string]: string } {
    const properties = ManifestUtil.parseCommonProperties(manifest);

    const telemetryProperties: { [p: string]: string } = {};
    const propertiesMap = new Map<string, any>(Object.entries(properties));
    propertiesMap.forEach((value, key) => {
      if (Array.isArray(value)) {
        telemetryProperties[key] = value.join(";");
      } else {
        telemetryProperties[key] = value;
      }
    });

    return telemetryProperties;
  }
}
