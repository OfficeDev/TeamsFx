// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  v2,
  AzureSolutionSettings,
  Plugin,
  ProjectSettings,
  UserError,
} from "@microsoft/teamsfx-api";
import "reflect-metadata";
import { SolutionError, SolutionSource } from "./constants";
export const ResourcePlugins = {};

export const ResourcePluginsV2 = {};

/**
 * @returns all registered resource plugins
 */
export function getAllResourcePlugins(): Plugin[] {
  const plugins: Plugin[] = [];
  return plugins;
}

/**
 * @returns all resource plugins implemented with v2 API
 */
export function getAllV2ResourcePlugins(): v2.ResourcePlugin[] {
  const plugins: v2.ResourcePlugin[] = [];
  return plugins;
}

/**
 * @returns all registered resource plugin map
 */
export function getAllResourcePluginMap(): Map<string, Plugin> {
  const map = new Map<string, Plugin>();
  const allPlugins = getAllResourcePlugins();
  for (const p of allPlugins) {
    map.set(p.name, p);
  }
  return map;
}

/**
 * @returns a map from plugin name to resource plugin v2
 */
export function getAllV2ResourcePluginMap(): Map<string, v2.ResourcePlugin> {
  const map = new Map<string, v2.ResourcePlugin>();
  const allPlugins = getAllV2ResourcePlugins();
  for (const p of allPlugins) {
    map.set(p.name, p);
  }
  return map;
}

/**
 * return activated resource plugin according to solution settings
 * @param solutionSettings Azure solution settings
 * @returns activated resource plugins
 */
export function getActivatedResourcePlugins(solutionSettings: AzureSolutionSettings): Plugin[] {
  const activatedPlugins = getAllResourcePlugins().filter(
    (p) => p.activate && p.activate(solutionSettings) === true
  );
  if (activatedPlugins.length === 0) {
    throw new UserError(
      SolutionSource,
      SolutionError.NoResourcePluginSelected,
      "No plugin selected"
    );
  }
  return activatedPlugins;
}

/**
 * return activated resource plugin according to solution settings
 * @param projectSettings project settings
 * @returns activated resource plugins
 */
export function getActivatedV2ResourcePlugins(
  projectSettings: ProjectSettings
): v2.ResourcePlugin[] {
  const activeResourcePlugins = (projectSettings.solutionSettings as AzureSolutionSettings)
    ?.activeResourcePlugins;
  if (!activeResourcePlugins) return [];
  const activatedPlugins = getAllV2ResourcePlugins().filter((p) =>
    activeResourcePlugins.includes(p.name)
  );
  return activatedPlugins;
}
