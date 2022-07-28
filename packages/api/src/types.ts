// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { Platform, Stage, VsCodeEnv } from "./constants";
import { TokenProvider } from "./utils/login";
import { Context } from "./v2/types";
import { AppManifestProvider } from "./v3/plugins";
import { EnvInfoV3 } from "./v3/types";

export type Json = Record<string, any>;

export type ConfigValue = any;

export type PluginIdentity = string;

export type PluginConfig = ConfigMap;
export type ReadonlyPluginConfig = ReadonlyMap<string, ConfigValue>;

export type SolutionConfig = Map<PluginIdentity, PluginConfig>;
export type ReadonlySolutionConfig = ReadonlyMap<PluginIdentity, ReadonlyPluginConfig>;

/**
 * Definition of option item in single selection or multiple selection
 */
export interface OptionItem {
  /**
   * unique identifier of the option item in the option list
   */
  id: string;
  /**
   * display name
   */
  label: string;
  /**
   * short description
   */
  description?: string;
  /**
   * detailed description
   */
  detail?: string;
  /**
   * customized user data, which is not displayed
   */
  data?: unknown;
  /**
   * CLI display name. CLI will use `cliName` as display name, and use `id` instead if `cliName` is undefined.
   */
  cliName?: string;
  /**
   * group name. If it's set, separator will be rendered on UI between groups.
   */
  groupName?: string;

  /**
   * Actions that can be made within the item.
   * @param An array of actions
   * @param `icon` is the icon id of the action item
   * @param `tooltip` is the hint of the action item
   * @param `command` is the command name that will be executed when current action triggered
   */
  buttons?: { iconPath: string; tooltip: string; command: string }[];
}

export class ConfigMap extends Map<string, ConfigValue> {
  getString(k: string, defaultValue?: string): string | undefined {
    const v = super.get(k);
    if (!v) return defaultValue;
    return String(v);
  }
  getBoolean(k: string, defaultValue?: boolean): boolean | undefined {
    const v = super.get(k);
    if (!v) return defaultValue;
    return Boolean(v);
  }
  getNumber(k: string, defaultValue?: number): number | undefined {
    const v = super.get(k);
    if (!v) return defaultValue;
    return Number(v);
  }
  getStringArray(k: string, defaultValue?: string[]): string[] | undefined {
    const v = super.get(k);
    if (!v) return defaultValue;
    return v as string[];
  }
  getNumberArray(k: string, defaultValue?: number[]): number[] | undefined {
    const v = super.get(k);
    if (!v) return defaultValue;
    return v as number[];
  }
  getBooleanArray(k: string, defaultValue?: boolean[]): boolean[] | undefined {
    const v = super.get(k);
    if (!v) return defaultValue;
    return v as boolean[];
  }
  getOptionItem(k: string, defaultValue?: OptionItem): OptionItem | undefined {
    const v = super.get(k);
    if (!v) return defaultValue;
    if (v instanceof ConfigMap) {
      return v.toJSON() as OptionItem;
    } else {
      return v as OptionItem;
    }
  }
  getOptionItemArray(k: string, defaultValue?: OptionItem[]): OptionItem[] | undefined {
    const v = super.get(k);
    if (!v) return defaultValue;
    return v as OptionItem[];
  }

  toJSON(): Json {
    const out: Json = {};
    for (const entry of super.entries()) {
      if (entry[1] instanceof ConfigMap) {
        out[entry[0]] = entry[1].toJSON();
      } else {
        out[entry[0]] = entry[1];
      }
    }
    return out;
  }

  public static fromJSON(obj?: Json): ConfigMap | undefined {
    if (!obj) return undefined;
    const map = new ConfigMap();
    for (const entry of Object.entries(obj)) {
      if (typeof entry[1] !== "object") {
        map.set(entry[0], entry[1]);
      } else {
        map.set(entry[0], this.fromJSON(entry[1]));
      }
    }
    return map;
  }
  constructor(entries?: readonly (readonly [string, ConfigValue])[] | null) {
    super(entries);
    Object.setPrototypeOf(this, ConfigMap.prototype);
  }
}

export function mergeConfigMap(lhs?: ConfigMap, rhs?: ConfigMap): ConfigMap | undefined {
  if (!lhs) {
    return rhs;
  }
  if (!rhs) {
    return lhs;
  }
  return new ConfigMap([...lhs.entries(), ...rhs.entries()]);
}

// eslint-disable-next-line @typescript-eslint/ban-types
export type Void = {};
export const Void = {};

export type ResourceTemplate = Record<string, ConfigValue>;

export type ResourceTemplates = {
  [k: string]: ResourceTemplate | undefined;
};

export type ResourceConfig = ResourceTemplate;

export type ResourceConfigs = ResourceTemplates;

export type ReadonlyResourceConfig = Readonly<ResourceConfig>;

export type ReadonlyResourceConfigs = Readonly<{
  [k: string]: ReadonlyResourceConfig | undefined;
}>;

/**
 * environment meta data
 */
export interface EnvMeta {
  name: string;
  local: boolean;
  sideloading: boolean;
}

/**
 * project static settings
 */
export interface ProjectSettings {
  appName: string;
  version?: string;
  projectId: string;
  programmingLanguage?: string;
  defaultFunctionName?: string;
  solutionSettings?: SolutionSettings;
  isFromSample?: boolean;
  isM365?: boolean;
  /**
   * pluginSettings is used for plugin settings irrelevant to environments
   */
  pluginSettings?: Json;
}

/**
 * solution settings
 */
export interface SolutionSettings extends Json {
  name: string;
  /**
   * solution settings schema version
   */
  version?: string;
}

export interface AzureSolutionSettings extends SolutionSettings {
  hostType: string;
  capabilities: string[];
  azureResources: string[];
  activeResourcePlugins: string[];
}

/**
 * local debug settings
 */
export interface LocalSettings {
  teamsApp?: ConfigMap;
  auth?: ConfigMap;
  frontend?: ConfigMap;
  backend?: ConfigMap;
  bot?: ConfigMap;
}

/**
 * project dynamic states
 */
export interface ProjectStates {
  solution: Record<string, ConfigValue>;
  resources: {
    [k: string]: Record<string, ConfigValue>;
  };
}
export interface Inputs extends Json {
  projectPath?: string;
  targetEnvName?: string;
  sourceEnvName?: string;
  targetResourceGroupName?: string;
  targetResourceLocationName?: string; // for vs to create a new resource group
  platform: Platform;
  stage?: Stage;
  vscodeEnv?: VsCodeEnv;
  ignoreConfigPersist?: boolean;
  ignoreEnvInfo?: boolean;
  env?: string;
  projectId?: string;
  existingResources?: string[];
  locale?: string;
  isM365?: boolean;
}

export interface ProjectConfig {
  settings?: ProjectSettings;
  config?: SolutionConfig | Json;
  localSettings?: LocalSettings | Json;
}

export interface ProjectConfigV3 {
  projectSettings: ProjectSettings;
  envInfos: {
    [key: string]: EnvInfoV3;
  };
}

export interface Component extends Json {
  name: string;
  hosting?: string;
  code?: string;
  deployType?: "folder" | "zip";
  language?: string;
  folder?: string;
  artifactFolder?: string;
  build?: boolean;
  provision?: boolean;
  deploy?: boolean;
  connections?: string[];
  sso?: boolean;
}
export interface ProjectSettingsV3 extends ProjectSettings {
  components: Component[];
}
export interface ContextV3 extends Context {
  manifestProvider: AppManifestProvider;
  projectSetting: ProjectSettingsV3;
  envInfo?: EnvInfoV3;
  tokenProvider?: TokenProvider;
  projectPath?: string;
}
export interface ResourceContextV3 extends ContextV3 {
  envInfo: EnvInfoV3;
  tokenProvider: TokenProvider;
}
export type MaybePromise<T> = T | Promise<T>;
