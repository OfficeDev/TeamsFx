// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks";

import { SampleUrlInfo } from "../component/generator/utils";
import { ErrorContextMW } from "../core/globalVars";
import { AccessGithubError } from "../error/common";
import { FeatureFlagName } from "./constants";

const packageJson = require("../../package.json");

const SampleConfigOwner = "OfficeDev";
const SampleConfigRepo = "TeamsFx-Samples";
const SampleConfigFile = ".config/samples-config-v3.json";
export const SampleConfigTag = "v2.4.0";
// prerelease tag is always using a branch.
export const SampleConfigBranchForPrerelease = "v3";

export interface SampleConfig {
  id: string;
  onboardDate: Date;
  title: string;
  shortDescription: string;
  fullDescription: string;
  // matches the Teams app type when creating a new project
  types: string[];
  tags: string[];
  time: string;
  configuration: string;
  suggested: boolean;
  thumbnailUrl: string;
  gifUrl?: string;
  // maximum TTK & CLI version to run sample
  maximumToolkitVersion?: string;
  maximumCliVersion?: string;
  // these 2 fields are used when external sample is upgraded and breaks in old TTK version.
  minimumToolkitVersion?: string;
  minimumCliVersion?: string;
  downloadUrlInfo: SampleUrlInfo;
}

interface SampleCollection {
  samples: SampleConfig[];
  filterOptions: {
    capabilities: string[];
    languages: string[];
    technologies: string[];
  };
}

type SampleConfigType = {
  samples: Array<Record<string, unknown>>;
  filterOptions: Record<string, Array<string>>;
};

class SampleProvider {
  private sampleCollection: SampleCollection | undefined;

  public get SampleCollection(): Promise<SampleCollection> {
    if (!this.sampleCollection) {
      return this.refreshSampleConfig();
    }
    return Promise.resolve(this.sampleCollection);
  }

  public async refreshSampleConfig(): Promise<SampleCollection> {
    const { samplesConfig, ref } = await this.fetchOnlineSampleConfig();
    this.sampleCollection = this.parseOnlineSampleConfig(samplesConfig, ref);
    return this.sampleCollection;
  }

  private async fetchOnlineSampleConfig() {
    const version: string = packageJson.version;
    const configBranchInEnv = process.env[FeatureFlagName.SampleConfigBranch];
    let samplesConfig: SampleConfigType | undefined;
    let ref = SampleConfigTag;

    // Set default value for branchOrTag
    if (version.includes("alpha")) {
      // daily build version always use 'dev' branch
      ref = "dev";
    } else if (version.includes("beta")) {
      // prerelease build version always use branch head for prerelease.
      ref = SampleConfigBranchForPrerelease;
    } else if (version.includes("rc")) {
      // if there is a breaking change, the tag is not used by any stable version.
      ref = SampleConfigTag;
    } else {
      // stable version uses the head of branch defined by feature flag when available
      ref = SampleConfigTag;
    }

    // Set branchOrTag value if branch in env is valid
    if (configBranchInEnv) {
      try {
        const data = await this.fetchRawFileContent(configBranchInEnv);
        ref = configBranchInEnv;
        samplesConfig = data as SampleConfigType;
      } catch (e: unknown) {}
    }

    if (samplesConfig === undefined) {
      samplesConfig = (await this.fetchRawFileContent(ref)) as SampleConfigType;
    }

    return { samplesConfig, ref };
  }

  @hooks([ErrorContextMW({ component: "SampleProvider" })])
  private parseOnlineSampleConfig(samplesConfig: SampleConfigType, ref: string): SampleCollection {
    const samples =
      samplesConfig?.samples.map((sample) => {
        const isExternal = sample["downloadUrlInfo"] ? true : false;
        let gifUrl =
          sample["gifPath"] !== undefined
            ? `https://raw.githubusercontent.com/${SampleConfigOwner}/${SampleConfigRepo}/${ref}/${
                sample["id"] as string
              }/${sample["gifPath"] as string}`
            : undefined;
        let thumbnailUrl = `https://raw.githubusercontent.com/${SampleConfigOwner}/${SampleConfigRepo}/${ref}/${
          sample["id"] as string
        }/${sample["thumbnailPath"] as string}`;
        if (isExternal) {
          const info = sample["downloadUrlInfo"] as SampleUrlInfo;
          gifUrl =
            sample["gifPath"] !== undefined
              ? `https://raw.githubusercontent.com/${info.owner}/${info.repository}/${info.ref}/${
                  info.dir
                }/${sample["gifPath"] as string}`
              : undefined;
          thumbnailUrl = `https://raw.githubusercontent.com/${info.owner}/${info.repository}/${
            info.ref
          }/${info.dir}/${sample["thumbnailPath"] as string}`;
        }
        return {
          ...sample,
          onboardDate: new Date(sample["onboardDate"] as string),
          downloadUrlInfo: isExternal
            ? sample["downloadUrlInfo"]
            : {
                owner: SampleConfigOwner,
                repository: SampleConfigRepo,
                ref: ref,
                dir: sample["id"] as string,
              },
          gifUrl: gifUrl,
          thumbnailUrl: thumbnailUrl,
        } as SampleConfig;
      }) || [];

    return {
      samples,
      filterOptions: {
        capabilities: samplesConfig?.filterOptions["capabilities"] || [],
        languages: samplesConfig?.filterOptions["languages"] || [],
        technologies: samplesConfig?.filterOptions["technologies"] || [],
      },
    };
  }

  private async fetchRawFileContent(branchOrTag: string): Promise<unknown> {
    const url = `https://raw.githubusercontent.com/${SampleConfigOwner}/${SampleConfigRepo}/${branchOrTag}/${SampleConfigFile}`;
    try {
      const response = await fetch(url, {
        method: "GET",
        headers: {
          Accept: "application/json",
        },
      });
      if (response) {
        return await response.json();
      }
    } catch (e) {
      throw new AccessGithubError(url, "SampleProvider", e);
    }
  }
}

export const sampleProvider = new SampleProvider();
