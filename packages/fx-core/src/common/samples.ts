import axios from "axios";
import { sendRequestWithTimeout } from "../component/generator/utils";
import sampleConfigV3 from "./samples-config-v3.json";
import { isVideoFilterEnabled } from "./featureFlags";
const packageJson = require("../../package.json");

class configInfo {
  static readonly owner = "OfficeDev";
  static readonly repo = "TeamsFx-Samples";
  static readonly tree = "v2.2.0";
  static readonly file = ".config/samples-config-v3.json";
}

class preReleaseConfig {
  static readonly baseUrl = "https://github.com/OfficeDev/TeamsFx-Samples/tree/dev/";
  static readonly defaultPackageLink =
    "https://github.com/OfficeDev/TeamsFx-Samples/archive/refs/heads/dev.zip";
}

export interface SampleInfo {
  id: string;
  title: string;
  shortDescription: string;
  fullDescription: string;
  tags: string[];
  time: string;
  configuration: string;
  suggested: boolean;
  url: string;
}

interface SampleCollection {
  samples: SampleInfo[];
}

class SampleProvider {
  private sampleCollection: SampleCollection | undefined;
  private sampleConfigs: any;

  public async fetchSampleConfig() {
    this.sampleConfigs = undefined;
    try {
      const fileResponse = await sendRequestWithTimeout(
        async () => {
          return await axios.get(
            `https://raw.githubusercontent.com/${configInfo.owner}/${configInfo.repo}/${configInfo.tree}/${configInfo.file}`,
            { responseType: "json" }
          );
        },
        1000,
        3
      );

      if (fileResponse && fileResponse.data) {
        this.sampleConfigs = fileResponse.data;
      }
    } catch (e) {
      this.sampleConfigs = undefined;
    }
  }
  public get SampleCollection(): SampleCollection {
    const samples = (this.sampleConfigs ?? sampleConfigV3).samples.map((sample: any) => {
      return {
        id: sample.id,
        title: sample.title,
        shortDescription: sample.shortDescription,
        fullDescription: sample.fullDescription,
        tags: sample.tags,
        time: sample.time,
        configuration: sample.configuration,
        suggested: sample.suggested,
        url: (sample as any).url
          ? (sample as any).url
          : `${
              (this.isStableRelease() ? this.sampleConfigs ?? sampleConfigV3 : preReleaseConfig)
                .baseUrl
            }${sample.id}`,
      } as SampleInfo;
    });

    // remove video filter sample app if feature flag is disabled.
    if (!isVideoFilterEnabled()) {
      const videoFilterSampleId = "teams-videoapp-sample";
      const index = samples.findIndex((sample: any) => sample.id === videoFilterSampleId);
      if (index !== -1) {
        samples.splice(index, 1);
      }
    }

    this.sampleCollection = {
      samples,
    };

    return this.sampleCollection;
  }

  private isStableRelease(): boolean {
    const version = packageJson.version;
    if (version.includes("alpha") || version.includes("beta") || version.includes("rc")) {
      return false;
    }
    return true;
  }
}

export const sampleProvider = new SampleProvider();
