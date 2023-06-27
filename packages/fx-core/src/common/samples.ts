import axios from "axios";
import { sendRequestWithTimeout } from "../component/generator/utils";
import sampleConfigV3 from "./samples-config-v3.json";
import { isVideoFilterEnabled } from "./tools";

class configInfo {
  static readonly owner = "OfficeDev";
  static readonly repo = "TeamsFx-Samples";
  static readonly tree = "v2.2.0";
  static readonly file = ".config/samples-config-v3.json";
}

export interface SampleInfo {
  id: string;
  title: string;
  shortDescription: string;
  fullDescription: string;
  tags: string[];
  time: string;
  configuration: string;
  link: string;
  suggested: boolean;
  url: string;
  relativePath?: string;
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
        link:
          (sample as any).packageLink ?? (this.sampleConfigs ?? sampleConfigV3).defaultPackageLink,
        suggested: sample.suggested,
        url:
          (sample as any).relativePath && (sample as any).url
            ? (sample as any).url
            : `${(this.sampleConfigs ?? sampleConfigV3).baseUrl}${sample.id}`,
        relativePath: (sample as any).relativePath,
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
}

export const sampleProvider = new SampleProvider();
