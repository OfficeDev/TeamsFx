// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import axios from "axios";
import { BM25, DocumentWithmetadata } from "../../rag/BM25";
import { SampleData } from "./sampleData";

export type WXPAppName = "Word" | "Excel" | "PowerPoint";

const sampleDirectoryUrl =
  "https://api.github.com/repos/OfficeDev/Office-agentsamples/contents/scenario-samples/";

export class OfficeAddinTemplateModelPorvider {
  private static instance: OfficeAddinTemplateModelPorvider;

  private bm25Models: { [x: string]: BM25 } = {};

  private constructor() {
    // Private constructor to prevent direct instantiation
  }

  public static getInstance(): OfficeAddinTemplateModelPorvider {
    if (!OfficeAddinTemplateModelPorvider.instance) {
      OfficeAddinTemplateModelPorvider.instance = new OfficeAddinTemplateModelPorvider();
    }
    return OfficeAddinTemplateModelPorvider.instance;
  }

  private async getSamples(name: WXPAppName): Promise<SampleData[]> {
    const returnData: SampleData[] = [];
    const fullUrl = sampleDirectoryUrl + name;
    const directoryResponse = await axios.get(fullUrl, {
      headers: {
        Accept: "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
      },
    });
    if (directoryResponse && directoryResponse.data && directoryResponse.data.length > 0) {
      const dataMap: {
        [x: string]: { Templates: [{ Description: string; SampleCodes: string }] } | null;
      } = {};
      for (const fileInfo of directoryResponse.data) {
        if (fileInfo.download_url) {
          dataMap[fileInfo.download_url] = null;
        }
      }
      const p = [];
      for (const fileInfo of directoryResponse.data) {
        if (fileInfo.download_url) {
          p.push(
            axios
              .get(fileInfo.download_url)
              .then((response) => {
                if (response.data) {
                  dataMap[fileInfo.download_url] = response.data;
                }
              })
              .catch((error) => {
                console.log(error);
              })
          );
        }
      }
      await Promise.all(p);
      for (const fileInfo of directoryResponse.data) {
        if (fileInfo.download_url) {
          if (dataMap[fileInfo.download_url]) {
            const metaData = dataMap[fileInfo.download_url];
            if (metaData && metaData.Templates && metaData.Templates.length > 0) {
              let count = 0;
              for (const template of metaData.Templates) {
                if (template.Description && template.SampleCodes) {
                  count++;
                  const sampleData = new SampleData(
                    (fileInfo.name as string) + "-" + count.toString(),
                    fileInfo.html_url as string,
                    template.SampleCodes,
                    template.Description,
                    "" /* definition*/,
                    "" /* usage */
                  );
                  returnData.push(sampleData);
                }
              }
            }
          }
        }
      }
    }
    return returnData;
  }

  public async getBM25Model(name: WXPAppName): Promise<BM25> {
    if (this.bm25Models[name]) {
      return this.bm25Models[name];
    }
    const samples = await this.getSamples(name);
    const documents: DocumentWithmetadata[] = samples.map((sample) => {
      return {
        documentText: sample.description,
        metadata: sample,
      };
    });
    const bm25 = new BM25(documents);
    this.bm25Models[name] = bm25;
    return bm25;
  }
}
