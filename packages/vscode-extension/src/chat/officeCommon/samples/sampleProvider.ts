// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  CancellationToken,
  ChatRequest,
  LanguageModelChatMessage,
  LanguageModelChatUserMessage,
} from "vscode";
import { BM25, BMDocument } from "../../rag/BM25";
import { OfficeAddinTemplateModelPorvider, WXPAppName } from "./officeAddinTemplateModelPorvider";
import { SampleData } from "./sampleData";
import { prepareDiscription } from "../../rag/ragUtil";

export class SampleProvider {
  private static instance: SampleProvider;

  private constructor() {
    // Private constructor to prevent direct instantiation
  }

  public static getInstance(): SampleProvider {
    if (!SampleProvider.instance) {
      SampleProvider.instance = new SampleProvider();
    }
    return SampleProvider.instance;
  }

  public async getTopKMostRelevantScenarioSampleCodes(
    request: ChatRequest,
    token: CancellationToken,
    host: string,
    scenario: string,
    k: number
  ): Promise<Map<string, SampleData>> {
    const samples: Map<string, SampleData> = new Map<string, SampleData>();
    try {
      const bm25: BM25 = await OfficeAddinTemplateModelPorvider.getInstance().getBM25Model(
        host as WXPAppName
      );
      const query = prepareDiscription(scenario);
      const documents: BMDocument[] = bm25.search(query, k);

      for (const doc of documents) {
        if (doc.document.metadata) {
          const sampleData = doc.document.metadata as SampleData;
          samples.set(sampleData.name, sampleData);
        }
      }
    } catch (error) {
      console.error(`Failed to fetch BM25 model.`);
    }
    return new Promise<Map<string, SampleData>>((resolve, reject) => {
      resolve(samples);
    });
  }
}
