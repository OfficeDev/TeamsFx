// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { excelJsApiDocs } from "./Excel_ObjsWithAPIs";
import { BM25, BMDocument, prepareDiscription } from "./ragUtil";
import { wordJsApiDocs } from "./word_docs";

export type DocumentMetadata = {
  description: string;
  codeSample: string;
};

export type API = {
  name: string;
  description: string;
  kind: string;
  signature: string;
  examples: string[];
};

//descrepted, only for wordJsApiDocs formatting
export function prepareDocs(): [string[], Map<string, API>] {
  const docs: string[] = [];
  const docsWithMetadata: Map<string, API> = new Map();
  excelJsApiDocs.forEach((object) => {
    object.apiList.forEach((api) => {
      if (api.description === undefined) {
        return;
      }
      const cleanDescription = prepareDiscription(api.description).join(" ");
      docs.push(cleanDescription);
      docsWithMetadata.set(cleanDescription, api);
    });
  });
  return [docs, docsWithMetadata];
}

// for new json array templates
export function prepareExamples(
  docs: DocumentMetadata[]
): [string[], Map<string, DocumentMetadata>] {
  const docsWithMetadata: Map<string, DocumentMetadata> = new Map();
  const cleanDocs: string[] = [];
  docs.forEach((doc) => {
    const cleanDescription = prepareDiscription(doc.description).join(" ");
    cleanDocs.push(cleanDescription);
    docsWithMetadata.set(cleanDescription, doc);
  });
  return [cleanDocs, docsWithMetadata];
}

//export function getStepsByResponse(response: string): string[] {
//  let steps: string[] = [];
//  const responseJson = parseCopilotResponseMaybeWithStrJson(response);
//  if (responseJson && responseJson.response) {
//    if (Array.isArray(responseJson.response)) {
//      responseJson.response.forEach((element: any) => {
//        if (element.type === "init_plan") {
//          steps = element.content.split(/\d\.\s*/).filter((step) => step !== "");
//        }
//      });
//    }
//  }
//  return steps;
//}

function splitStep(step: string): string[] {
  return step
    .replace(/[^a-zA-Z0-9 ]/g, "")
    .toLowerCase()
    .split(" ");
}

export function searchTopKBySteps(
  steps: string[],
  docs: string[],
  docsWithMetadata: Map<string, API>,
  topK = 3
): API[] {
  const matchedAPIs: Set<string> = new Set();
  steps.forEach((step) => {
    const results = BM25(docs, prepareDiscription(step), undefined, (firstEl, secondEl) => {
      return secondEl.score - firstEl.score;
    }) as BMDocument[];

    // first only take the topK results
    results.slice(0, topK < results.length ? topK : results.length).forEach((result) => {
      if (result.score > 2) {
        matchedAPIs.add(result.document);
      }
    });
  });
  const apiSample: API[] = Array.from(matchedAPIs)
    .map((api) => docsWithMetadata.get(api))
    .filter((api) => api !== undefined) as API[];
  return apiSample;
}

export function searchTopKByqueryAndDocs(
  query: string,
  docs: string[],
  docsWithMetadata: Map<string, DocumentMetadata>,
  topK = 2,
  scoreThreshold = 2
): DocumentMetadata[] {
  const results = BM25(docs, prepareDiscription(query), undefined, (firstEl, secondEl) => {
    return secondEl.score - firstEl.score;
  }) as BMDocument[];
  const matchedDocs: DocumentMetadata[] = [];
  results.slice(0, topK < results.length ? topK : results.length).forEach((result) => {
    if (result.score >= scoreThreshold) {
      matchedDocs.push(docsWithMetadata.get(result.document) as DocumentMetadata);
    }
  });
  return matchedDocs;
}
