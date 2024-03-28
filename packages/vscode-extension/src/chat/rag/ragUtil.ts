// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { stemmer } from "./porterStemmer";
import * as stopwords from "./stop_words_english.json";

export function filterStopWords(texts: string[]): string[] {
  return texts.filter((word) => !stopwords.includes(word));
}

export function keepLetters(text: string): string {
  return text.replace(/[^a-zA-Z ]/g, "");
}

export function stemText(texts: string[]): string[] {
  return texts.map(stemmer);
}

//export function lancasterStemText(texts: string[]): string[] {
//  return texts.map(word => lancasterStemmer(word, {}));
//}

export function prepareDiscription(text: string): string[] {
  return stemText(filterStopWords(keepLetters(text).split(" ")));
}

// BM25

export const getWordCount = (corpus: string) => {
  return ((corpus || "").match(/\w+/g) || []).length;
};

/** Number of occurences of a word in a string. */
export const getTermFrequency = (term: string, corpus: string) => {
  return ((corpus || "").match(new RegExp(term, "g")) || []).length;
};

/** Inverse document frequency. */
export const getIDF = (term: string, documents: string[]) => {
  // Number of relevant documents.
  const relevantDocuments = documents.filter((document: string) => document.includes(term)).length;
  return Math.log((documents.length - relevantDocuments + 0.5) / (relevantDocuments + 0.5) + 1);
};

/** Represents a document; useful when sorting results.
 */
export interface BMDocument {
  /** The document is originally scoreed. */
  document: string;
  /** The score that the document recieves. */
  score: number;
}

/** Constants that are free parameters used in BM25, specifically when generating inverse document frequency. */
export interface BMConstants {
  /** Free parameter. Is 0.75 by default.  */
  b?: number;
  /** Free parameter. Is 1.2 by default. Generally in range [1.2, 2.0] */
  k1?: number;
}

/** If returns positive, the sorting results in secondEl coming before firstEl, else, firstEl comes before secondEL  */
export type BMSorter = (firstEl: BMDocument, secondEl: BMDocument) => number;

/** Implementation of Okapi BM25 algorithm.
 *  @param documents: Collection of documents.
 *  @param keywords: query terms.
 *  @param constants: Contains free parameters k1 and b. b=0.75 and k1=1.2 by default.
 *  @param sort: A function that allows you to sort queries by a given rule. If not provided, returns results corresponding to the original order.
 * If this option is provided, the return type will not be an array of scores but an array of documents with their scores.
 */
export function BM25(
  documents: string[],
  keywords: string[],
  constants?: BMConstants,
  sorter?: BMSorter
): number[] | BMDocument[] {
  const b = constants && constants.b ? constants.b : 0.75;
  const k1 = constants && constants.k1 ? constants.k1 : 1.2;
  const documentLengths = documents.map((document: string) => getWordCount(document));
  const averageDocumentLength = documentLengths.reduce((a, b) => a + b, 0) / documents.length;
  const scores = documents.map((document: string, index: number) => {
    const score = keywords
      .map((keyword: string) => {
        const inverseDocumentFrequency = getIDF(keyword, documents);
        const termFrequency = getTermFrequency(keyword, document);
        const documentLength = documentLengths[index];
        return (
          (inverseDocumentFrequency * (termFrequency * (k1 + 1))) /
          (termFrequency + k1 * (1 - b + (b * documentLength) / averageDocumentLength))
        );
      })
      .reduce((a: number, b: number) => a + b, 0);
    if (sorter) {
      return { score, document } as BMDocument;
    }
    return score;
  });
  // sort the results
  if (sorter) {
    return (scores as BMDocument[]).sort(sorter);
  }
  return scores as number[];
}
