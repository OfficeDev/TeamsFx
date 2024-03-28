// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import axios from "axios";
import { sendRequestWithTimeout } from "@microsoft/teamsfx-core/build/component/generator/utils";

export async function fetchRawFileContent(url: string): Promise<string> {
  try {
    const fileResponse = await sendRequestWithTimeout(
      async () => {
        return await axios.get(url);
      },
      1000,
      3
    );

    if (fileResponse && fileResponse.data) {
      return fileResponse.data as string;
    }

    return "";
  } catch (e) {
    throw new Error(`Cannot fetch ${url}.`);
  }
}

export function compressCode(code: string): string {
  // Remove comments
  let result = code.replace(/\/\*[\s\S]*?\*\/|\/\/.*/g, "");
  // Remove unnecessary spaces
  result = result.replace(/ +/g, " ");
  return result;
}

// For test purpose
export async function sleep(second: number): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, second * 1000);
  });
}

export async function sleepRandom(minSecond: number, maxSecond: number): Promise<void> {
  const second = Math.floor(Math.random() * (maxSecond - minSecond + 1)) + minSecond;
  return sleep(second);
}

// For test purpose
export async function writeLogToFile(log: string): Promise<void> {
  const filePath = "C:\\temp\\codeGenLog.txt";
  const fs = require("fs");
  await fs.appendFileSync(filePath, log);
}

export function deepClone<T>(obj: T): T {
  return JSON.parse(JSON.stringify(obj));
}
