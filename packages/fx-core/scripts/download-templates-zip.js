// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
const axios = require("axios");
const semver = require("semver");
const process = require("process");
const fs = require("fs-extra");
const path = require("path");
const config = require("../src/common/templates-config.json");

const scenrioSupportedLanguages = new Map([
  [
    ["function-base", "default", "function"],
    ["js", "ts"],
  ],
  [
    ["function-triggers", "HTTPTrigger", "function"],
    ["js", "ts"],
  ],
  [
    ["tab", "default", "frontend"],
    ["js", "ts", "csharp"],
  ],
  [
    ["bot", "default", "bot"],
    ["js", "ts", "csharp"],
  ],
  [["blazor-base", "default", "dotnet"], ["csharp"]],
]);

let stepId = 0;

async function step(desc, fn) {
  const id = ++stepId;
  try {
    console.log(`step ${id} start: ${desc}`);
    const ret = await fn();
    return ret;
  } catch (e) {
    console.log(e.toString());
    console.log(`step ${id} failed`);
    process.exit(-1);
  }
}

async function downloadTemplates(version) {
  const tag = config.tagPrefix + version;
  console.log(`Start to download templates with tag: ${tag}`);

  for (let scenrio of Array.from(scenrioSupportedLanguages.keys())) {
    for (let lang of scenrioSupportedLanguages.get(scenrio)) {
      const fileName = `${scenrio[0]}.${lang}.${scenrio[1]}.zip`;
      step(`Download ${config.templateDownloadBaseURL}/${tag}/${fileName}`, async () => {
        const res = await axios.get(`${config.templateDownloadBaseURL}/${tag}/${fileName}`, {
          responseType: "arraybuffer",
        });
        const folder = path.join(__dirname, "..", "templates", "plugins", "resource", scenrio[2]);
        await fs.ensureDir(folder);
        await fs.writeFile(path.join(folder, `${fileName}`), res.data);
      });
    }
  }
}

function selectVersion(tagList) {
  const versionList = tagList
    .filter((tag) => tag.startsWith(config.tagPrefix))
    .map((tag) => tag.replace(config.tagPrefix, ""));
  return semver.maxSatisfying(versionList, config.version);
}

function selectVersionFromShellArgument() {
  const tagList = process.argv.slice(2);
  return selectVersion(tagList);
}

async function selectVersionFromRemoteTagList() {
  const rawTagList = await step(`Download tag list from ${config.tagListURL}`, async () => {
    res = await axios.get(config.tagListURL);
    return res.data;
  });
  const tagList = rawTagList.toString().replace(/\r/g, "").split("\n");
  return selectVersion(tagList);
}

async function main() {
  const selectedVersion =
    selectVersionFromShellArgument() || (await selectVersionFromRemoteTagList());
  if (!selectVersion) {
    console.error(`Failed to find a tag for the version, ${config.version}`);
    process.exit(-1);
  }
  await downloadTemplates(selectedVersion);
}

main();
