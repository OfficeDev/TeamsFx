// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { OptionItem } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../../common/localizeUtils";

export const githubOption: OptionItem = {
  id: "github",
  label: "GitHub",
  detail: "",
};

export const azdoOption: OptionItem = {
  id: "azdo",
  label: "Azure DevOps",
  detail: "",
};

export const jenkinsOption: OptionItem = {
  id: "jenkins",
  label: "Jenkins",
  detail: "",
};

export const ciOption: OptionItem = {
  id: "ci",
  label: "CI",
  detail: getLocalizedString("plugins.cicd.ciOption.detail"),
};

export const cdOption: OptionItem = {
  id: "cd",
  label: "CD",
  detail: getLocalizedString("plugins.cicd.cdOption.detail"),
};

export const provisionOption: OptionItem = {
  id: "provision",
  label: "Provision",
  detail: getLocalizedString("plugins.cicd.provisionOption.detail"),
};

export const publishOption: OptionItem = {
  id: "publish",
  label: "Publish to Teams",
  detail: getLocalizedString("plugins.cicd.publishOption.detail"),
};

const templateIdLabelMap = new Map<string, string>([
  [ciOption.id, ciOption.label],
  [cdOption.id, cdOption.label],
  [provisionOption.id, provisionOption.label],
  [publishOption.id, publishOption.label],
]);

const providerIdLabelMap = new Map<string, string>([
  [githubOption.id, githubOption.label],
  [azdoOption.id, azdoOption.label],
  [jenkinsOption.id, jenkinsOption.label],
]);

export function templateIdToLabel(templateId: string): string {
  return templateIdLabelMap.get(templateId) ?? templateId;
}

export function providerIdToLabel(providerId: string): string {
  return providerIdLabelMap.get(providerId) ?? providerId;
}

export enum questionNames {
  Provider = "provider",
  Template = "template",
  Environment = "target-env",
}
