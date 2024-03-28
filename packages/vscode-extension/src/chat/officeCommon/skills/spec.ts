// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { deepClone } from "../Utils";

// Licensed under the MIT license.
export class Spec {
  public userInput: string;
  public taskSummary: string;
  public sections: string[];
  public inspires: string[];
  public resources: string[];
  public appendix: {
    host: string;
    codeSnippet: string;
    codeExplanation: string;
    codeTaskBreakdown: string[];
    isCustomFunction: boolean;
    telemetryData: {
      properties: { [key: string]: string };
      measurements: { [key: string]: number };
    };
    tempAppLocation: string;
  };

  constructor(userInput: string) {
    this.userInput = userInput;
    this.taskSummary = "";
    this.sections = [];
    this.inspires = [];
    this.resources = [];
    this.appendix = {
      host: "",
      codeSnippet: "",
      codeExplanation: "",
      codeTaskBreakdown: [],
      isCustomFunction: false,
      telemetryData: {
        properties: {},
        measurements: {},
      },
      tempAppLocation: "",
    };
  }

  public clone(other: Spec): Spec {
    this.userInput = other.userInput;
    this.taskSummary = other.taskSummary;
    this.sections = other.sections;
    this.inspires = other.inspires;
    this.resources = other.resources;
    this.appendix = deepClone(other.appendix);
    return this;
  }
}
