// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { OfficeAddinChatCommand } from "../../consts";
import { Explainer } from "./codeExplainer";
import { CodeGenerator } from "./codeGenerator";
import { CodeMerger } from "./codeMerger";
import { ISkill } from "./iSkill"; // Replace this import statement
import { Printer } from "./printer";
import { projectCreator } from "./projectCreator";
import { SkillSet } from "./skillset";

export class SkillsManager {
  private static instance: SkillsManager;
  private projectCreator: ISkill;
  private codeGenerator: ISkill;
  private codeExplainer: ISkill;
  private codeMerger: ISkill;
  private printer: ISkill;

  private constructor() {
    // Private constructor to prevent direct instantiation
    this.codeGenerator = new CodeGenerator();
    this.printer = new Printer();
    this.codeExplainer = new Explainer();
    this.projectCreator = new projectCreator();
    this.codeMerger = new CodeMerger();
  }

  public static getInstance(): SkillsManager {
    if (!SkillsManager.instance) {
      SkillsManager.instance = new SkillsManager();
    }
    return SkillsManager.instance;
  }

  public getCapableSkills(capability: OfficeAddinChatCommand): ISkill[] {
    const capableSkills: ISkill[] = [];
    switch (capability) {
      case OfficeAddinChatCommand.GenerateCode:
        capableSkills.push(new SkillSet([this.codeGenerator], 2));
        capableSkills.push(this.codeExplainer);
        capableSkills.push(this.printer);
        break;
      case OfficeAddinChatCommand.Create:
        capableSkills.push(new SkillSet([this.codeGenerator], 2));
        capableSkills.push(this.codeExplainer);
        capableSkills.push(this.printer);
        capableSkills.push(this.projectCreator);
        capableSkills.push(this.codeMerger);
        break;
      default:
        break;
    }

    return capableSkills;
  }
}
