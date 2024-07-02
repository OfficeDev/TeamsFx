// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IQTreeNode } from "@microsoft/teamsfx-api";
import {
  createProjectCliHelpNode,
  createProjectQuestionNode,
  createSampleProjectQuestionNode,
} from "./create";
import {
  addPluginQuestionNode,
  addWebPartQuestionNode,
  apiSpecApiKeyQuestion,
  copilotPluginAddAPIQuestionNode,
  createNewEnvQuestionNode,
  deployAadManifestQuestionNode,
  grantPermissionQuestionNode,
  listCollaboratorQuestionNode,
  oauthQuestion,
  previewWithTeamsAppManifestQuestionNode,
  selectTeamsAppManifestQuestionNode,
  uninstallQuestionNode,
  validateTeamsAppQuestionNode,
} from "./other";
export * from "./constants";
export * from "./create";
export * from "./inputs";
export * from "./options";

export class QuestionNodes {
  createProject(): IQTreeNode {
    return createProjectQuestionNode();
  }
  createSampleProject(): IQTreeNode {
    return createSampleProjectQuestionNode();
  }
  createProjectCliHelp(): IQTreeNode {
    return createProjectCliHelpNode();
  }
  addWebpart(): IQTreeNode {
    return addWebPartQuestionNode();
  }
  selectTeamsAppManifest(): IQTreeNode {
    return selectTeamsAppManifestQuestionNode();
  }
  validateTeamsApp(): IQTreeNode {
    return validateTeamsAppQuestionNode();
  }
  previewWithTeamsAppManifest(): IQTreeNode {
    return previewWithTeamsAppManifestQuestionNode();
  }
  listCollaborator(): IQTreeNode {
    return listCollaboratorQuestionNode();
  }
  grantPermission(): IQTreeNode {
    return grantPermissionQuestionNode();
  }
  deployAadManifest(): IQTreeNode {
    return deployAadManifestQuestionNode();
  }
  createNewEnv(): IQTreeNode {
    return createNewEnvQuestionNode();
  }
  copilotPluginAddAPI(): IQTreeNode {
    return copilotPluginAddAPIQuestionNode();
  }
  apiKey(): IQTreeNode {
    return apiSpecApiKeyQuestion();
  }
  oauth(): IQTreeNode {
    return oauthQuestion();
  }
  addPlugin(): IQTreeNode {
    return addPluginQuestionNode();
  }
  uninstall(): IQTreeNode {
    return uninstallQuestionNode();
  }
}

export const questionNodes = new QuestionNodes();
