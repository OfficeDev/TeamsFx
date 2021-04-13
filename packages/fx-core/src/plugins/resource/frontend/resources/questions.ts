// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { NodeType, QTreeNode } from "fx-api";
import { TabLanguage } from "./templateInfo";

export class QuestionKey {
    static readonly TabLanguage = "TabLanguage";
    static readonly TabScope = "TabScope";
}

export class TabScope {
    static readonly PersonalTab = "personal tab";
    static readonly GroupTab = "group tab";
}

export const tabLanguageQuestion = new QTreeNode({
    name: QuestionKey.TabLanguage,
    description: "Select language for tab frontend project",
    type: NodeType.singleSelect,
    option: [TabLanguage.JavaScript, TabLanguage.TypeScript],
});

export const tabScopeQuestion = new QTreeNode({
    name: QuestionKey.TabScope,
    description: "Select tab scope",
    type: NodeType.singleSelect,
    option: [TabScope.PersonalTab, TabScope.GroupTab],
    default: TabScope.PersonalTab,
});
