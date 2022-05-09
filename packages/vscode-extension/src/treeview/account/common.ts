// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";

export enum AccountItemStatus {
  SignedOut,
  SigningIn,
  SignedIn,
}

export const loadingIcon = new vscode.ThemeIcon("loading~spin");
