// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

"use strict";

import * as TypeMoq from "typemoq";
import * as vscode from "vscode";
import * as vscodeMocks from "./vsc";
import { Disposable } from "./vsc/extHostedTypes";
import { vscMockTelemetryReporter } from "./vsc/telemetryReporter";
const Module = require("module");

type VSCode = typeof vscode;

const mockedVSCode: Partial<VSCode> = {};
export const mockedVSCodeNamespaces: { [P in keyof VSCode]?: TypeMoq.IMock<VSCode[P]> } = {};
const originalLoad = Module._load;

export function initialize() {
  generateMock("languages");
  generateMock("env");
  generateMock("debug");
  generateMock("scm");
  generateNotebookMocks();

  // Use mock clipboard fo testing purposes.
  const clipboard = new MockClipboard();
  mockedVSCodeNamespaces.env?.setup((e) => e.clipboard).returns(() => clipboard);
  mockedVSCodeNamespaces.env?.setup((e) => e.appName).returns(() => "Insider");

  // When upgrading to npm 9-10, this might have to change, as we could have explicit imports (named imports).
  Module._load = function (request: any, _parent: any) {
    if (request === "vscode") {
      return mockedVSCode;
    }
    if (request === "@vscode/extension-telemetry") {
      return { default: vscMockTelemetryReporter as any };
    }
    // less files need to be in import statements to be converted to css
    // But we don't want to try to load them in the mock vscode
    if (/\.less$/.test(request)) {
      return;
    }
    // eslint-disable-next-line prefer-rest-params
    return originalLoad.apply(this, arguments);
  };
}

mockedVSCode.MarkdownString = vscodeMocks.MarkdownString;
mockedVSCode.Hover = vscodeMocks.Hover;
mockedVSCode.Disposable = vscodeMocks.Disposable as any;
mockedVSCode.ExtensionKind = vscodeMocks.ExtensionKind;
mockedVSCode.CodeAction = vscodeMocks.CodeAction;
mockedVSCode.EventEmitter = vscodeMocks.EventEmitter;
mockedVSCode.CancellationTokenSource = vscodeMocks.CancellationTokenSource;
mockedVSCode.CompletionItemKind = vscodeMocks.CompletionItemKind;
mockedVSCode.SymbolKind = vscodeMocks.SymbolKind;
mockedVSCode.IndentAction = vscodeMocks.IndentAction;
mockedVSCode.Uri = vscodeMocks.vscUri.URI as any;
mockedVSCode.Range = vscodeMocks.vscMockExtHostedTypes.Range;
mockedVSCode.Position = vscodeMocks.vscMockExtHostedTypes.Position;
mockedVSCode.Selection = vscodeMocks.vscMockExtHostedTypes.Selection;
mockedVSCode.Location = vscodeMocks.vscMockExtHostedTypes.Location;
mockedVSCode.SymbolInformation = vscodeMocks.vscMockExtHostedTypes.SymbolInformation;
mockedVSCode.CallHierarchyItem = vscodeMocks.vscMockExtHostedTypes.CallHierarchyItem;
mockedVSCode.CompletionItem = vscodeMocks.vscMockExtHostedTypes.CompletionItem;
mockedVSCode.CompletionItemKind = vscodeMocks.vscMockExtHostedTypes.CompletionItemKind;
mockedVSCode.CodeLens = vscodeMocks.vscMockExtHostedTypes.CodeLens;
mockedVSCode.Diagnostic = vscodeMocks.vscMockExtHostedTypes.Diagnostic;
mockedVSCode.DiagnosticSeverity = vscodeMocks.vscMockExtHostedTypes.DiagnosticSeverity;
mockedVSCode.SnippetString = vscodeMocks.vscMockExtHostedTypes.SnippetString;
mockedVSCode.ConfigurationTarget = vscodeMocks.vscMockExtHostedTypes.ConfigurationTarget;
mockedVSCode.StatusBarAlignment = vscodeMocks.vscMockExtHostedTypes.StatusBarAlignment;
mockedVSCode.SignatureHelp = vscodeMocks.vscMockExtHostedTypes.SignatureHelp;
mockedVSCode.DocumentLink = vscodeMocks.vscMockExtHostedTypes.DocumentLink;
mockedVSCode.TextEdit = vscodeMocks.vscMockExtHostedTypes.TextEdit;
mockedVSCode.WorkspaceEdit = vscodeMocks.vscMockExtHostedTypes.WorkspaceEdit;
mockedVSCode.RelativePattern = vscodeMocks.vscMockExtHostedTypes.RelativePattern;
mockedVSCode.ProgressLocation = vscodeMocks.vscMockExtHostedTypes.ProgressLocation;
mockedVSCode.ViewColumn = vscodeMocks.vscMockExtHostedTypes.ViewColumn;
mockedVSCode.TextEditorRevealType = vscodeMocks.vscMockExtHostedTypes.TextEditorRevealType;
mockedVSCode.TreeItem = vscodeMocks.vscMockExtHostedTypes.TreeItem;
mockedVSCode.TreeItemCollapsibleState = vscodeMocks.vscMockExtHostedTypes.TreeItemCollapsibleState;
mockedVSCode.CodeActionKind = vscodeMocks.CodeActionKind;
mockedVSCode.CompletionItemKind = vscodeMocks.CompletionItemKind;
mockedVSCode.CompletionTriggerKind = vscodeMocks.CompletionTriggerKind;
mockedVSCode.DebugAdapterExecutable = vscodeMocks.DebugAdapterExecutable;
mockedVSCode.DebugAdapterServer = vscodeMocks.DebugAdapterServer;
mockedVSCode.QuickInputButtons = vscodeMocks.vscMockExtHostedTypes.QuickInputButtons;
mockedVSCode.FileType = vscodeMocks.FileType;
mockedVSCode.UIKind = vscodeMocks.UIKind;
mockedVSCode.FileSystemError = vscodeMocks.vscMockExtHostedTypes.FileSystemError;
mockedVSCode.QuickPickItemKind = vscodeMocks.QuickPickItemKind;
mockedVSCode.ThemeIcon = vscodeMocks.vscMockExtHostedTypes.ThemeIcon;
mockedVSCode.ThemeColor = vscodeMocks.vscMockExtHostedTypes.ThemeColor;
mockedVSCode.Task = vscodeMocks.vscMockExtHostedTypes.Task;
(mockedVSCode as any).NotebookCellKind = vscodeMocks.vscMockExtHostedTypes.NotebookCellKind;
(mockedVSCode as any).CellOutputKind = vscodeMocks.vscMockExtHostedTypes.CellOutputKind;
(mockedVSCode as any).NotebookCellRunState = vscodeMocks.vscMockExtHostedTypes.NotebookCellRunState;
(mockedVSCode as any).TypeHierarchyItem = vscodeMocks.vscMockExtHostedTypes.TypeHierarchyItem;
(mockedVSCode as any).ProtocolTypeHierarchyItem =
  vscodeMocks.vscMockExtHostedTypes.ProtocolTypeHierarchyItem;
(mockedVSCode as any).CancellationError = vscodeMocks.vscMockExtHostedTypes.CancellationError;
(mockedVSCode as any).LSPCancellationError = vscodeMocks.vscMockExtHostedTypes.LSPCancellationError;

// Setup window APIs
(mockedVSCode as any).window = {
  activeTextEditor: undefined,
  activeTerminal: undefined,
  terminals: [],
  showInformationMessage: () => {},
  showErrorMessage: () => {
    return Promise.resolve("success");
  },
  showWarningMessage: () => {},
  createOutputChannel: () => {},
  registerTreeDataProvider: () => {},
  withProgress: async (options: any, task: (progress: any, token: any) => Promise<any>) => {
    return await task({ report: () => {} }, new vscodeMocks.CancellationToken());
  },
  createQuickPick: () => {},
};
(mockedVSCode as any).workspace = {
  workspaceFolders: undefined,
  openTextDocument: () => {},
};

// Setup extensions APIs
mockedVSCode.extensions = {
  getExtension: () => {
    return undefined;
  },
  onDidChange: () => {
    return new Disposable(() => {
      return;
    });
  },
  all: [],
};

// Setup commands APIs
mockedVSCode.commands = {
  executeCommand: () => {
    const res: any = "success";
    return Promise.resolve(res);
  },
  registerCommand: (command: string, callback: (...args: any[]) => any, thisArg?: any) => {
    return new Disposable(() => {
      return;
    });
  },
  getCommands: (filter: boolean | undefined) => {
    return Promise.resolve([]);
  },
  registerTextEditorCommand: () => {
    return new Disposable(() => {
      return;
    });
  },
};

function generateNotebookMocks() {
  const mockedObj = TypeMoq.Mock.ofType<Record<string, unknown>>();
  (mockedVSCode as any).notebook = mockedObj.object;
  (mockedVSCodeNamespaces as any).notebook = mockedObj as any;
}

function generateMock<K extends keyof VSCode>(name: K): void {
  const mockedObj = TypeMoq.Mock.ofType<VSCode[K]>();
  (mockedVSCode as any)[name] = mockedObj.object;
  mockedVSCodeNamespaces[name] = mockedObj as any;
}

class MockClipboard {
  private text = "";
  public readText(): Promise<string> {
    return Promise.resolve(this.text);
  }
  public async writeText(value: string): Promise<void> {
    this.text = value;
  }
}
