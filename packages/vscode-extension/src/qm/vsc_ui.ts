// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { remove } from "lodash";
import {
  Disposable,
  InputBox,
  QuickInputButtons,
  QuickPick,
  QuickPickItem,
  Uri,
  window,
  env,
  ProgressLocation,
  ExtensionContext,
  commands,
  extensions,
  QuickInputButton,
  ThemeIcon,
  QuickPickItemKind,
} from "vscode";
import {
  UserCancelError,
  FxError,
  InputResult,
  SingleSelectResult,
  MultiSelectResult,
  InputTextResult,
  SelectFileResult,
  SelectFilesResult,
  SelectFolderResult,
  OptionItem,
  Result,
  SelectFileConfig,
  SelectFilesConfig,
  SelectFolderConfig,
  SingleSelectConfig,
  MultiSelectConfig,
  InputTextConfig,
  ExecuteFuncConfig,
  RunnableTask,
  UIConfig,
  err,
  assembleError,
  ok,
  TaskConfig,
  UserInteraction,
  Colors,
  IProgressHandler,
  SystemError,
  StaticOptions,
} from "@microsoft/teamsfx-api";
import { ExtensionErrors, ExtensionSource } from "../error";
import { sleep } from "../utils/commonUtils";
import { ProgressHandler } from "../progressHandler";
import * as exp from "../exp";
import { TreatmentVariables } from "../exp/treatmentVariables";
import * as packageJson from "../../package.json";
import * as ExtTelemetry from "../telemetry/telemetry";
import { getDefaultString, localize } from "../utils/localizeUtils";

export interface FxQuickPickItem extends QuickPickItem {
  id: string;
  data?: unknown;
}

function getOptionItem(item: FxQuickPickItem): OptionItem {
  return {
    id: item.id,
    label: item.label,
    description: item.description,
    detail: item.detail,
    data: item.data,
  };
}

function convertToFxQuickPickItems(options: StaticOptions): FxQuickPickItem[] {
  if (options && options.length > 0 && typeof options[0] === "string") {
    return (options as string[]).map((option: string) => {
      return { id: option, label: option };
    });
  } else {
    const result: FxQuickPickItem[] = [];
    const candidates = [...options];
    while (candidates.length > 0) {
      const groupName = (candidates[0] as OptionItem).groupName;
      const group = remove(
        candidates as OptionItem[],
        (option: OptionItem) => option.groupName === groupName
      );
      if (groupName) {
        result.push({
          id: groupName,
          label: groupName,
          kind: QuickPickItemKind.Separator,
        });
      }
      result.push(
        ...group.map((option) => {
          return {
            id: option.id,
            label: option.label,
            description: option.description,
            detail: option.detail,
            data: option.data,
            buttons: option.buttons?.map((button) => {
              return { iconPath: new ThemeIcon(button.iconPath), tooltip: button.tooltip };
            }),
          };
        })
      );
    }
    return result;
  }
}

function toIdSet(items: ({ id: string } | string)[]): Set<string> {
  const set = new Set<string>();
  for (const i of items) {
    if (typeof i === "string") set.add(i);
    else set.add(i.id);
  }
  return set;
}

export function cloneSet(set: Set<string>): Set<string> {
  const res = new Set<string>();
  for (const e of set) res.add(e);
  return res;
}

function isSame(set1: Set<string>, set2: Set<string>): boolean {
  for (const i of set1) {
    if (!set2.has(i)) return false;
  }
  for (const i of set2) {
    if (!set1.has(i)) return false;
  }
  return true;
}

export class VsCodeUI implements UserInteraction {
  context: ExtensionContext;
  constructor(context: ExtensionContext) {
    this.context = context;
  }

  async selectOption(option: SingleSelectConfig): Promise<Result<SingleSelectResult, FxError>> {
    if (option.options.length === 0) {
      return err(
        new SystemError(
          ExtensionSource,
          ExtensionErrors.EmptySelectOption,
          getDefaultString("teamstoolkit.qm.emptySelection"),
          localize("teamstoolkit.qm.emptySelection")
        )
      );
    }
    const disposables: Disposable[] = [];
    try {
      const quickPick = window.createQuickPick<FxQuickPickItem>();
      quickPick.title = option.title;
      const buttons: QuickInputButton[] = option.buttons
        ? option.buttons.map((button) => {
            return {
              iconPath: new ThemeIcon(button.icon),
              tooltip: button.tooltip,
            } as QuickInputButton;
          })
        : [];
      if (option.step && option.step > 1) {
        quickPick.buttons = [QuickInputButtons.Back, ...buttons];
      } else {
        quickPick.buttons = buttons;
      }
      quickPick.placeholder = option.placeholder;
      quickPick.ignoreFocusOut = true;
      quickPick.matchOnDescription = true;
      quickPick.matchOnDetail = true;
      quickPick.canSelectMany = false;
      return await new Promise<Result<SingleSelectResult, FxError>>(
        async (resolve): Promise<void> => {
          // set items
          const options = option.options;
          quickPick.items = convertToFxQuickPickItems(option.options);
          // set default
          if (option.default) {
            // let defaultOption: string | OptionItem | undefined;
            if (options && options.length > 0 && typeof options[0] === "string") {
              const defaultOption = (options as string[]).find((o) => o == option.default);
              if (defaultOption) {
                const newItems = (options as string[]).filter((o) => o != option.default);
                newItems.unshift(defaultOption);
                quickPick.items = convertToFxQuickPickItems(newItems);
              }
            } else {
              const defaultOption = (options as OptionItem[]).find((o) => o.id == option.default);
              if (defaultOption) {
                const newItems = (options as OptionItem[]).filter((o) => o.id != option.default);
                newItems.unshift(defaultOption);
                quickPick.items = convertToFxQuickPickItems(newItems);
              }
            }
          }

          const onDidAccept = async () => {
            const selectedItems = quickPick.selectedItems;
            if (selectedItems && selectedItems.length > 0) {
              const item = selectedItems[0];
              let result: string | OptionItem;
              if (
                typeof option.options[0] === "string" ||
                option.returnObject === undefined ||
                option.returnObject === false
              )
                result = item.id;
              else result = getOptionItem(item);
              resolve(ok({ type: "success", result: result }));
            }
          };

          disposables.push(
            quickPick.onDidAccept(onDidAccept),
            quickPick.onDidHide(() => {
              resolve(err(UserCancelError));
            }),
            quickPick.onDidTriggerButton((button) => {
              if (button === QuickInputButtons.Back) resolve(ok({ type: "back" }));
              else if (option.buttons && buttons.indexOf(button) !== -1) {
                const curButton = option.buttons?.find((btn) => {
                  return (
                    btn.icon === (button.iconPath as ThemeIcon).id && btn.tooltip === button.tooltip
                  );
                });
                if (curButton) {
                  commands.executeCommand(curButton.command);
                }
              } else {
                quickPick.selectedItems = quickPick.activeItems;
                onDidAccept();
              }
            }),
            quickPick.onDidTriggerItemButton((event) => {
              const itemOptions: StaticOptions = option.options;
              if (itemOptions.length > 0 && typeof itemOptions[0] === "string") {
                return;
              }
              const triggerItem: OptionItem | undefined = (itemOptions as OptionItem[]).find(
                (singleOption: string | OptionItem) => {
                  if (typeof singleOption !== "string") {
                    return (singleOption as OptionItem).id === event.item.id;
                  }
                }
              );
              if (triggerItem) {
                const triggerButton = triggerItem.buttons?.find((button) => {
                  return button.iconPath === (event.button.iconPath as ThemeIcon).id;
                });
                if (triggerButton) {
                  commands.executeCommand(triggerButton.command, event.item);
                }
              }
            })
          );

          disposables.push(quickPick);
          quickPick.show();
        }
      );
    } finally {
      disposables.forEach((d) => {
        d.dispose();
      });
    }
  }

  async selectOptions(option: MultiSelectConfig): Promise<Result<MultiSelectResult, FxError>> {
    if (option.options.length === 0) {
      return err(
        new SystemError(
          ExtensionSource,
          ExtensionErrors.EmptySelectOption,
          getDefaultString("teamstoolkit.qm.emptySelection"),
          localize("teamstoolkit.qm.emptySelection")
        )
      );
    }
    const disposables: Disposable[] = [];
    try {
      const quickPick: QuickPick<FxQuickPickItem> = window.createQuickPick<FxQuickPickItem>();
      quickPick.title = option.title;
      if (option.step && option.step > 1) {
        quickPick.buttons = [QuickInputButtons.Back];
      }
      quickPick.placeholder = option.placeholder
        ? option.placeholder + localize("teamstoolkit.qm.multiSelectKeyboard")
        : localize("teamstoolkit.qm.multiSelectKeyboard");
      quickPick.ignoreFocusOut = true;
      quickPick.matchOnDescription = true;
      quickPick.matchOnDetail = true;
      quickPick.canSelectMany = true;
      const preIds: Set<string> = new Set<string>();
      return await new Promise<Result<MultiSelectResult, FxError>>(
        async (resolve): Promise<void> => {
          // set items
          quickPick.items = convertToFxQuickPickItems(option.options);
          const optionMap = new Map<string, FxQuickPickItem>();
          for (const item of quickPick.items) {
            optionMap.set(item.id, item);
          }

          // set default values
          if (option.default) {
            const ids = option.default as string[];
            const selectedItems: FxQuickPickItem[] = [];
            preIds.clear();
            for (const id of ids) {
              const item = optionMap.get(id);
              if (item) {
                selectedItems.push(item);
                preIds.add(id);
              }
            }
            quickPick.selectedItems = selectedItems;
          }

          const onDidAccept = async () => {
            const strArray = Array.from(quickPick.selectedItems.map((i) => i.id));
            if (option.validation) {
              const validateRes = await option.validation(strArray);
              if (validateRes) {
                return;
              }
            }
            let result: OptionItem[] | string[] = strArray;
            if (
              typeof option.options[0] === "string" ||
              option.returnObject === undefined ||
              option.returnObject === false
            )
              result = strArray;
            else result = quickPick.selectedItems.map((i) => getOptionItem(i));
            resolve(ok({ type: "success", result: result }));
          };

          disposables.push(
            quickPick.onDidAccept(onDidAccept),
            quickPick.onDidHide(() => {
              resolve(err(UserCancelError));
            }),
            quickPick.onDidTriggerButton((button) => {
              if (button === QuickInputButtons.Back) resolve(ok({ type: "back" }));
              else onDidAccept();
            })
          );

          if (option.onDidChangeSelection) {
            const changeHandler = async function (items: readonly FxQuickPickItem[]): Promise<any> {
              let currentIds = new Set<string>();
              for (const item of items) {
                currentIds.add(item.id);
              }
              if (option.onDidChangeSelection) {
                const currentClone = cloneSet(currentIds);
                currentIds = await option.onDidChangeSelection(currentIds, preIds);
                const selectedItems: FxQuickPickItem[] = [];
                preIds.clear();
                for (const id of currentIds) {
                  const item = optionMap.get(id);
                  if (item) {
                    selectedItems.push(item);
                    preIds.add(id);
                  }
                }
                if (!isSame(currentClone, currentIds)) {
                  quickPick.selectedItems = selectedItems;
                }
              }
            };
            disposables.push(quickPick.onDidChangeSelection(changeHandler));
          }

          disposables.push(quickPick);
          quickPick.show();
        }
      );
    } finally {
      disposables.forEach((d) => {
        d.dispose();
      });
    }
  }

  async inputText(option: InputTextConfig): Promise<Result<InputTextResult, FxError>> {
    const disposables: Disposable[] = [];
    try {
      const inputBox: InputBox = window.createInputBox();
      inputBox.title = option.title;
      if (option.step && option.step > 1) {
        inputBox.buttons = [QuickInputButtons.Back];
      }
      inputBox.placeholder = option.placeholder;
      inputBox.value = option.default || "";
      inputBox.ignoreFocusOut = true;
      inputBox.password = option.password === true;
      inputBox.prompt = option.prompt;
      return await new Promise<Result<InputTextResult, FxError>>((resolve): void => {
        const onDidAccept = async () => {
          const validationRes = option.validation
            ? await option.validation(inputBox.value)
            : undefined;
          if (!validationRes) {
            resolve(ok({ type: "success", result: inputBox.value }));
          } else {
            inputBox.validationMessage = validationRes;
          }
        };
        disposables.push(
          inputBox.onDidChangeValue(async (text) => {
            if (option.validation) {
              const validationRes = option.validation ? await option.validation(text) : undefined;
              if (!!validationRes) {
                inputBox.validationMessage = validationRes;
              } else {
                inputBox.validationMessage = undefined;
              }
            }
          }),
          inputBox.onDidAccept(onDidAccept),
          inputBox.onDidHide(() => {
            resolve(err(UserCancelError));
          }),
          inputBox.onDidTriggerButton((button) => {
            if (button === QuickInputButtons.Back) resolve(ok({ type: "back" }));
            else onDidAccept();
          })
        );
        disposables.push(inputBox);
        inputBox.show();
      });
    } finally {
      disposables.forEach((d) => {
        d.dispose();
      });
    }
  }

  async selectFolder(config: SelectFolderConfig): Promise<Result<SelectFolderResult, FxError>> {
    return this.selectFileInQuickPick(config, "folder", config.default);
  }

  async selectFile(config: SelectFileConfig): Promise<Result<SelectFileResult, FxError>> {
    return this.selectFileInQuickPick(config, "file", config.default);
  }

  async selectFiles(config: SelectFilesConfig): Promise<Result<SelectFilesResult, FxError>> {
    return this.selectFileInQuickPick(
      config,
      "files",
      config.default ? config.default.join(";") : undefined
    );
  }

  async selectFileInQuickPick(
    config: SelectFileConfig,
    type: "file" | "files" | "folder",
    defaultValue?: string
  ): Promise<Result<SelectFileResult, FxError>>;
  async selectFileInQuickPick(
    config: SelectFilesConfig,
    type: "file" | "files" | "folder",
    defaultValue?: string
  ): Promise<Result<SelectFilesResult, FxError>>;
  async selectFileInQuickPick(
    config: SelectFolderConfig,
    type: "file" | "files" | "folder",
    defaultValue?: string
  ): Promise<Result<SelectFolderResult, FxError>>;
  async selectFileInQuickPick(
    config: UIConfig<any>,
    type: "file" | "files" | "folder",
    defaultValue?: string
  ): Promise<Result<InputResult<string[] | string>, FxError>> {
    /// TODO: use generic constraints.
    const disposables: Disposable[] = [];
    try {
      const quickPick: QuickPick<QuickPickItem> = window.createQuickPick();
      quickPick.title = config.title;
      if (config.step && config.step > 1) {
        quickPick.buttons = [QuickInputButtons.Back];
      }
      quickPick.ignoreFocusOut = true;
      quickPick.placeholder = config.placeholder;
      quickPick.matchOnDescription = false;
      quickPick.matchOnDetail = false;
      quickPick.canSelectMany = false;
      let fileSelectorIsOpen = false;
      return await new Promise(async (resolve) => {
        const onDidAccept = () => {
          const result = quickPick.items[0].detail;
          if (result && result.length > 0) {
            if (type === "files") {
              resolve(ok({ type: "success", result: result.split(";") }));
            } else {
              resolve(ok({ type: "success", result: result }));
            }
          }
        };

        disposables.push(
          quickPick.onDidHide(() => {
            if (fileSelectorIsOpen === false) resolve(err(UserCancelError));
          }),
          quickPick.onDidTriggerButton((button) => {
            if (button === QuickInputButtons.Back) resolve(ok({ type: "back" }));
            else onDidAccept();
          })
        );

        /// set items
        quickPick.items = [
          {
            label:
              config.prompt ||
              localize(
                type === "folder" ? "teamstoolkit.qm.selectFolder" : "teamstoolkit.qm.selectFile"
              ),
            detail: defaultValue,
          },
        ];
        const showFileSelectDialog = async function (defaultUrl?: string) {
          fileSelectorIsOpen = true;
          const uriList: Uri[] | undefined = await window.showOpenDialog({
            defaultUri: defaultUrl ? Uri.file(defaultUrl) : undefined,
            canSelectFiles: type === "file" || type === "files",
            canSelectFolders: type === "folder",
            canSelectMany: type === "files",
            title: config.title,
          });
          fileSelectorIsOpen = false;
          if (uriList && uriList.length > 0) {
            if (type === "files") {
              const results = uriList.map((u) => u.fsPath);
              const resultString = results.join(";");
              quickPick.items = [
                {
                  label: config.prompt || localize("teamstoolkit.qm.selectFile"),
                  detail: resultString,
                },
              ];
              resolve(ok({ type: "success", result: results }));
            } else {
              const result = uriList[0].fsPath;
              quickPick.items = [
                {
                  label:
                    config.prompt ||
                    localize(
                      type === "folder"
                        ? "teamstoolkit.qm.selectFolder"
                        : "teamstoolkit.qm.selectFile"
                    ),
                  detail: result,
                },
              ];
              resolve(ok({ type: "success", result: result }));
            }
          } else {
            resolve(err(UserCancelError));
          }
        };
        const onDidChangeSelection = async function (
          items: readonly QuickPickItem[]
        ): Promise<any> {
          const defaultUrl = items[0].detail;
          await showFileSelectDialog(defaultUrl);
        };
        disposables.push(quickPick.onDidChangeSelection(onDidChangeSelection));
        disposables.push(quickPick);
        quickPick.show();
        await showFileSelectDialog(defaultValue);
      });
    } finally {
      disposables.forEach((d) => {
        d.dispose();
      });
    }
  }

  async openUrl(link: string): Promise<Result<boolean, FxError>> {
    const uri = Uri.parse(link);
    return new Promise(async (resolve) => {
      env.openExternal(uri).then((v) => {
        if (v) resolve(ok(v));
        else resolve(err(UserCancelError));
      });
    });
  }

  public async showMessage(
    level: "info" | "warn" | "error",
    message: string,
    modal: boolean,
    ...items: string[]
  ): Promise<Result<string | undefined, FxError>>;

  public async showMessage(
    level: "info" | "warn" | "error",
    message: Array<{ content: string; color: Colors }>,
    modal: boolean,
    ...items: string[]
  ): Promise<Result<string | undefined, FxError>>;

  async showMessage(
    level: "info" | "warn" | "error",
    message: string | Array<{ content: string; color: Colors }>,
    modal: boolean,
    ...items: string[]
  ): Promise<Result<string | undefined, FxError>> {
    if (message instanceof Array) {
      message = message.map((x) => x.content).join("");
    }
    return new Promise(async (resolve) => {
      const option = { modal: modal };
      try {
        let promise: Thenable<string | undefined>;
        switch (level) {
          case "info": {
            promise = window.showInformationMessage(message as string, option, ...items);
            break;
          }
          case "warn": {
            promise = window.showWarningMessage(message as string, option, ...items);
            break;
          }
          case "error":
            promise = window.showErrorMessage(message as string, option, ...items);
        }
        promise.then((v) => {
          if (v) resolve(ok(v));
          else resolve(err(UserCancelError));
        });
      } catch (error) {
        resolve(err(assembleError(error)));
      }
    });
  }

  public createProgressBar(title: string, totalSteps: number): IProgressHandler {
    return new ProgressHandler(title, totalSteps);
  }

  async runWithProgress<T>(
    task: RunnableTask<T>,
    config: TaskConfig,
    ...args: any
  ): Promise<Result<T, FxError>> {
    return new Promise(async (resolve) => {
      window.withProgress(
        {
          location: ProgressLocation.Notification,
          cancellable: config.cancellable,
        },
        async (progress, token): Promise<any> => {
          if (config.cancellable === true) {
            token.onCancellationRequested(() => {
              if (task.cancel) task.cancel();
              resolve(err(UserCancelError));
            });
          }
          let lastReport = 0;
          const showProgress = config.showProgress === true;
          const total = task.total ? task.total : 1;
          const head = task.name ? task.name : "";
          const report = (task: RunnableTask<T>) => {
            const current = task.current ? task.current : 0;
            const body = showProgress
              ? `: ${Math.round((current * 100) / total)} %`
              : `: [${current + 1}/${total}]`;
            const tail = task.message
              ? ` ${task.message}`
              : localize("teamstoolkit.progressHandler.prepareTask");
            const message = `${head}${body}${tail}`;
            if (showProgress)
              progress.report({
                increment: ((current - lastReport) * 100) / total,
                message: message,
              });
            else progress.report({ message: message });
          };
          task
            .run(args)
            .then(async (v) => {
              report(task);
              await sleep(100);
              resolve(v);
            })
            .catch((e) => {
              resolve(err(assembleError(e)));
            });
          let current;
          if (showProgress) {
            report(task);
            do {
              current = task.current ? task.current : 0;
              const inc = ((current - lastReport) * 100) / total;
              const delta = current - lastReport;
              if (inc > 0) {
                report(task);
                lastReport += delta;
              }
              await sleep(100);
            } while (current < total && !task.isCanceled);
            report(task);
            await sleep(100);
          } else {
            do {
              report(task);
              await sleep(100);
              current = task.current ? task.current : 0;
            } while (current < total && !task.isCanceled);
          }
          if (task.isCanceled) resolve(err(UserCancelError));
        }
      );
    });
  }

  async reload(): Promise<Result<boolean, FxError>> {
    return new Promise(async (resolve) => {
      // The following code only fixes the bug that cause telemetry event lost for projectMigrator().
      // When this reload() function has more users, they may need to dispose() more resources that allocated in activate().
      const extension = extensions.getExtension(`${packageJson.publisher}.${packageJson.name}`);
      if (!extension?.isActive) {
        // When our extension is not activated, we can determine this is in the vscode extension activate() context.
        // Since we are not activated yet, vscode will not deactivate() and dispose() our resourses (which have been allocated in activate()).
        // This may cause resource leaks.For example, buffered events in TelemetryReporter is not sent.
        // So manually dispose them.
        ExtTelemetry.reporter?.dispose();
      }

      commands.executeCommand("workbench.action.reloadWindow").then((v) => {
        if (v) resolve(ok(v as boolean));
        else resolve(err(UserCancelError));
      });
    });
  }

  async executeFunction(config: ExecuteFuncConfig) {
    const quickPick = window.createQuickPick<FxQuickPickItem>();
    quickPick.title = config.title;
    quickPick.busy = true;
    quickPick.show();
    try {
      return await config.func(config.inputs);
    } finally {
      quickPick.hide();
      quickPick.dispose();
    }
  }
}
