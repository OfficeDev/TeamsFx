// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
 
import { Disposable, InputBox, QuickInputButtons, QuickPick, QuickPickItem, Uri, window } from "vscode";
import { FxInputBoxOption, FxOpenDialogOption, FxQuickPickOption, InputResult, InputResultType, OptionItem, returnSystemError, UserInterface } from "teamsfx-api";
import { ExtensionErrors, ExtensionSource } from "../error";
 

 
export interface FxQuickPickItem extends QuickPickItem {
  id: string;
  data?: unknown;
}

export class VsCodeUi implements UserInterface{
  async showQuickPick (option: FxQuickPickOption) : Promise<InputResult>{
    const disposables: Disposable[] = [];
    // let isPrompting = false;
    try {
      const quickPick: QuickPick<QuickPickItem> = window.createQuickPick();
      disposables.push(quickPick);
      quickPick.title = option.title;
      if (option.backButton) quickPick.buttons = [QuickInputButtons.Back];
      quickPick.placeholder = option.placeholder;
      quickPick.ignoreFocusOut = true;
      quickPick.matchOnDescription = true;
      quickPick.matchOnDetail = true;
      quickPick.canSelectMany = option.canSelectMany;
  
      return await new Promise<InputResult>(
        async (resolve): Promise<void> => {
          disposables.push(
            quickPick.onDidAccept(async () => {
              const selectedItems = quickPick.selectedItems as FxQuickPickItem[];
              if (option.canSelectMany) {
                const strArray = Array.from(selectedItems.map((i) => i.id));
                let result: OptionItem[] | string[] = strArray;
                if (option.returnObject) {
                  result = selectedItems.map((i) => {
                    const item: OptionItem = {
                      id: i.id,
                      label: i.label,
                      description: i.description,
                      detail: i.detail,
                      data: i.data
                    };
                    return item;
                  });
                }
                resolve({
                  type: InputResultType.sucess,
                  result: result
                });
              } else {
                const item: FxQuickPickItem = quickPick.selectedItems[0] as FxQuickPickItem;
                let result: string | OptionItem = item.id;
                if (option.returnObject) {
                  result = {
                    id: item.id,
                    label: item.label,
                    description: item.description,
                    detail: item.detail,
                    data: item.data
                  };
                }
                resolve({ type: InputResultType.sucess, result: result });
              }
            }),
            quickPick.onDidHide(() => {
              resolve({ type: InputResultType.cancel });
            })
          );
          if (option.backButton) {
            disposables.push(
              quickPick.onDidTriggerButton((_btn) => {
                resolve({ type: InputResultType.back });
              })
            );
          }
          // isPrompting = true;
          try {
            const isStringArray = !!(typeof option.items[0] === "string");
            /// set items
            if (isStringArray) {
              quickPick.items = (option.items as string[]).map((i: string) => {
                return { label: i, id: i };
              });
            } else {
              quickPick.items = (option.items as OptionItem[]).map((i: OptionItem) => {
                return {
                  id: i.id,
                  label: i.label,
                  description: i.description,
                  detail: i.detail,
                  data: i.data
                };
              });
            }
  
            /// set default values
            if (option.defaultValue) {
              const modsItems = quickPick.items as FxQuickPickItem[];
              if (option.canSelectMany) {
                const defaultStringArrayValue = option.defaultValue as string[];
                quickPick.selectedItems = modsItems.filter((i) =>
                  defaultStringArrayValue.includes(i.id)
                );
              } else {
                const defaultStringValue = option.defaultValue as string;
                const newitems = modsItems.filter((i) => i.id !== defaultStringValue);
                for (const i of modsItems) {
                  if (i.id === defaultStringValue) {
                    newitems.unshift(i);
                    break;
                  }
                }
                quickPick.items = newitems;
              }
            }
            quickPick.show();
          } catch (err) {
            resolve({
              type: InputResultType.error,
              error: returnSystemError(err, ExtensionSource, ExtensionErrors.UnknwonError)
            });
          }
        }
      );
    } finally {
      // isPrompting = false;
      disposables.forEach((d) => {
        d.dispose();
      });
    }
  }


  async showInputBox(option: FxInputBoxOption) : Promise<InputResult>{
    const disposables: Disposable[] = [];
    // let isPrompting = false;
    try {
      const inputBox: InputBox = window.createInputBox();
      disposables.push(inputBox);
      inputBox.title = option.title;
      if (option.backButton) inputBox.buttons = [QuickInputButtons.Back];
      inputBox.value = option.defaultValue || "";
      inputBox.ignoreFocusOut = true;
      inputBox.password = option.password;
      inputBox.placeholder = option.placeholder;
      inputBox.prompt = option.prompt;
      let latestValidation: Promise<string | undefined | null> = option.validation
        ? Promise.resolve(await option.validation(inputBox.value))
        : Promise.resolve("");
      return await new Promise<InputResult>((resolve, reject): void => {
        disposables.push(
          inputBox.onDidChangeValue(async (text) => {
            if (option.validation) {
              const validationRes: Promise<string | undefined | null> = Promise.resolve(
                await option.validation(text)
              );
              latestValidation = validationRes;
              let message: string | undefined | null = await validationRes;
              if(message === undefined && option.number){
                const num = Number(text);
                if(isNaN(num)){
                  message = text + " is not a valid number";
                }
              }
              if (validationRes === latestValidation) {
                inputBox.validationMessage = message || "";
              }
            }
          }),
          inputBox.onDidAccept(async () => {
            // Run final validation and resolve if value passes
            inputBox.enabled = false;
            inputBox.busy = true;
            const message: string | undefined | null = await latestValidation;
            if (!message) {
              resolve({ type: InputResultType.sucess, result: inputBox.value });
            } else {
              inputBox.validationMessage = message;
            }
            inputBox.enabled = true;
            inputBox.busy = false;
          }),
          inputBox.onDidHide(() => {
            resolve({ type: InputResultType.cancel });
          })
        );
        if (option.backButton) {
          disposables.push(
            inputBox.onDidTriggerButton((_btn) => {
              resolve({ type: InputResultType.back });
            })
          );
        }
        inputBox.show();
        // isPrompting = true;
      });
    } finally {
      // isPrompting = false;
      disposables.forEach((d) => {
        d.dispose();
      });
    }
  }

  async showOpenDialog (option: FxOpenDialogOption):Promise<InputResult>{
    while (true) {
      const uri = await window.showOpenDialog({
        defaultUri: option.defaultUri ? Uri.file(option.defaultUri) : undefined,
        canSelectFiles: false,
        canSelectFolders: true,
        canSelectMany: false,
        title: option.title
      });
      const res = uri && uri.length > 0 ? uri[0].fsPath : undefined;
      if (!res) {
        return { type: InputResultType.cancel };
      }
      if(!option.validation){
        return { type: InputResultType.sucess, result: res };
      }
      const validationRes = await option.validation(res);
      if (!validationRes) {
        return { type: InputResultType.sucess, result: res };
      }
      else {
        await window.showErrorMessage(validationRes);
      }
    }
  }
}
  
