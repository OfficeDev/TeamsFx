// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  FolderQuestion,
  OptionItem,
  Platform,
  SingleSelectQuestion,
  TextInputQuestion,
  FuncQuestion,
  Inputs,
  LocalEnvironmentName,
  StaticOptions,
  MultiSelectQuestion,
} from "@microsoft/teamsfx-api";
import * as jsonschema from "jsonschema";
import * as path from "path";
import * as fs from "fs-extra";
import * as os from "os";
import { environmentManager } from "./environment";
import { ConstantString } from "../common/constants";
import { sampleProvider } from "../common/samples";
import { isAadManifestEnabled, isExistingTabAppEnabled, isM365AppEnabled } from "../common/tools";
import { isBotNotificationEnabled, isPreviewFeaturesEnabled } from "../common/featureFlags";
import { getLocalizedString } from "../common/localizeUtils";
import {
  BotOptionItem,
  MessageExtensionItem,
  NotificationOptionItem,
  TabOptionItem,
  TabSPFxItem,
  M365SsoLaunchPageOptionItem,
  M365SearchAppOptionItem,
  CommandAndResponseOptionItem,
  TabNonSsoItem,
  ExistingTabOptionItem,
  TabNewUIOptionItem,
  TabSPFxNewUIItem,
  MessageExtensionNewUIItem,
  BotNewUIOptionItem,
} from "../plugins/solution/fx-solution/question";
import { resourceGroupHelper } from "../plugins/solution/fx-solution/utils/ResourceGroupHelper";
import { ResourceManagementClient } from "@azure/arm-resources";

export enum CoreQuestionNames {
  AppName = "app-name",
  DefaultAppNameFunc = "default-app-name-func",
  Folder = "folder",
  ProjectPath = "projectPath",
  ProgrammingLanguage = "programming-language",
  Capabilities = "capabilities",
  Features = "features",
  Solution = "solution",
  CreateFromScratch = "scratch",
  Runtime = "runtime",
  Samples = "samples",
  Stage = "stage",
  SubStage = "substage",
  SourceEnvName = "sourceEnvName",
  TargetEnvName = "targetEnvName",
  TargetResourceGroupName = "targetResourceGroupName",
  NewResourceGroupName = "newResourceGroupName",
  NewResourceGroupLocation = "newResourceGroupLocation",
  NewTargetEnvName = "newTargetEnvName",
  ExistingTabEndpoint = "existing-tab-endpoint",
  SafeProjectName = "safeProjectName",
}

export const ProjectNamePattern =
  '^(?=(.*[\\da-zA-Z]){2})[a-zA-Z][^"<>:\\?/*|\u0000-\u001F]*[^"\\s.<>:\\?/*|\u0000-\u001F]$';

export function createAppNameQuestion(validateProjectPathExistence = true): TextInputQuestion {
  const question: TextInputQuestion = {
    type: "text",
    name: CoreQuestionNames.AppName,
    title: "Application name",
    validation: {
      validFunc: async (input: string, previousInputs?: Inputs): Promise<string | undefined> => {
        const schema = {
          pattern: ProjectNamePattern,
          maxLength: 30,
        };
        const appName = input as string;
        const validateResult = jsonschema.validate(appName, schema);
        if (validateResult.errors && validateResult.errors.length > 0) {
          if (validateResult.errors[0].name === "pattern") {
            return getLocalizedString("core.QuestionAppName.validation.pattern");
          }
          if (validateResult.errors[0].name === "maxLength") {
            return getLocalizedString("core.QuestionAppName.validation.maxlength");
          }
        }
        if (validateProjectPathExistence && previousInputs && previousInputs.folder) {
          const folder = previousInputs.folder as string;
          if (folder) {
            const projectPath = path.resolve(folder, appName);
            const exists = await fs.pathExists(projectPath);
            if (exists)
              return getLocalizedString("core.QuestionAppName.validation.pathExist", projectPath);
          }
        }
        return undefined;
      },
    },
    placeholder: "Application name",
  };
  return question;
}

export const DefaultAppNameFunc: FuncQuestion = {
  type: "func",
  name: CoreQuestionNames.DefaultAppNameFunc,
  func: (inputs: Inputs) => {
    const appName = path.basename(inputs.projectPath ?? "");
    const schema = {
      pattern: ProjectNamePattern,
      maxLength: 30,
    };
    const validateResult = jsonschema.validate(appName, schema);
    if (validateResult.errors && validateResult.errors.length > 0) {
      return undefined;
    }

    return appName;
  },
};

export const QuestionRootFolder: FolderQuestion = {
  type: "folder",
  name: CoreQuestionNames.Folder,
  title: getLocalizedString("core.question.workspaceFolder.title"),
  placeholder: getLocalizedString("core.question.workspaceFolder.placeholder"),
  default: path.join(os.homedir(), ConstantString.RootFolder),
};

export const ProgrammingLanguageQuestionForDotNet: SingleSelectQuestion = {
  name: CoreQuestionNames.ProgrammingLanguage,
  title: "Programming Language",
  type: "singleSelect",
  staticOptions: [{ id: "csharp", label: "C#" }],
  skipSingleOption: true,
};

export const ProgrammingLanguageQuestion: SingleSelectQuestion = {
  name: CoreQuestionNames.ProgrammingLanguage,
  title: "Programming Language",
  type: "singleSelect",
  staticOptions: [
    { id: "javascript", label: "JavaScript" },
    { id: "typescript", label: "TypeScript" },
  ],
  dynamicOptions: (inputs: Inputs): StaticOptions => {
    if (inputs.platform === Platform.VS) {
      return [{ id: "csharp", label: "C#" }];
    }
    const capabilities = inputs[CoreQuestionNames.Capabilities] as string[];
    if (capabilities && capabilities.includes && capabilities.includes(TabSPFxItem.id))
      return [{ id: "typescript", label: "TypeScript" }];
    return [
      { id: "javascript", label: "JavaScript" },
      { id: "typescript", label: "TypeScript" },
    ];
  },
  skipSingleOption: true,
  default: (inputs: Inputs) => {
    if (isPreviewFeaturesEnabled()) {
      const capability = inputs[CoreQuestionNames.Capabilities] as string;
      if (capability && capability === TabSPFxItem.id) {
        return "typescript";
      }
      const feature = inputs[CoreQuestionNames.Features] as string;
      if (feature && feature === TabSPFxItem.id) {
        return "typescript";
      }
    } else {
      const capabilities = inputs[CoreQuestionNames.Capabilities] as string[];
      if (capabilities && capabilities.includes && capabilities.includes(TabSPFxItem.id))
        return "typescript";
    }
    return "javascript";
  },
  placeholder: (inputs: Inputs): string => {
    if (isPreviewFeaturesEnabled()) {
      const capability = inputs[CoreQuestionNames.Capabilities] as string;
      if (capability && capability === TabSPFxItem.id) {
        return getLocalizedString("core.ProgrammingLanguageQuestion.placeholder.spfx");
      }
      const feature = inputs[CoreQuestionNames.Features] as string;
      if (feature && feature === TabSPFxItem.id) {
        return getLocalizedString("core.ProgrammingLanguageQuestion.placeholder.spfx");
      }
    } else {
      const capabilities = inputs[CoreQuestionNames.Capabilities] as string[];
      if (capabilities && capabilities.includes && capabilities.includes(TabSPFxItem.id))
        return getLocalizedString("core.ProgrammingLanguageQuestion.placeholder.spfx");
    }
    return getLocalizedString("core.ProgrammingLanguageQuestion.placeholder");
  },
};

function hasCapability(items: string[], optionItem: OptionItem): boolean {
  return items.includes(optionItem.id) || items.includes(optionItem.label);
}

function setIntersect<T>(set1: Set<T>, set2: Set<T>): Set<T> {
  return new Set([...set1].filter((item) => set2.has(item)));
}

function setDiff<T>(set1: Set<T>, set2: Set<T>): Set<T> {
  return new Set([...set1].filter((item) => !set2.has(item)));
}

function setUnion<T>(...sets: Set<T>[]): Set<T> {
  return new Set(([] as T[]).concat(...sets.map((set) => [...set])));
}

// Each set is mutually exclusive. Handle conflict by removing items conflicting with the newly added items.
// Assuming intersection of all sets are empty sets and no conflicts in newly added items.
//
// For example: sets = [[1, 2], [3, 4]], previous = [1, 2, 5], current = [1, 2, 4, 5].
// So the newly added one is [4]. Remove all items from `current` that conflict with [4].
// Result = [4, 5].
export function handleSelectionConflict<T>(
  sets: Set<T>[],
  previous: Set<T>,
  current: Set<T>
): Set<T> {
  const allSets = setUnion(...sets);
  const addedItems = setDiff(current, previous);

  for (const set of sets) {
    if (setIntersect(set, addedItems).size > 0) {
      return setUnion(setIntersect(set, current), setDiff(current, allSets));
    }
  }

  // If newly added items are not in any sets, do nothing.
  return current;
}

export function validateConflict<T>(sets: Set<T>[], current: Set<T>): string | undefined {
  const all = setUnion(...sets);
  const currentIntersectAll = setIntersect(all, current);
  for (const set of sets) {
    if (setIntersect(set, current).size > 0) {
      const currentIntersectSet = setIntersect(set, current);
      if (currentIntersectSet.size < currentIntersectAll.size) {
        return getLocalizedString(
          "core.capability.validation",
          `[${Array.from(current).join(", ")}]`,
          Array.from(sets)
            .map((set) => `[${Array.from(set).join(", ")}]`)
            .join(", ")
        );
      }
    }
  }
  return undefined;
}

export function createCapabilityQuestion(): MultiSelectQuestion {
  let staticOptions: StaticOptions;
  if (isBotNotificationEnabled()) {
    // new capabilities question order
    staticOptions = [
      ...[CommandAndResponseOptionItem, NotificationOptionItem],
      ...(isExistingTabAppEnabled() ? [ExistingTabOptionItem] : []),
      ...(isAadManifestEnabled() ? [TabNonSsoItem] : []),
      ...[TabNewUIOptionItem, TabSPFxNewUIItem, MessageExtensionNewUIItem],
      ...(isM365AppEnabled() ? [M365SsoLaunchPageOptionItem, M365SearchAppOptionItem] : []),
    ];
  } else {
    staticOptions = [
      ...[TabOptionItem, BotOptionItem, MessageExtensionItem, TabSPFxItem],
      ...(isAadManifestEnabled() ? [TabNonSsoItem] : []),
      ...(isExistingTabAppEnabled() ? [ExistingTabOptionItem] : []),
      ...(isM365AppEnabled() ? [M365SsoLaunchPageOptionItem, M365SearchAppOptionItem] : []),
    ];
  }
  return {
    name: CoreQuestionNames.Capabilities,
    title: isBotNotificationEnabled()
      ? getLocalizedString("core.createCapabilityQuestion.titleNew")
      : getLocalizedString("core.createCapabilityQuestion.title"),
    type: "multiSelect",
    staticOptions: staticOptions,
    default: isBotNotificationEnabled() ? [CommandAndResponseOptionItem.id] : [TabOptionItem.id],
    placeholder: getLocalizedString("core.createCapabilityQuestion.placeholder"),
    validation: {
      validFunc: validateCapabilities,
    },
    onDidChangeSelection: onChangeSelectionForCapabilities,
  };
}

export function createCapabilityForDotNet(): SingleSelectQuestion {
  const staticOptions: StaticOptions = [
    NotificationOptionItem,
    CommandAndResponseOptionItem,
    TabOptionItem,
    MessageExtensionItem,
  ];
  return {
    name: CoreQuestionNames.Capabilities,
    title: getLocalizedString("core.createCapabilityQuestion.title"),
    type: "singleSelect",
    staticOptions: staticOptions,
    placeholder: getLocalizedString("core.createCapabilityQuestion.placeholder"),
  };
}

export function createCapabilityQuestionPreview(): SingleSelectQuestion {
  // new capabilities question order
  const staticOptions: StaticOptions = [
    NotificationOptionItem,
    CommandAndResponseOptionItem,
    TabNewUIOptionItem,
    TabSPFxNewUIItem,
    TabNonSsoItem,
    BotNewUIOptionItem,
    MessageExtensionNewUIItem,
    M365SsoLaunchPageOptionItem,
    M365SearchAppOptionItem,
  ];

  if (isExistingTabAppEnabled()) {
    staticOptions.splice(2, 0, ExistingTabOptionItem);
  }

  return {
    name: CoreQuestionNames.Capabilities,
    title: getLocalizedString("core.createCapabilityQuestion.titleNew"),
    type: "singleSelect",
    staticOptions: staticOptions,
    placeholder: getLocalizedString("core.createCapabilityQuestion.placeholder"),
  };
}

export function validateCapabilities(inputs: string[]): string | undefined {
  if (inputs.length === 0) {
    return getLocalizedString("core.createCapabilityQuestion.placeholder");
  }
  const set = new Set<string>();
  inputs.forEach((i) => set.add(i));
  let result = validateConflict(
    [
      new Set([BotOptionItem.id, MessageExtensionItem.id]),
      new Set([NotificationOptionItem.id]),
      new Set([CommandAndResponseOptionItem.id]),
    ],
    set
  );
  if (result) return result;
  result = validateConflict(
    [
      new Set([
        TabOptionItem.id,
        TabNonSsoItem.id,
        BotOptionItem.id,
        MessageExtensionItem.id,
        NotificationOptionItem.id,
        CommandAndResponseOptionItem.id,
      ]),
      new Set([TabSPFxItem.id]),
    ],
    set
  );
  if (result) return result;
  result = validateConflict([new Set([TabOptionItem.id]), new Set([TabNonSsoItem.id])], set);
  if (result) return result;
  result = validateConflict(
    [
      new Set([
        TabOptionItem.id,
        TabNonSsoItem.id,
        TabSPFxItem.id,
        BotOptionItem.id,
        MessageExtensionItem.id,
        NotificationOptionItem.id,
        CommandAndResponseOptionItem.id,
        ExistingTabOptionItem.id,
      ]),
      new Set([M365SsoLaunchPageOptionItem.id]),
      new Set([M365SearchAppOptionItem.id]),
    ],
    set
  );
  return result;
}

export async function onChangeSelectionForCapabilities(
  currentSelectedIds: Set<string>,
  previousSelectedIds: Set<string>
): Promise<Set<string>> {
  let result = handleSelectionConflict(
    [
      new Set([BotOptionItem.id, MessageExtensionItem.id]),
      new Set([NotificationOptionItem.id]),
      new Set([CommandAndResponseOptionItem.id]),
    ],
    previousSelectedIds,
    currentSelectedIds
  );
  result = handleSelectionConflict(
    [
      new Set([
        TabOptionItem.id,
        TabNonSsoItem.id,
        BotOptionItem.id,
        MessageExtensionItem.id,
        NotificationOptionItem.id,
        CommandAndResponseOptionItem.id,
      ]),
      new Set([TabSPFxItem.id]),
      new Set([ExistingTabOptionItem.id]),
      new Set([M365SsoLaunchPageOptionItem.id]),
      new Set([M365SearchAppOptionItem.id]),
    ],
    previousSelectedIds,
    result
  );
  result = handleSelectionConflict(
    [new Set([TabOptionItem.id]), new Set([TabNonSsoItem.id])],
    previousSelectedIds,
    result
  );
  return result;
}
export const QuestionSelectTargetEnvironment: SingleSelectQuestion = {
  type: "singleSelect",
  name: CoreQuestionNames.TargetEnvName,
  title: getLocalizedString("core.QuestionSelectTargetEnvironment.title"),
  staticOptions: [],
  skipSingleOption: true,
  forgetLastValue: true,
};

export function getQuestionNewTargetEnvironmentName(projectPath: string): TextInputQuestion {
  const WINDOWS_MAX_PATH_LENGTH = 260;
  return {
    type: "text",
    name: CoreQuestionNames.NewTargetEnvName,
    title: getLocalizedString("core.getQuestionNewTargetEnvironmentName.title"),
    validation: {
      validFunc: async (input: string): Promise<string | undefined> => {
        const targetEnvName = input;
        const match = targetEnvName.match(environmentManager.envNameRegex);
        if (!match) {
          return getLocalizedString("core.getQuestionNewTargetEnvironmentName.validation1");
        }

        const envFilePath = environmentManager.getEnvConfigPath(targetEnvName, projectPath);
        if (os.type() === "Windows_NT" && envFilePath.length >= WINDOWS_MAX_PATH_LENGTH) {
          return getLocalizedString("core.getQuestionNewTargetEnvironmentName.validation2");
        }

        if (targetEnvName === LocalEnvironmentName) {
          return getLocalizedString(
            "core.getQuestionNewTargetEnvironmentName.validation3",
            LocalEnvironmentName
          );
        }

        const envConfigs = await environmentManager.listRemoteEnvConfigs(projectPath);
        if (envConfigs.isErr()) {
          return getLocalizedString("core.getQuestionNewTargetEnvironmentName.validation4");
        }

        const found =
          envConfigs.value.find(
            (env) => env.localeCompare(targetEnvName, undefined, { sensitivity: "base" }) === 0
          ) !== undefined;
        if (found) {
          return getLocalizedString(
            "core.getQuestionNewTargetEnvironmentName.validation5",
            targetEnvName
          );
        } else {
          return undefined;
        }
      },
    },
    placeholder: getLocalizedString("core.getQuestionNewTargetEnvironmentName.placeholder"),
  };
}

export const QuestionSelectSourceEnvironment: SingleSelectQuestion = {
  type: "singleSelect",
  name: CoreQuestionNames.SourceEnvName,
  title: getLocalizedString("core.QuestionSelectSourceEnvironment.title"),
  staticOptions: [],
  skipSingleOption: true,
  forgetLastValue: true,
};

export const QuestionSelectResourceGroup: SingleSelectQuestion = {
  type: "singleSelect",
  name: CoreQuestionNames.TargetResourceGroupName,
  title: getLocalizedString("core.QuestionSelectResourceGroup.title"),
  staticOptions: [],
  skipSingleOption: true,
  forgetLastValue: true,
};
export function newResourceGroupNameQuestion(
  rmClient: ResourceManagementClient
): TextInputQuestion {
  const question = QuestionNewResourceGroupName;
  question.validation = {
    validFunc: async (input: string): Promise<string | undefined> => {
      const name = input as string;
      // https://docs.microsoft.com/en-us/rest/api/resources/resource-groups/create-or-update#uri-parameters
      const match = name.match(/^[-\w._()]+$/);
      if (!match) {
        return getLocalizedString("core.QuestionNewResourceGroupName.validation");
      }
      const maybeExist = await resourceGroupHelper.checkResourceGroupExistence(name, rmClient);
      if (maybeExist.isErr()) {
        return maybeExist.error.message;
      }
      if (maybeExist.value) {
        return `resource group already exists: ${name}`;
      }
      return undefined;
    },
  };
  return question;
}
export const QuestionNewResourceGroupName: TextInputQuestion = {
  type: "text",
  name: CoreQuestionNames.NewResourceGroupName,
  title: getLocalizedString("core.QuestionNewResourceGroupName.title"),
  placeholder: getLocalizedString("core.QuestionNewResourceGroupName.placeholder"),
  // default resource group name will change with env name
  forgetLastValue: true,
};

export const QuestionNewResourceGroupLocation: SingleSelectQuestion = {
  type: "singleSelect",
  name: CoreQuestionNames.NewResourceGroupLocation,
  title: getLocalizedString("core.QuestionNewResourceGroupLocation.title"),
  staticOptions: [],
};

export const ScratchOptionYesVSC: OptionItem = {
  id: "yes",
  label: `$(new-folder) ${getLocalizedString("core.ScratchOptionYesVSC.label")}`,
  detail: getLocalizedString("core.ScratchOptionYesVSC.detail"),
};

export const ScratchOptionNoVSC: OptionItem = {
  id: "no",
  label: `$(heart) ${getLocalizedString("core.ScratchOptionNoVSC.label")}`,
  detail: getLocalizedString("core.ScratchOptionNoVSC.detail"),
};

export const RuntimeOptionNodeJs: OptionItem = {
  id: "node",
  label: "Node.js",
  detail: getLocalizedString("core.RuntimeOptionNodeJS.detail"),
};

export const RuntimeOptionDotNet: OptionItem = {
  id: "dotnet",
  label: ".NET Core",
  detail: getLocalizedString("core.RuntimeOptionDotNet.detail"),
};

export const ScratchOptionYes: OptionItem = {
  id: "yes",
  label: getLocalizedString("core.ScratchOptionYes.label"),
  detail: getLocalizedString("core.ScratchOptionYes.detail"),
};

export const ScratchOptionNo: OptionItem = {
  id: "no",
  label: getLocalizedString("core.ScratchOptionNo.label"),
  detail: getLocalizedString("core.ScratchOptionNo.detail"),
};

// This question should only exist on CLI
export function getRuntimeQuestion(): SingleSelectQuestion {
  return {
    type: "singleSelect",
    name: CoreQuestionNames.Runtime,
    title: getLocalizedString("core.getRuntimeQuestion.title"),
    staticOptions: [RuntimeOptionNodeJs, RuntimeOptionDotNet],
    default: RuntimeOptionNodeJs.id,
    placeholder: getLocalizedString("core.getRuntimeQuestion.placeholder"),
  };
}

export function getCreateNewOrFromSampleQuestion(platform: Platform): SingleSelectQuestion {
  const staticOptions: OptionItem[] = [];
  if (platform === Platform.VSCode) {
    staticOptions.push(ScratchOptionYesVSC);
    staticOptions.push(ScratchOptionNoVSC);
  } else {
    staticOptions.push(ScratchOptionYes);
    staticOptions.push(ScratchOptionNo);
  }
  return {
    type: "singleSelect",
    name: CoreQuestionNames.CreateFromScratch,
    title: getLocalizedString("core.getCreateNewOrFromSampleQuestion.title"),
    staticOptions,
    default: ScratchOptionYes.id,
    placeholder: getLocalizedString("core.getCreateNewOrFromSampleQuestion.placeholder"),
    skipSingleOption: true,
  };
}

export const SampleSelect: SingleSelectQuestion = {
  type: "singleSelect",
  name: CoreQuestionNames.Samples,
  title: getLocalizedString("core.SampleSelect.title"),
  staticOptions: sampleProvider.SampleCollection.samples.map((sample) => {
    return {
      id: sample.id,
      label: sample.title,
      description: `${sample.time} • ${sample.configuration}`,
      detail: sample.shortDescription,
      data: sample.link,
    } as OptionItem;
  }),
  placeholder: getLocalizedString("core.SampleSelect.placeholder"),
  buttons: [
    {
      icon: "library",
      tooltip: getLocalizedString("core.SampleSelect.buttons.viewSamples"),
      command: "fx-extension.openSamples",
    },
  ],
};

export const ExistingTabEndpointQuestion: TextInputQuestion = {
  type: "text",
  name: CoreQuestionNames.ExistingTabEndpoint,
  title: getLocalizedString("core.ExistingTabEndpointQuestion.title"),
  default: "https://localhost:3000",
  placeholder: getLocalizedString("core.ExistingTabEndpointQuestion.placeholder"),
  validation: {
    validFunc: async (endpoint: string): Promise<string | undefined> => {
      const match = endpoint.match(/^https:\/\/[\S]+$/i);
      if (!match) {
        return getLocalizedString("core.ExistingTabEndpointQuestion.validation");
      }

      return undefined;
    },
  },
};
