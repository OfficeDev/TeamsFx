import { MultiSelectQuestion, OptionItem, SingleSelectQuestion } from "@microsoft/teamsfx-api";

export enum QuestionName {
  ExampleSingleSelectQuestion = "example-single-select",
  ExampleMultiSelectQuestion = "example-multi-select",
}

export const SingleSelectOptionOne: OptionItem = {
  id: "option1",
  label: "Option 1 label",
  detail: "Option 1 detail",
  groupName: "group1",
};

export const SingleSelectOptionTwo: OptionItem = {
  id: "option2",
  label: "Option 2 label",
  detail: "Option 2 detail",
  groupName: "group1",
};

export const SingleSelectOptionThree: OptionItem = {
  id: "option3",
  label: "Option 3 label",
  detail: "Option 3 detail",
  groupName: "group2",
};

// TODO: localize the strings
export const ExampleSingleSelectQuestion: SingleSelectQuestion = {
  type: "singleSelect",
  name: QuestionName.ExampleSingleSelectQuestion,
  title: "This is a single select question",
  staticOptions: [SingleSelectOptionOne, SingleSelectOptionTwo, SingleSelectOptionThree],
  default: SingleSelectOptionOne.id,
  placeholder: "This is placeholder",
};

export const MultiSelectOptionOne: OptionItem = {
  id: "multi-option1",
  label: "Option 1 label",
  detail: "Option 1 detail",
};

export const MultiSelectOptionTwo: OptionItem = {
  id: "multi-option2",
  label: "Option 2 label",
  detail: "Option 2 detail",
};

export const ExampleMultiSelectQuestion: MultiSelectQuestion = {
  name: QuestionName.ExampleMultiSelectQuestion,
  title: "This is a multi-select question",
  type: "multiSelect",
  staticOptions: [MultiSelectOptionOne, MultiSelectOptionTwo],
  default: undefined,
  placeholder: "This is placeholder",
};
