// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { FuncQuestion, MultiSelectQuestion, NodeType, OptionItem, SingleSelectQuestion } from "fx-api";

export const TabOptionItem: OptionItem = {
    id: "Tab",
    label: "Tab",
    cliName: "tab",
    description: "Tabs embeds a web app experience in a tab in a Teams chat, channel, or personal workspace.",
};

export const BotOptionItem: OptionItem = {
    id: "Bot",
    label: "Bot",
    cliName: "bot",
    description:
        "Bots allow you to interact with and obtain information in a text/search/conversational manner.",
};

export const MessageExtensionItem: OptionItem = {
    id: "MessageExtension",
    label: "Messaging Extension",
    description:
        "Messaging Extensions allow users to interact with a web service through buttons and forms in the Microsoft Teams client.",
};

export enum AzureSolutionQuestionNames {
    Capabilities = "capabilities",
    TabScopes = "tab-scopes",
    HostType = "host-type",
    AzureResources = "azure-resources",
    PluginSelectionDeploy = "deploy-plugin",
    AddResources = "add-azure-resources",
    AppName = "app-name",
    AskSub = "subscription",
    ProgrammingLanguage = "programming-language",
}

export const HostTypeOptionAzure: OptionItem = {
    id:"Azure",
    label: "Azure",
    cliName: "azure",
    description: "Azure Cloud",
};

export const HostTypeOptionSPFx: OptionItem = {
    id:"SPFx",
    label: "SPFx",
    cliName: "spfx",
    description: "SharePoint Framework",
};

export const AzureResourceSQL: OptionItem = {
    id:"sql",
    label: "Azure SQL Database",
    description: "Azure SQL Database depends on Azure Functions.",
};

export const AzureResourceFunction: OptionItem = {
    id:"function",
    label: "Azure Functions",
    description: "Application backend.",
};

export const AzureResourceApim: OptionItem = {
    id:"apim",
    label: "apim",
    description: "Register APIs in Azure API Management",
};
 
export function createCapabilityQuestion(): MultiSelectQuestion {
    return {
        name: AzureSolutionQuestionNames.Capabilities,
        title: "Add capabilities",
        prompt: "Choose capabilities for your application",
        type: NodeType.multiSelect,
        option: [TabOptionItem, BotOptionItem, MessageExtensionItem],
        default: [TabOptionItem.id]
    };
}

export const FrontendHostTypeQuestion: SingleSelectQuestion = {
    name: AzureSolutionQuestionNames.HostType,
    title: "Select frontend hosting type",
    type: NodeType.singleSelect,
    option: [HostTypeOptionAzure, HostTypeOptionSPFx],
    default: HostTypeOptionAzure.id,
};

export const AzureResourcesQuestion: MultiSelectQuestion = {
    name: AzureSolutionQuestionNames.AzureResources,
    title: "Additional cloud resources",
    type: NodeType.multiSelect,
    option: [AzureResourceSQL, AzureResourceFunction],
    default: [],
    onDidChangeSelection:async function(selectedItems: OptionItem[]) : Promise<string[]>{
        const hasSQL = selectedItems.some(i=>i.id === AzureResourceSQL.id);
        if(hasSQL){
            return [AzureResourceSQL.id, AzureResourceFunction.id];
        }
        return selectedItems.map(i=>i.id);
    }
};

// export const AddAzureResourceQuestion: MultiSelectQuestion = {
//     name: AzureSolutionQuestionNames.AddResources,
//     title: 'Select Azure resources to add',
//     type: NodeType.multiSelect,
//     option: [AzureResourceSQL, AzureResourceFunction, AzureResourceApim],
//     default: [],
// };

export function createAddAzureResourceQuestion(alreadyHaveFunction: boolean, alreadhHaveSQL: boolean, alreadyHaveAPIM: boolean): MultiSelectQuestion {
    const options:OptionItem[] = [AzureResourceFunction];
    if(!alreadhHaveSQL) options.push(AzureResourceSQL);
    if(!alreadyHaveAPIM) options.push(AzureResourceApim);
    return {
        name: AzureSolutionQuestionNames.AddResources,
        title: "Select Azure resources to add",
        type: NodeType.multiSelect,
        option: options,
        default: [],
        onDidChangeSelection:async function(selectedItems: OptionItem[]) : Promise<string[]>{
            const hasSQL = selectedItems.some(i=>i.id === AzureResourceSQL.id);
            const hasAPIM = selectedItems.some(i=>i.id === AzureResourceApim.id);
            const ids = selectedItems.map(i=>i.id);
            /// when SQL or APIM is selected and function is not selected, then function must be selected
            if( (hasSQL||hasAPIM) && !alreadyHaveFunction && !ids.includes(AzureResourceFunction.id)){
                ids.push(AzureResourceFunction.id);
            }
            return ids;
        }
    };
}

export function createAddCapabilityQuestion(alreadyHaveTab: boolean, alreadyHaveBot: boolean): MultiSelectQuestion {
    const options:OptionItem[] = [];
    if(!alreadyHaveTab) options.push(TabOptionItem);
    if(!alreadyHaveBot) options.push(BotOptionItem);
    return {
        name: AzureSolutionQuestionNames.Capabilities,
        title: "Select capabilities to add",
        type: NodeType.multiSelect,
        option: options,
        default: []
    };
}

export const DeployPluginSelectQuestion: MultiSelectQuestion = {
    name: AzureSolutionQuestionNames.PluginSelectionDeploy,
    title: `Select resource(s) to deploy`,
    type: NodeType.multiSelect,
    skipSingleOption: true,
    option: [],
    default: []
};


export const AskSubscriptionQuestion: FuncQuestion = {
    name: AzureSolutionQuestionNames.AskSub,
    title: "Please select a subscription",
    type: NodeType.func,
    namespace: "fx-solution-azure",
    method: "askSubscription"
};

export const ProgrammingLanguageQuestion: SingleSelectQuestion = {
    name: AzureSolutionQuestionNames.ProgrammingLanguage,
    title: "Select programming language for your project",
    type: NodeType.singleSelect,
    option: ["javascript", "typescript"],
    default: "javascript",
    skipSingleOption: true
};
