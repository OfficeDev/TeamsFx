// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { AssertConfigNotEmpty, BuildError, NoValidOpenApiDocument } from "../error";
import {
    LogProvider,
    Dialog,
    OptionItem,
    SingleSelectQuestion,
    NodeType,
    Question,
    Validation,
    PluginContext,
    FuncQuestion,
    TextInputQuestion,
} from "fx-api";
import { ApimDefaultValues, ApimPluginConfigKeys, QuestionConstants, TeamsToolkitComponent } from "../constants";
import { ApimPluginConfig, SolutionConfig } from "../model/config";
import { ApimService } from "./apimService";
import { OpenApiProcessor } from "../util/openApiProcessor";
import { NameSanitizer } from "../util/nameSanitizer";
import { Telemetry } from "../telemetry";
import { buildAnswer } from "../model/answer";

export interface IQuestionService {
    // Control whether the question is displayed to the user.
    condition?(parentAnswerPath: string): { target?: string; } & Validation;

    // Define the method name
    funcName: string;

    // Generate the options / default value / answer of the question.
    executeFunc(ctx: PluginContext): Promise<string | OptionItem | OptionItem[]>;

    // Generate the question
    getQuestion(): Question;

    // Validate the answer of the question.
    validate?(answer: string): boolean;
}

class BaseQuestionService {
    protected readonly dialog: Dialog;
    protected readonly logger?: LogProvider;
    protected readonly telemetry?: Telemetry;

    constructor(dialog: Dialog, telemetry?: Telemetry, logger?: LogProvider) {
        this.dialog = dialog;
        this.telemetry = telemetry;
        this.logger = logger;
    }
}

export class ApimServiceQuestion extends BaseQuestionService implements IQuestionService {
    private readonly apimService: ApimService;
    public readonly funcName = QuestionConstants.Apim.funcName;

    constructor(apimService: ApimService, dialog: Dialog, telemetry: Telemetry, logger?: LogProvider) {
        super(dialog, telemetry, logger);
        this.apimService = apimService;
    }

    public async executeFunc(ctx: PluginContext): Promise<OptionItem[]> {
        const apimServiceList = await this.apimService.listService();
        const existingOptions = apimServiceList.map((apimService) => {
            return { id: apimService.serviceName, label: apimService.serviceName, description: apimService.resourceGroupName, data: apimService };
        });
        const newOption = { id: QuestionConstants.Apim.createNewApimOption, label: QuestionConstants.Apim.createNewApimOption };
        return [newOption, ...existingOptions];
    }

    public getQuestion(): SingleSelectQuestion {
        return {
            type: NodeType.singleSelect,
            name: QuestionConstants.Apim.questionName,
            description: QuestionConstants.Apim.description,
            option: {
                namespace: QuestionConstants.namespace,
                method: QuestionConstants.Apim.funcName,
            },
            returnObject: true,
            skipSingleOption: false
        };
    }
}

export class OpenApiDocumentQuestion extends BaseQuestionService implements IQuestionService {
    private readonly openApiProcessor: OpenApiProcessor;
    public readonly funcName = QuestionConstants.OpenApiDocument.funcName;

    constructor(openApiProcessor: OpenApiProcessor, dialog: Dialog, telemetry: Telemetry, logger?: LogProvider) {
        super(dialog, telemetry, logger);
        this.openApiProcessor = openApiProcessor;
    }

    public async executeFunc(ctx: PluginContext): Promise<OptionItem[]> {
        const filePath2OpenApiMap = await this.openApiProcessor.listOpenApiDocument(
            ctx.root,
            QuestionConstants.OpenApiDocument.excludeFolders,
            QuestionConstants.OpenApiDocument.openApiDocumentFileExtensions
        );

        if (filePath2OpenApiMap.size === 0) {
            throw BuildError(NoValidOpenApiDocument);
        }

        const result: OptionItem[] = [];
        filePath2OpenApiMap.forEach((value, key) => result.push({ id: key, label: key, data: value }));
        return result;
    }

    public getQuestion(): SingleSelectQuestion {
        return {
            type: NodeType.singleSelect,
            name: QuestionConstants.OpenApiDocument.questionName,
            description: QuestionConstants.OpenApiDocument.description,
            option: {
                namespace: QuestionConstants.namespace,
                method: QuestionConstants.OpenApiDocument.funcName,
            },
            returnObject: true,
            skipSingleOption: false
        };
    }
}

export class ExistingOpenApiDocumentFunc extends BaseQuestionService implements IQuestionService {
    private readonly openApiProcessor: OpenApiProcessor;
    public readonly funcName = QuestionConstants.ExistingOpenApiDocument.funcName;

    constructor(openApiProcessor: OpenApiProcessor, dialog: Dialog, telemetry: Telemetry, logger?: LogProvider) {
        super(dialog, telemetry, logger);
        this.openApiProcessor = openApiProcessor;
    }

    public async executeFunc(ctx: PluginContext): Promise<OptionItem> {
        const apimConfig = new ApimPluginConfig(ctx.config);
        const openApiDocumentPath = AssertConfigNotEmpty(
            TeamsToolkitComponent.ApimPlugin,
            ApimPluginConfigKeys.apiDocumentPath,
            apimConfig.apiDocumentPath
        );
        const openApiDocument = await this.openApiProcessor.loadOpenApiDocument(openApiDocumentPath, ctx.root);
        return { id: openApiDocumentPath, label: openApiDocumentPath, data: openApiDocument };
    }

    public getQuestion(): FuncQuestion {
        return {
            type: NodeType.func,
            name: QuestionConstants.ExistingOpenApiDocument.questionName,
            namespace: QuestionConstants.namespace,
            method: QuestionConstants.ExistingOpenApiDocument.funcName,
        };
    }
}

export class ApiPrefixQuestion extends BaseQuestionService implements IQuestionService {
    public readonly funcName = QuestionConstants.ApiPrefix.funcName;

    constructor(dialog: Dialog, telemetry: Telemetry, logger?: LogProvider) {
        super(dialog, telemetry, logger);
    }

    public async executeFunc(ctx: PluginContext): Promise<string> {
        const apiTitle = buildAnswer(ctx)?.openApiDocumentSpec?.info.title;
        return !!apiTitle ? NameSanitizer.sanitizeApiNamePrefix(apiTitle) : ApimDefaultValues.apiPrefix;
    }

    public getQuestion(): TextInputQuestion {
        return {
            type: NodeType.text,
            name: QuestionConstants.ApiPrefix.questionName,
            description: QuestionConstants.ApiPrefix.description,
            default: {
                namespace: QuestionConstants.namespace,
                method: QuestionConstants.ApiPrefix.funcName,
            },
        };
    }
}

export class ApiVersionQuestion extends BaseQuestionService implements IQuestionService {
    private readonly apimService: ApimService;
    public readonly funcName = QuestionConstants.ApiVersion.funcName;

    constructor(apimService: ApimService, dialog: Dialog, telemetry: Telemetry, logger?: LogProvider) {
        super(dialog, telemetry, logger);
        this.apimService = apimService;
    }

    public async executeFunc(ctx: PluginContext): Promise<OptionItem[]> {
        const apimConfig = new ApimPluginConfig(ctx.config);
        const solutionConfig = new SolutionConfig(ctx.configOfOtherPlugins);
        const answer = buildAnswer(ctx);
        const resourceGroupName = apimConfig.resourceGroupName ?? solutionConfig.resourceGroupName;
        const serviceName = AssertConfigNotEmpty(TeamsToolkitComponent.ApimPlugin, ApimPluginConfigKeys.serviceName, apimConfig.serviceName);
        const apiPrefix =
            answer.apiPrefix ?? AssertConfigNotEmpty(TeamsToolkitComponent.ApimPlugin, ApimPluginConfigKeys.apiPrefix, apimConfig.apiPrefix);
        const versionSetId = apimConfig.versionSetId ?? NameSanitizer.sanitizeVersionSetId(apiPrefix, solutionConfig.resourceNameSuffix);

        const apiContracts = await this.apimService.listApi(resourceGroupName, serviceName, versionSetId);

        const existingApiVersionOptions: OptionItem[] = apiContracts.map((api) => {
            return { label: api.apiVersion, description: api.name, detail: api.displayName, data: api } as OptionItem;
        });
        const createNewApiVersionOption: OptionItem = { id: QuestionConstants.ApiVersion.createNewApiVersionOption, label: QuestionConstants.ApiVersion.createNewApiVersionOption };
        return [createNewApiVersionOption, ...existingApiVersionOptions];
    }

    public getQuestion(): SingleSelectQuestion {
        return {
            type: NodeType.singleSelect,
            name: QuestionConstants.ApiVersion.questionName,
            description: QuestionConstants.ApiVersion.description,
            option: {
                namespace: QuestionConstants.namespace,
                method: QuestionConstants.ApiVersion.funcName,
            },
            returnObject: true,
            skipSingleOption: false
        };
    }
}

export class NewApiVersionQuestion extends BaseQuestionService implements IQuestionService {
    public readonly funcName = QuestionConstants.NewApiVersion.funcName;

    constructor(dialog: Dialog, telemetry: Telemetry, logger?: LogProvider) {
        super(dialog, telemetry, logger);
    }

    public condition(): { target?: string; } & Validation {
        return {
            target: "$parent.id",
            equals: QuestionConstants.ApiVersion.createNewApiVersionOption,
        };
    }

    public async executeFunc(ctx: PluginContext): Promise<string> {
        const apiVersion = buildAnswer(ctx)?.openApiDocumentSpec?.info.version;
        return !!apiVersion ? NameSanitizer.sanitizeApiVersionIdentity(apiVersion) : ApimDefaultValues.apiVersion;
    }

    public getQuestion(): TextInputQuestion {
        return {
            type: NodeType.text,
            name: QuestionConstants.NewApiVersion.questionName,
            description: QuestionConstants.NewApiVersion.description,
            default: {
                namespace: QuestionConstants.namespace,
                method: QuestionConstants.NewApiVersion.funcName,
            },
        };
    }
}
