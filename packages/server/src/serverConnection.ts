// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CancellationToken, MessageConnection } from "vscode-jsonrpc";
import {
  FxError,
  Inputs,
  Void,
  Tools,
  Result,
  Func,
  ok,
  Stage,
  QTreeNode,
} from "@microsoft/teamsfx-api";
import { FxCore, Correlator, getSideloadingStatus } from "@microsoft/teamsfx-core";
import { IServerConnection, Namespaces } from "./apis";
import LogProvider from "./providers/logger";
import TokenProvider from "./providers/tokenProvider";
import TelemetryReporter from "./providers/telemetry";
import UserInteraction from "./providers/userInteraction";
import { callFunc } from "./customizedFuncAdapter";
import { standardizeResult } from "./utils";
import { environmentManager } from "@microsoft/teamsfx-core";

export default class ServerConnection implements IServerConnection {
  public static readonly namespace = Namespaces.Server;
  private readonly connection: MessageConnection;
  private readonly tools: Tools;
  private readonly core: FxCore;

  constructor(connection: MessageConnection) {
    this.connection = connection;
    this.tools = {
      logProvider: new LogProvider(connection),
      tokenProvider: new TokenProvider(connection),
      telemetryReporter: new TelemetryReporter(connection),
      ui: new UserInteraction(connection),
    };
    this.core = new FxCore(this.tools);

    [
      this.getQuestionsRequest.bind(this),
      this.createProjectRequest.bind(this),
      this.localDebugRequest.bind(this),
      this.provisionResourcesRequest.bind(this),
      this.deployArtifactsRequest.bind(this),
      this.buildArtifactsRequest.bind(this),
      this.publishApplicationRequest.bind(this),
      this.deployTeamsAppManifestRequest.bind(this),
      this.getSideloadingStatusRequest.bind(this),

      this.customizeLocalFuncRequest.bind(this),
      this.customizeValidateFuncRequest.bind(this),
      this.customizeOnSelectionChangeFuncRequest.bind(this),
      this.addSsoRequest.bind(this),
    ].forEach((fn) => {
      /// fn.name = `bound ${functionName}`
      connection.onRequest(`${ServerConnection.namespace}/${fn.name.split(" ")[1]}`, fn);
    });
  }

  public listen() {
    this.connection.listen();
  }

  public async getQuestionsRequest(
    stage: Stage,
    inputs: Inputs,
    token: CancellationToken
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    const corrId = inputs.correlationId ? inputs.correlationId : "";
    const res = await Correlator.runWithId(
      corrId,
      (stage, inputs) => this.core.getQuestions(stage, inputs),
      stage,
      inputs
    );
    return standardizeResult(res);
  }

  public async createProjectRequest(
    inputs: Inputs,
    token: CancellationToken
  ): Promise<Result<string, FxError>> {
    const corrId = inputs.correlationId ? inputs.correlationId : "";
    const res = await Correlator.runWithId(
      corrId,
      (params) => this.core.createProject(params),
      inputs
    );
    return standardizeResult(res);
  }

  public async localDebugRequest(
    inputs: Inputs,
    token: CancellationToken
  ): Promise<Result<Void, FxError>> {
    const corrId = inputs.correlationId ? inputs.correlationId : "";
    const res = await Correlator.runWithId(
      corrId,
      (params) => this.core.localDebug(params),
      inputs
    );
    return standardizeResult(res);
  }

  public async provisionResourcesRequest(
    inputs: Inputs,
    token: CancellationToken
  ): Promise<Result<Void, FxError>> {
    const corrId = inputs.correlationId ? inputs.correlationId : "";
    const res = await Correlator.runWithId(
      corrId,
      (params) => this.core.provisionResources(params),
      inputs
    );
    return standardizeResult(res);
  }

  public async deployArtifactsRequest(
    inputs: Inputs,
    token: CancellationToken
  ): Promise<Result<Void, FxError>> {
    const corrId = inputs.correlationId ? inputs.correlationId : "";
    const res = await Correlator.runWithId(
      corrId,
      (params) => this.core.deployArtifacts(params),
      inputs
    );
    return standardizeResult(res);
  }

  public async buildArtifactsRequest(
    inputs: Inputs,
    token: CancellationToken
  ): Promise<Result<any, FxError>> {
    const corrId = inputs.correlationId ? inputs.correlationId : "";
    const func: Func = {
      namespace: "fx-solution-azure",
      method: "buildPackage",
      params: {
        type: inputs.env == environmentManager.getLocalEnvName() ? "localDebug" : "remote",
        env: inputs.env,
      },
    };
    const res = await Correlator.runWithId(
      corrId,
      (func, inputs) => this.core.executeUserTask(func, inputs),
      func,
      inputs
    );
    return standardizeResult(res);
  }

  public async publishApplicationRequest(
    inputs: Inputs,
    token: CancellationToken
  ): Promise<Result<Void, FxError>> {
    const corrId = inputs.correlationId ? inputs.correlationId : "";
    const res = await Correlator.runWithId(
      corrId,
      (params) => this.core.publishApplication(params),
      inputs
    );
    return standardizeResult(res);
  }

  public async deployTeamsAppManifestRequest(
    inputs: Inputs,
    token: CancellationToken
  ): Promise<Result<any, FxError>> {
    const corrId = inputs.correlationId ? inputs.correlationId : "";
    const func: Func = {
      namespace: "fx-solution-azure/fx-resource-appstudio",
      method: "updateManifest",
      params: {
        envName: environmentManager.getDefaultEnvName(),
      },
    };
    const res = await Correlator.runWithId(
      corrId,
      (func, inputs) => this.core.executeUserTask(func, inputs),
      func,
      inputs
    );
    return standardizeResult(res);
  }

  public async addSsoRequest(
    inputs: Inputs,
    token: CancellationToken
  ): Promise<Result<any, FxError>> {
    const corrId = inputs.correlationId ? inputs.correlationId : "";
    const func: Func = {
      namespace: "fx-solution-azure",
      method: "addSso",
      params: {
        envName: environmentManager.getDefaultEnvName(),
      },
    };
    const res = await Correlator.runWithId(
      corrId,
      (func, inputs) => this.core.executeUserTask(func, inputs),
      func,
      inputs
    );
    return standardizeResult(res);
  }

  public async customizeLocalFuncRequest(
    funcId: number,
    params: Inputs,
    token: CancellationToken
  ): Promise<Result<Void, FxError>> {
    const res = await callFunc("LocalFunc", funcId, params);
    return standardizeResult(res);
  }

  public async customizeValidateFuncRequest(
    funcId: number,
    answer: any,
    previousAnswers: Inputs | undefined,
    token: CancellationToken
  ): Promise<Result<any, FxError>> {
    const res = await callFunc("ValidateFunc", funcId, answer, previousAnswers);
    return standardizeResult(res);
  }

  public async customizeOnSelectionChangeFuncRequest(
    funcId: number,
    currentSelectedIds: Set<string>,
    previousSelectedIds: Set<string>,
    token: CancellationToken
  ): Promise<Result<any, FxError>> {
    const res = await callFunc(
      "OnSelectionChangeFunc",
      funcId,
      currentSelectedIds,
      previousSelectedIds
    );
    return standardizeResult(res);
  }

  public async getSideloadingStatusRequest(
    accountToken: {
      token: string;
    },
    token: CancellationToken
  ): Promise<Result<string, FxError>> {
    const res = await getSideloadingStatus(accountToken.token);
    return ok(String(res));
  }
}
