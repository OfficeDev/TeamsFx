// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import "mocha";
import {
  match,
  SinonMatcher,
  SinonSandbox,
  SinonStub,
  SinonStubbedInstance,
  SinonStubbedMember,
  StubbableType,
} from "sinon";
import { ApimService } from "../../../../src/plugins/resource/apim/services/apimService";
import {
  Api,
  ApiManagementClient,
  ApiManagementService,
  ApiVersionSet,
  ProductApi,
} from "@azure/arm-apimanagement";
import { Providers } from "@azure/arm-resources";
import { TokenCredentialsBase } from "@azure/ms-rest-nodeauth";
import {
  ApiCreateOrUpdateParameter,
  ApiManagementServiceResource,
  ApiVersionSetContract,
} from "@azure/arm-apimanagement/src/models";
import axios, { AxiosInstance } from "axios";
import { IAadInfo } from "../../../../src/plugins/resource/apim/interfaces/IAadResource";

export type StubbedClass<T> = SinonStubbedInstance<T> & T;

export function createSinonStubInstance<T>(
  sandbox: SinonSandbox,
  constructor: StubbableType<T>,
  overrides?: { [K in keyof T]?: SinonStubbedMember<T[K]> }
): StubbedClass<T> {
  const stub = sandbox.createStubInstance<T>(constructor, overrides);
  return stub as unknown as StubbedClass<T>;
}

export const DefaultTestInput = {
  subscriptionId: "test-subscription-id",
  resourceGroup: {
    existing: "test-resource-group",
    new: "test-resource-group-new",
  },
  apimServiceName: {
    existing: "test-service",
    new: "test-service-new",
    error: "test-service-error",
  },
  versionSet: {
    existing: "test-version-set",
    new: "test-version-set-new",
    error: "test-version-set-error",
  },
  apiId: {
    existing: "test-api-id",
    new: "test-api-id-new",
    error: "test-api-id-error",
  },
  productId: {
    existing: "test-product-id",
    new: "test-product-id-new",
    error: "test-product-id-error",
  },
  aadDisplayName: {
    new: "test-aad-display-name-new",
    error: "test-aad-display-name-error",
  },
  aadObjectId: {
    created: "test-aad-object-id-created",
    new: "test-aad-object-id-new",
    error: "test-aad-object-id-error",
  },
  aadSecretDisplayName: {
    new: "test-aad-secret-display-name-new",
    error: "test-aad-secret-display-name-error",
  },
  aadClientId: {
    new: "test-aad-client-id-new",
    created: "test-aad-client-id-created",
    error: "test-aad-client-id-error",
  },
};

export const DefaultTestOutput = {
  createAad: {
    id: "test-aad-object-id-created",
    appId: "test-aad-client-id-created",
  },
  addSecret: {
    secretText: "test-secret-text",
  },
  getAad: {
    id: "test-aad-object-id-created",
    appId: "test-aad-client-id-created",
    displayName: "test-aad-display-name-created",
    requiredResourceAccess: [],
    web: {
      redirectUris: [],
      implicitGrantSettings: { enableIdTokenIssuance: false },
    },
  },
};

export function mockApimService(sandbox: SinonSandbox): {
  apimService: ApimService;
  apiManagementClient: StubbedClass<ApiManagementClient>;
  credential: StubbedClass<MockTokenCredentials>;
} {
  const apiManagementClient = createSinonStubInstance(sandbox, ApiManagementClient);
  const resourceProviderClient = createSinonStubInstance(sandbox, Providers);
  const credential = createSinonStubInstance(sandbox, MockTokenCredentials);
  const apimService = new ApimService(
    apiManagementClient,
    resourceProviderClient,
    credential,
    DefaultTestInput.subscriptionId
  );

  return { apimService, apiManagementClient, credential };
}

export function mockApiManagementService(sandbox: SinonSandbox): any {
  const apiManagementServiceStub = sandbox.createStubInstance(ApiManagementService);
  const getStub = apiManagementServiceStub.get as unknown as sinon.SinonStub<
    [string, string],
    Promise<any>
  >;
  getStub.rejects(UnexpectedInputError);
  getStub
    .withArgs(DefaultTestInput.resourceGroup.existing, DefaultTestInput.apimServiceName.new)
    .rejects(
      buildError({
        code: "ResourceNotFound",
        statusCode: 404,
        message: `The Resource 'Microsoft.ApiManagement/service/${DefaultTestInput.apimServiceName.new}' under resource group '${DefaultTestInput.resourceGroup.existing}' was not found. For more details please go to https://aka.ms/ARMResourceNotFoundFix`,
      })
    );
  getStub.withArgs(DefaultTestInput.resourceGroup.new, match.any).rejects(
    buildError({
      code: "ResourceGroupNotFound",
      statusCode: 404,
      message: `Resource group '${DefaultTestInput.resourceGroup.new}' could not be found.`,
    })
  );
  getStub
    .withArgs(DefaultTestInput.resourceGroup.existing, DefaultTestInput.apimServiceName.existing)
    .resolves({});

  const createOrUpdateStub = apiManagementServiceStub.createOrUpdate as unknown as SinonStub<
    [string, string, ApiManagementServiceResource],
    Promise<any>
  >;
  createOrUpdateStub.rejects(UnexpectedInputError);
  createOrUpdateStub.withArgs(match.any, DefaultTestInput.apimServiceName.error, match.any).rejects(
    buildError({
      code: "TestError",
      statusCode: 400,
      message: "Mock test error",
    })
  );
  createOrUpdateStub.withArgs(match.any, match.any, match.any).resolves({});

  return apiManagementServiceStub;
}

export function mockApiVersionSet(sandbox: SinonSandbox): any {
  const apiVersionSet = sandbox.createStubInstance(ApiVersionSet);
  const createOrUpdateStub = apiVersionSet.createOrUpdate as unknown as SinonStub<
    [string, string, string, ApiVersionSetContract],
    Promise<any>
  >;
  createOrUpdateStub.rejects(UnexpectedInputError);
  createOrUpdateStub.withArgs(match.any, match.any, match.any, match.any).resolves({});

  const getStub = apiVersionSet.get as unknown as SinonStub<[string, string, string], Promise<any>>;
  getStub.rejects(UnexpectedInputError);
  getStub.withArgs(match.any, match.any, DefaultTestInput.versionSet.new).rejects(
    buildError({
      code: "ResourceNotFound",
      statusCode: 404,
      message: `The version set '${DefaultTestInput.versionSet.new}' was not found.`,
    })
  );
  getStub.withArgs(match.any, match.any, DefaultTestInput.versionSet.existing).resolves({});

  return apiVersionSet;
}

export function mockApi(sandbox: SinonSandbox): any {
  const apiStub = sandbox.createStubInstance(Api);
  const createOrUpdateStub = apiStub.createOrUpdate as unknown as SinonStub<
    [string, string, string, ApiCreateOrUpdateParameter],
    Promise<any>
  >;
  createOrUpdateStub.rejects(UnexpectedInputError);
  createOrUpdateStub
    .withArgs(match.any, match.any, DefaultTestInput.apiId.error, match.any)
    .rejects(
      buildError({
        code: "TestError",
        statusCode: 400,
        message: "Mock test error",
      })
    );
  createOrUpdateStub.withArgs(match.any, match.any, match.any, match.any).resolves({});
  return apiStub;
}

export function mockProductApi(sandbox: SinonSandbox): any {
  const productApi = sandbox.createStubInstance(ProductApi);

  // Mock productApi.createOrUpdate
  const productApiStub = productApi.createOrUpdate as unknown as SinonStub<
    [string, string, string, string],
    Promise<any>
  >;
  productApiStub.rejects(UnexpectedInputError);
  // createOrUpdate (failed)
  productApiStub
    .withArgs(match.any, match.any, DefaultTestInput.productId.error, DefaultTestInput.apiId.error)
    .rejects(
      buildError({
        code: "TestError",
        statusCode: 400,
        message: "Mock test error",
      })
    );
  // createOrUpdate (success)
  productApiStub.withArgs(match.any, match.any, match.any, match.any).resolves({});

  // Mock productApi.checkEntityExists
  const checkEntityExistsStub = productApi.checkEntityExists as unknown as SinonStub<
    [string, string, string, string],
    Promise<any>
  >;
  checkEntityExistsStub.rejects(UnexpectedInputError);
  checkEntityExistsStub
    .withArgs(match.any, match.any, DefaultTestInput.productId.new, DefaultTestInput.apiId.new)
    .rejects(
      buildError({
        code: "ResourceNotFound",
        statusCode: 404,
        message: `The product api '${DefaultTestInput.versionSet.new}' was not found.`,
      })
    );
  checkEntityExistsStub
    .withArgs(
      match.any,
      match.any,
      DefaultTestInput.productId.existing,
      DefaultTestInput.apiId.existing
    )
    .resolves({});

  return productApi;
}

export class MockTokenCredentials extends TokenCredentialsBase {
  public async getToken(): Promise<any> {
    return undefined;
  }
}

export function mockCredential(
  sandbox: SinonSandbox,
  credential: StubbedClass<MockTokenCredentials>,
  token: any
): void {
  credential.getToken = sandbox.stub<[], Promise<any>>().resolves(token);
}

export type MockAxiosInput = {
  aadDisplayName?: { error?: string };
  aadObjectId?: { created?: string };
  aadClientId?: { created?: string };
};

export type MockAxiosOutput = {
  createAad?: {
    id: string;
    appId: string;
  };
  addSecret?: {
    secretText: string;
  };
  getAad?: IAadInfo;
};

export function mockAxios(
  sandbox: SinonSandbox,
  mockInput: MockAxiosInput = DefaultTestInput,
  mockOutput: MockAxiosOutput = DefaultTestOutput
): {
  axiosInstance: AxiosInstance;
  requestStub: any;
} {
  const mockAxiosInstance: any = axios.create();
  const requestStub = sandbox.stub(mockAxiosInstance, "request").rejects(UnexpectedInputError);

  // Create AAD (success)
  requestStub
    .withArgs(aadMatcher.createAad.and(match.has("data")))
    .resolves(buildAxiosResponse(mockOutput.createAad ?? DefaultTestOutput.createAad));

  // Create AAD (failed)
  if (mockInput?.aadDisplayName?.error) {
    requestStub
      .withArgs(
        aadMatcher.createAad.and(match.has("data", { displayName: mockInput.aadDisplayName.error }))
      )
      .rejects(buildError({ message: "error" }));
  }

  // Add secret
  requestStub
    .withArgs(aadMatcher.addSecret)
    .resolves(buildAxiosResponse(mockOutput.addSecret ?? DefaultTestOutput.addSecret));
  // Update AAD
  requestStub.withArgs(aadMatcher.updateAad).resolves(buildAxiosResponse({}));

  // Get AAD (not found)
  requestStub.withArgs(aadMatcher.getAad).resolves({});

  // Get AAD (existing)
  if (mockInput?.aadObjectId?.created) {
    requestStub
      .withArgs(
        aadMatcher.getAad.and(match.has("url", `/applications/${mockInput.aadObjectId.created}`))
      )
      .resolves(buildAxiosResponse(mockOutput?.getAad ?? DefaultTestOutput.getAad));
  }

  // Get ServicePrincipal (not found)
  requestStub.withArgs(aadMatcher.getServicePrincipals).resolves(
    buildAxiosResponse({
      value: [],
    })
  );

  // Get ServicePrincipal (existing)
  if (mockInput?.aadClientId?.created) {
    requestStub
      .withArgs(
        aadMatcher.getServicePrincipals.and(
          match.has("url", `/servicePrincipals?$filter=appId eq '${mockInput.aadClientId.created}'`)
        )
      )
      .resolves(
        buildAxiosResponse({
          value: [{}],
        })
      );
  }

  // Create ServicePrincipal
  requestStub.withArgs(aadMatcher.createServicePrincipal).resolves(buildAxiosResponse({}));

  mockAxiosInstance.request = requestStub;
  return { axiosInstance: mockAxiosInstance, requestStub: requestStub };
}

export const aadMatcher = {
  createAad: match.has("method", "post").and(match.has("url", "/applications")),
  addSecret: match
    .has("method", "post")
    .and(urlMatcher(["applications", undefined, "addPassword"])),
  updateAad: match.has("method", "patch").and(urlMatcher(["applications", undefined])),
  getAad: match.has("method", "get").and(urlMatcher(["applications", undefined])),
  getServicePrincipals: match
    .has("method", "get")
    .and(urlMatcher(["servicePrincipals", undefined])),
  createServicePrincipal: match.has("method", "post").and(urlMatcher(["servicePrincipals"])),
  body: (matcher: SinonMatcher | any) => match.has("data", matcher),
};

function urlMatcher(urls: (string | undefined)[]): SinonMatcher {
  return match.has(
    "url",
    match.string.and(
      match((value: string) => {
        const res = value.split(/\/|\?/);
        if (res.length < urls.length + 1) {
          return false;
        }
        for (let i = 0; i < urls.length; ++i) {
          if (urls[i] && res[i + 1] !== urls[i]) {
            return false;
          }
        }
        return true;
      })
    )
  );
}

function buildAxiosResponse(obj: any): any {
  return { data: obj };
}

function buildError(obj: any): Error {
  const error = new Error();
  return Object.assign(error, obj);
}

const UnexpectedInputError = new Error("Unexpected input");
