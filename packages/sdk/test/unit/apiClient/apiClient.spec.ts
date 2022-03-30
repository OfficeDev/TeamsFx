// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import * as sinon from "sinon";
import { MockAuthProvider } from "./mockAuthProvider";
import axios, { AxiosRequestConfig } from "axios";
import { createApiClient } from "../../../src";
import MockAdapter from "axios-mock-adapter";

describe("ApiClient Tests - Node", () => {
  it("createApiClient should return axios instance with base url and config updated", async function () {
    // Arrange
    const apiBaseUrl = "https://fake-api-endpoint";
    const mockAuthProvider = new MockAuthProvider();
    sinon
      .stub(mockAuthProvider, "AddAuthenticationInfo")
      .callsFake(async (config: AxiosRequestConfig): Promise<AxiosRequestConfig> => {
        config.headers = {
          Authorization: "fake-token",
        };
        return config;
      });
    const axiosMockAdapter = new MockAdapter(axios);
    axiosMockAdapter.onGet("/foo").replyOnce(200);

    // Act
    const apiClient = createApiClient(apiBaseUrl, mockAuthProvider);
    const res = await apiClient.get("/foo");

    // Assert
    assert.equal(res.config.baseURL, apiBaseUrl);
    assert.equal(res.config.url, "/foo");
    assert.equal(res.config.headers!["Authorization"], "fake-token");
  });
});
