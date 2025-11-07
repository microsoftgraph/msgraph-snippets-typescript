// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ClientSecretCredential, TokenCredential } from '@azure/identity';
import { AzureIdentityAuthenticationProvider } from '@microsoft/kiota-authentication-azure';
import {
  createGraphServiceClient,
  GraphRequestAdapter,
  GraphServiceClient,
} from '@microsoft/msgraph-sdk';
import { getDefaultMiddlewares } from '@microsoft/msgraph-sdk-core';
import {
  ChaosHandler,
  KiotaClientFactory,
} from '@microsoft/kiota-http-fetchlibrary';
import {
  fetch,
  ProxyAgent,
  RequestInit as UndiciRequestInit,
} from 'undici-types';

export function createWithChaosHandler(
  credential: TokenCredential,
  scopes: string[],
): GraphServiceClient {
  // <ChaosHandlerSnippet>
  // credential is one of the credential classes from @azure/identity
  // scopes is an array of permission scope strings
  const authProvider = new AzureIdentityAuthenticationProvider(
    credential,
    scopes,
  );

  // Create a chaos handler (from @microsoft/kiota-http-fetchlibrary)
  const chaosHandler = new ChaosHandler();

  // Get the default middleware stack
  const middlewares = getDefaultMiddlewares();

  // Add the chaos handler to the middleware stack
  middlewares.push(chaosHandler);

  // Create an HttpClient with middlewares
  var httpClient = KiotaClientFactory.create(undefined, middlewares);

  // Create request adapter with the HttpClient
  var requestAdapter = new GraphRequestAdapter(
    authProvider,
    undefined,
    undefined,
    httpClient,
  );

  const graphClient = createGraphServiceClient(requestAdapter);
  // </ChaosHandlerSnippet>

  return graphClient;
}

export function createWithProxy(scopes: string[]): GraphServiceClient {
  // <ProxySnippet>
  // Setup proxy for the token credential from @azure/identity
  const credential = new ClientSecretCredential(
    'YOUR_TENANT_ID',
    'YOUR_CLIENT_ID',
    'YOUR_CLIENT_SECRET',
    {
      proxyOptions: {
        host: 'localhost',
        port: 8888,
      },
    },
  );

  // scopes is an array of permission scope strings
  const authProvider = new AzureIdentityAuthenticationProvider(
    credential,
    scopes,
  );

  // Create a new HTTPS proxy agent (from https-proxy-agent)
  const proxyAgent = new ProxyAgent('http://localhost:8888');

  const customFetch = async (request: string, init: RequestInit) => {
    // import { RequestInit as UndiciRequestInit } from 'undici-types'
    const requestInit: UndiciRequestInit = {
      ...(init as UndiciRequestInit),
      dispatcher: proxyAgent,
    };
    const response = await fetch(request, requestInit);
    return response as unknown as Response;
  };

  // Create an HttpClient with custom fetch callback
  var httpClient = KiotaClientFactory.create(customFetch);

  // Create request adapter with the HttpClient
  var requestAdapter = new GraphRequestAdapter(
    authProvider,
    undefined,
    undefined,
    httpClient,
  );

  const graphClient = createGraphServiceClient(requestAdapter);
  // </ProxySnippet>

  return graphClient;
}
