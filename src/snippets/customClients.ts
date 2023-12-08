// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TokenCredential } from '@azure/identity';
import {
  Client,
  AuthenticationHandler,
  ChaosHandler,
  HTTPMessageHandler,
} from '@microsoft/microsoft-graph-client';
// prettier-ignore
import { TokenCredentialAuthenticationProvider }
  from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { HttpsProxyAgent } from 'https-proxy-agent';

export function createWithChaosHandler(
  credential: TokenCredential,
  scopes: string[],
): Client {
  // <ChaosHandlerSnippet>
  // credential is one of the credential classes from @azure/identity
  // scopes is an array of permission scope strings
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: scopes,
  });

  // Create an authentication handler (from @microsoft/microsoft-graph-client)
  const authHandler = new AuthenticationHandler(authProvider);

  // Create a chaos handler (from @microsoft/microsoft-graph-client)
  const chaosHandler = new ChaosHandler();

  // Create a standard HTTP handler (from @microsoft/microsoft-graph-client)
  const httpHandler = new HTTPMessageHandler();

  // Use setNext to chain handlers together
  // auth -> chaos -> http
  authHandler.setNext(chaosHandler);
  chaosHandler.setNext(httpHandler);

  // Pass the first middleware in the chain in the middleWare property
  const graphClient = Client.initWithMiddleware({
    middleware: authHandler,
  });
  // </ChaosHandlerSnippet>

  return graphClient;
}

export function createWithProxy(
  credential: TokenCredential,
  scopes: string[],
): Client {
  // <ProxySnippet>
  // credential is one of the credential classes from @azure/identity
  // scopes is an array of permission scope strings
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: scopes,
  });

  // Create a new HTTPS proxy agent (from https-proxy-agent)
  const proxyAgent = new HttpsProxyAgent('http://localhost:8888');

  // Create a client with the proxy
  const graphClient = Client.initWithMiddleware({
    authProvider: authProvider,
    fetchOptions: {
      agent: proxyAgent,
    },
  });
  // </ProxySnippet>

  return graphClient;
}
