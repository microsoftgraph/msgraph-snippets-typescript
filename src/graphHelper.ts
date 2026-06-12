// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import 'isomorphic-fetch';
import {
  DeviceCodeCredential,
  DeviceCodePromptCallback,
} from '@azure/identity';
import {
  AuthenticationHandler,
  Client,
  HTTPMessageHandler,
} from '@microsoft/microsoft-graph-client';
// prettier-ignore
import { TokenCredentialAuthenticationProvider }
  from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import { AppConfig } from './appConfig.js';
import ClientLoggingMiddleware from './clientLoggingMiddleware.js';

export function getGraphClientForUser(
  appConfig: AppConfig,
  deviceCodePrompt: DeviceCodePromptCallback,
): Client {
  const credential = new DeviceCodeCredential({
    clientId: appConfig.clientId,
    tenantId: appConfig.tenantId,
    userPromptCallback: deviceCodePrompt,
  });

  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: appConfig.graphUserScopes,
  });

  return Client.initWithMiddleware({ authProvider: authProvider });
}

export function getDebugGraphClientForUser(
  appConfig: AppConfig,
  deviceCodePrompt: DeviceCodePromptCallback,
): Client {
  const credential = new DeviceCodeCredential({
    clientId: appConfig.clientId,
    tenantId: appConfig.tenantId,
    userPromptCallback: deviceCodePrompt,
  });

  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: appConfig.graphUserScopes,
  });

  const authHandler = new AuthenticationHandler(authProvider);
  const loggingHandler = new ClientLoggingMiddleware();
  const httpHandler = new HTTPMessageHandler();

  authHandler.setNext(loggingHandler);
  loggingHandler.setNext(httpHandler);

  return Client.initWithMiddleware({
    middleware: authHandler,
  });
}
