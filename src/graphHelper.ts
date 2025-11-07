// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  DeviceCodeCredential,
  DeviceCodePromptCallback,
} from '@azure/identity';
import { AzureIdentityAuthenticationProvider } from '@microsoft/kiota-authentication-azure';
import {
  createGraphServiceClient,
  GraphRequestAdapter,
  GraphServiceClient,
} from '@microsoft/msgraph-sdk';
import '@microsoft/msgraph-sdk-drives';
import '@microsoft/msgraph-sdk-groups';
import '@microsoft/msgraph-sdk-teams';
import '@microsoft/msgraph-sdk-users';
import { AppConfig } from './appConfig.js';

export function getGraphClientForUser(
  appConfig: AppConfig,
  deviceCodePrompt: DeviceCodePromptCallback,
): GraphServiceClient {
  const credential = new DeviceCodeCredential({
    clientId: appConfig.clientId,
    tenantId: appConfig.tenantId,
    userPromptCallback: deviceCodePrompt,
  });

  const authProvider = new AzureIdentityAuthenticationProvider(
    credential,
    appConfig.graphUserScopes,
  );

  const requestAdapter = new GraphRequestAdapter(authProvider);

  return createGraphServiceClient(requestAdapter);
}

/*
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
*/
