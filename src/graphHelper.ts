// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import 'isomorphic-fetch';
import {
  DeviceCodeCredential,
  DeviceCodePromptCallback,
} from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
// prettier-ignore
import { TokenCredentialAuthenticationProvider }
  from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { AppConfig } from './appConfig';

export function getGraphClientForUser(
  appConfig: AppConfig,
  deviceCodePrompt: DeviceCodePromptCallback
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
