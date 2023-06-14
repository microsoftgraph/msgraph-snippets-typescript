// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  AzureAuthorityHosts,
  InteractiveBrowserCredential,
} from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';

export function createClientForUSGov(): Client {
  // <NationalCloudSnippet>
  // Create the InteractiveBrowserCredential using details
  // from app registered in the Azure AD for US Government portal
  const credential = new InteractiveBrowserCredential({
    clientId: 'YOUR_CLIENT_ID',
    tenantId: 'YOUR_TENANT_ID',
    // https://login.microsoftonline.us
    authorityHost: AzureAuthorityHosts.AzureGovernment,
    redirectUri: 'YOUR_REDIRECT_URI',
  });

  // Create the authentication provider
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.us/.default'],
  });

  // Create the Microsoft Graph client object using
  // the Microsoft Graph for US Government L4 endpoint
  // NOTE: Do not include the version in the baseUrl
  const graphClient = Client.initWithMiddleware({
    authProvider: authProvider,
    baseUrl: 'https://graph.microsoft.us',
  });
  // </NationalCloudSnippet>

  return graphClient;
}
