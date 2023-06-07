// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  AuthorizationCodeCredential,
  ClientCertificateCredential,
  ClientSecretCredential,
  DeviceCodeCredential,
  InteractiveBrowserCredential,
  OnBehalfOfCredential,
  UsernamePasswordCredential,
} from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
// prettier-ignore
import { TokenCredentialAuthenticationProvider }
  from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
// prettier-ignore
import { AuthCodeMSALBrowserAuthenticationProvider }
  from '@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser';
import { PublicClientApplication, InteractionType } from '@azure/msal-browser';

export async function createWithMsalBrowser(): Promise<Client> {
  // <BrowserSnippet>
  // @azure/msal-browser
  const pca = new PublicClientApplication({
    auth: {
      clientId: 'YOUR_CLIENT_ID',
      authority: `https://login.microsoft.online/${'YOUR_TENANT_ID'}`,
      redirectUri: 'YOUR_REDIRECT_URI',
    },
  });

  // Authenticate to get the user's account
  const authResult = await pca.acquireTokenPopup({
    scopes: ['User.Read'],
  });

  if (!authResult.account) {
    throw new Error('Could not authenticate');
  }

  // @microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser
  const authProvider = new AuthCodeMSALBrowserAuthenticationProvider(pca, {
    account: authResult.account,
    interactionType: InteractionType.Popup,
    scopes: ['User.Read'],
  });

  const graphClient = Client.initWithMiddleware({ authProvider: authProvider });
  // </BrowserSnippet>

  return graphClient;
}

export function createWithAuthorizationCode(): Client {
  // <AuthorizationCodeSnippet>
  // @azure/identity
  const credential = new AuthorizationCodeCredential(
    'YOUR_TENANT_ID',
    'YOUR_CLIENT_ID',
    'YOUR_CLIENT_SECRET',
    'AUTHORIZATION_CODE',
    'REDIRECT_URL'
  );

  // @microsoft/microsoft-graph-client/authProviders/azureTokenCredentials
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['User.Read'],
  });

  const graphClient = Client.initWithMiddleware({ authProvider: authProvider });
  // </AuthorizationCodeSnippet>

  return graphClient;
}

export function createWithClientSecret(): Client {
  // <ClientSecretSnippet>
  // @azure/identity
  const credential = new ClientSecretCredential(
    'YOUR_TENANT_ID',
    'YOUR_CLIENT_ID',
    'YOUR_CLIENT_SECRET'
  );

  // @microsoft/microsoft-graph-client/authProviders/azureTokenCredentials
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    // The client credentials flow requires that you request the
    // /.default scope, and pre-configure your permissions on the
    // app registration in Azure. An administrator must grant consent
    // to those permissions beforehand.
    scopes: ['https://graph.microsoft.com/.default'],
  });

  const graphClient = Client.initWithMiddleware({ authProvider: authProvider });
  // </ClientSecretSnippet>

  return graphClient;
}

export function createWithClientCertificate(): Client {
  // <ClientCertificateSnippet>
  // @azure/identity
  const credential = new ClientCertificateCredential(
    'YOUR_TENANT_ID',
    'YOUR_CLIENT_ID',
    'YOUR_CERTIFICATE_PATH'
  );

  // @microsoft/microsoft-graph-client/authProviders/azureTokenCredentials
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    // The client credentials flow requires that you request the
    // /.default scope, and pre-configure your permissions on the
    // app registration in Azure. An administrator must grant consent
    // to those permissions beforehand.
    scopes: ['https://graph.microsoft.com/.default'],
  });

  const graphClient = Client.initWithMiddleware({ authProvider: authProvider });
  // </ClientCertificateSnippet>

  return graphClient;
}

export function createWithOnBehalfOf(): Client {
  // <OnBehalfOfSnippet>
  // @azure/identity
  const credential = new OnBehalfOfCredential({
    tenantId: 'YOUR_TENANT_ID',
    clientId: 'YOUR_CLIENT_ID',
    clientSecret: 'YOUR_CLIENT_SECRET',
    userAssertionToken: 'JWT_TOKEN_TO_EXCHANGE',
  });

  // @microsoft/microsoft-graph-client/authProviders/azureTokenCredentials
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default'],
  });

  const graphClient = Client.initWithMiddleware({ authProvider: authProvider });
  // </OnBehalfOfSnippet>

  return graphClient;
}

export function createWithDeviceCode(): Client {
  // <DeviceCodeSnippet>
  // @azure/identity
  const credential = new DeviceCodeCredential({
    tenantId: 'YOUR_TENANT_ID',
    clientId: 'YOUR_CLIENT_ID',
    userPromptCallback: (info) => {
      console.log(info.message);
    },
  });

  // @microsoft/microsoft-graph-client/authProviders/azureTokenCredentials
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['User.Read'],
  });

  const graphClient = Client.initWithMiddleware({ authProvider: authProvider });
  // </DeviceCodeSnippet>

  return graphClient;
}

export function createWithInteractive(): Client {
  // <InteractiveSnippet>
  // @azure/identity
  const credential = new InteractiveBrowserCredential({
    tenantId: 'YOUR_TENANT_ID',
    clientId: 'YOUR_CLIENT_ID',
    redirectUri: 'http://localhost',
  });

  // @microsoft/microsoft-graph-client/authProviders/azureTokenCredentials
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['User.Read'],
  });

  const graphClient = Client.initWithMiddleware({ authProvider: authProvider });
  // </InteractiveSnippet>

  return graphClient;
}

export function createWithUserNamePassword(): Client {
  // <UserNamePasswordSnippet>
  // @azure/identity
  const credential = new UsernamePasswordCredential(
    'YOUR_TENANT_ID',
    'YOUR_CLIENT_ID',
    'YOUR_USER_NAME',
    'YOUR_PASSWORD'
  );

  // @microsoft/microsoft-graph-client/authProviders/azureTokenCredentials
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['User.Read'],
  });

  const graphClient = Client.initWithMiddleware({ authProvider: authProvider });
  // </UserNamePasswordSnippet>

  return graphClient;
}
