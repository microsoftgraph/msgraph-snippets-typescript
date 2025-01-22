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
import {
  createGraphServiceClient,
  GraphRequestAdapter,
  GraphServiceClient,
} from '@microsoft/msgraph-sdk';
import { AzureIdentityAuthenticationProvider } from '@microsoft/kiota-authentication-azure';

export function createWithAuthorizationCode(): GraphServiceClient {
  // <AuthorizationCodeSnippet>
  // @azure/identity
  const credential = new AuthorizationCodeCredential(
    'YOUR_TENANT_ID',
    'YOUR_CLIENT_ID',
    'YOUR_CLIENT_SECRET',
    'AUTHORIZATION_CODE',
    'REDIRECT_URL',
  );

  const scopes = ['User.Read'];

  // @microsoft/kiota-authentication-azure
  const authProvider = new AzureIdentityAuthenticationProvider(
    credential,
    scopes,
  );

  const requestAdapter = new GraphRequestAdapter(authProvider);

  const graphClient = createGraphServiceClient(requestAdapter);
  // </AuthorizationCodeSnippet>

  return graphClient;
}

export function createWithClientSecret(): GraphServiceClient {
  // <ClientSecretSnippet>
  // @azure/identity
  const credential = new ClientSecretCredential(
    'YOUR_TENANT_ID',
    'YOUR_CLIENT_ID',
    'YOUR_CLIENT_SECRET',
  );

  // The client credentials flow requires that you request the
  // /.default scope, and pre-configure your permissions on the
  // app registration in Azure. An administrator must grant consent
  // to those permissions beforehand.
  const scopes = ['https://graph.microsoft.com/.default'];

  // @microsoft/kiota-authentication-azure
  const authProvider = new AzureIdentityAuthenticationProvider(
    credential,
    scopes,
  );

  const requestAdapter = new GraphRequestAdapter(authProvider);

  const graphClient = createGraphServiceClient(requestAdapter);
  // </ClientSecretSnippet>

  return graphClient;
}

export function createWithClientCertificate(): GraphServiceClient {
  // <ClientCertificateSnippet>
  // @azure/identity
  const credential = new ClientCertificateCredential(
    'YOUR_TENANT_ID',
    'YOUR_CLIENT_ID',
    'YOUR_CERTIFICATE_PATH',
  );

  // The client credentials flow requires that you request the
  // /.default scope, and pre-configure your permissions on the
  // app registration in Azure. An administrator must grant consent
  // to those permissions beforehand.
  const scopes = ['https://graph.microsoft.com/.default'];

  // @microsoft/kiota-authentication-azure
  const authProvider = new AzureIdentityAuthenticationProvider(
    credential,
    scopes,
  );

  const requestAdapter = new GraphRequestAdapter(authProvider);

  const graphClient = createGraphServiceClient(requestAdapter);
  // </ClientCertificateSnippet>

  return graphClient;
}

export function createWithOnBehalfOf(): GraphServiceClient {
  // <OnBehalfOfSnippet>
  // @azure/identity
  const credential = new OnBehalfOfCredential({
    tenantId: 'YOUR_TENANT_ID',
    clientId: 'YOUR_CLIENT_ID',
    clientSecret: 'YOUR_CLIENT_SECRET',
    userAssertionToken: 'JWT_TOKEN_TO_EXCHANGE',
  });

  const scopes = ['https://graph.microsoft.com/.default'];

  // @microsoft/kiota-authentication-azure
  const authProvider = new AzureIdentityAuthenticationProvider(
    credential,
    scopes,
  );

  const requestAdapter = new GraphRequestAdapter(authProvider);

  const graphClient = createGraphServiceClient(requestAdapter);
  // </OnBehalfOfSnippet>

  return graphClient;
}

export function createWithDeviceCode(): GraphServiceClient {
  // <DeviceCodeSnippet>
  // @azure/identity
  const credential = new DeviceCodeCredential({
    tenantId: 'YOUR_TENANT_ID',
    clientId: 'YOUR_CLIENT_ID',
    userPromptCallback: (info) => {
      console.log(info.message);
    },
  });

  const scopes = ['User.Read'];

  // @microsoft/kiota-authentication-azure
  const authProvider = new AzureIdentityAuthenticationProvider(
    credential,
    scopes,
  );

  const requestAdapter = new GraphRequestAdapter(authProvider);

  const graphClient = createGraphServiceClient(requestAdapter);
  // </DeviceCodeSnippet>

  return graphClient;
}

export function createWithInteractive(): GraphServiceClient {
  // <InteractiveSnippet>
  // @azure/identity
  const credential = new InteractiveBrowserCredential({
    tenantId: 'YOUR_TENANT_ID',
    clientId: 'YOUR_CLIENT_ID',
    redirectUri: 'http://localhost',
  });

  const scopes = ['User.Read'];

  // @microsoft/kiota-authentication-azure
  const authProvider = new AzureIdentityAuthenticationProvider(
    credential,
    scopes,
  );

  const requestAdapter = new GraphRequestAdapter(authProvider);

  const graphClient = createGraphServiceClient(requestAdapter);
  // </InteractiveSnippet>

  return graphClient;
}

export function createWithUserNamePassword(): GraphServiceClient {
  // <UserNamePasswordSnippet>
  // @azure/identity
  const credential = new UsernamePasswordCredential(
    'YOUR_TENANT_ID',
    'YOUR_CLIENT_ID',
    'YOUR_USER_NAME',
    'YOUR_PASSWORD',
  );

  const scopes = ['User.Read'];

  // @microsoft/kiota-authentication-azure
  const authProvider = new AzureIdentityAuthenticationProvider(
    credential,
    scopes,
  );

  const requestAdapter = new GraphRequestAdapter(authProvider);

  const graphClient = createGraphServiceClient(requestAdapter);
  // </UserNamePasswordSnippet>

  return graphClient;
}
