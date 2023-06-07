// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as config from 'config';

export interface AppConfig {
  clientId: string;
  tenantId: string;
  graphUserScopes: string[];
}

const appConfig: AppConfig = {
  clientId: config.get<string>('clientId'),
  tenantId: config.get<string>('tenantId'),
  graphUserScopes: config.get<string[]>('graphUserScopes'),
};

if (!appConfig.clientId || appConfig.clientId.length <= 0) {
  throw new Error('clientId missing or empty from config.');
}

if (!appConfig.tenantId || appConfig.clientId.length <= 0) {
  throw new Error('tenantId missing or empty from config.');
}

if (!appConfig.graphUserScopes || appConfig.graphUserScopes.length <= 0) {
  throw new Error('graphUserScopes missing or empty from config.');
}

export default appConfig;
