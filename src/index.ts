// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as readline from 'readline-sync';
import { DeviceCodeInfo } from '@azure/identity';
import appConfig from './appConfig.js';
import { getGraphClientForUser } from './graphHelper.js';
import runBatchSamples from './snippets/batchRequests.js';
import runRequestSamples from './snippets/createRequests.js';
import runLargeFileUploadSamples from './snippets/largeFileUpload.js';
//import runPagingSamples from './snippets/paging.js';

async function main() {
  const userClient = getGraphClientForUser(
    appConfig,
    (info: DeviceCodeInfo) => {
      console.log(info.message);
    },
  );

  try {
    const me = await userClient.me.get();
    console.log(`Hello, ${me?.displayName}!`);
  } catch (err) {
    console.log(`Error getting user: ${JSON.stringify(err, null, 2)}`);
  }

  let choice = 0;

  const choices = [
    'Run batch samples',
    'Run create request samples',
    'Run large file upload samples',
    'Run paging samples',
  ];

  while (choice != -1) {
    choice = readline.keyInSelect(choices, 'Select an option', {
      cancel: 'Exit',
    });

    switch (choice) {
      case -1:
        // Exit
        console.log('Goodbye...');
        break;
      case 0:
        await runBatchSamples(userClient);
        break;
      case 1:
        await runRequestSamples(userClient);
        break;
      case 2:
        //await runLargeFileUploadSamples(userClient, appConfig.largeFilePath);
        break;
      case 3:
        //await runPagingSamples(userClient);
        break;
      default:
        console.log('Invalid choice! Please try again.');
    }
  }
}

main();
