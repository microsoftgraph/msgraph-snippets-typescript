// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Client,
  GraphRequestOptions,
  PageCollection,
  PageIterator,
  PageIteratorCallback,
} from '@microsoft/microsoft-graph-client';
import { Message } from '@microsoft/microsoft-graph-types';

export default async function runPagingSamples(
  graphClient: Client,
): Promise<void> {
  await iterateAllMessages(graphClient);
  await iterateAllMessagesWithPause(graphClient);
  await manuallyPageAllMessages(graphClient);
}

async function iterateAllMessages(graphClient: Client): Promise<void> {
  // <PagingSnippet>
  const response: PageCollection = await graphClient
    .api('/me/messages?$top=10&$select=sender,subject,body')
    .header('Prefer', 'outlook.body-content-type="text"')
    .get();

  // A callback function to be called for every item in the collection.
  // This call back should return boolean indicating whether not to
  // continue the iteration process.
  const callback: PageIteratorCallback = (message: Message) => {
    console.log(message.subject);
    return true;
  };

  // A set of request options to be applied to
  // all subsequent page requests
  const requestOptions: GraphRequestOptions = {
    // Re-add the header to subsequent requests
    headers: {
      Prefer: 'outlook.body-content-type="text"',
    },
  };

  // Creating a new page iterator instance with client a graph client
  // instance, page collection response from request and callback
  const pageIterator = new PageIterator(
    graphClient,
    response,
    callback,
    requestOptions,
  );

  // This iterates the collection until the nextLink is drained out.
  await pageIterator.iterate();
  // </PagingSnippet>
}

async function iterateAllMessagesWithPause(graphClient: Client): Promise<void> {
  // <ResumePagingSnippet>
  let count = 0;
  const pauseAfter = 25;

  const response: PageCollection = await graphClient
    .api('/me/messages?$top=10&$select=sender,subject,body')
    .get();

  const callback: PageIteratorCallback = (message: Message) => {
    console.log(message.subject);
    count++;

    // If we've iterated over the limit,
    // stop the iteration by returning false
    return count < pauseAfter;
  };

  const pageIterator = new PageIterator(graphClient, response, callback);
  await pageIterator.iterate();

  while (!pageIterator.isComplete()) {
    console.log('Iteration paused for 5 seconds...');
    await new Promise((resolve) => setTimeout(resolve, 5000));

    // Reset count
    count = 0;
    await pageIterator.resume();
  }
  // </ResumePagingSnippet>
}

async function manuallyPageAllMessages(graphClient: Client): Promise<void> {
  // <ManualPagingSnippet>
  let response: PageCollection = await graphClient
    .api('/me/messages?$top=10')
    .get();

  while (response.value.length > 0) {
    for (const message of response.value as Message[]) {
      console.log(message.subject);
    }

    if (response['@odata.nextLink']) {
      response = await graphClient.api(response['@odata.nextLink']).get();
    } else {
      break;
    }
  }
  // </ManualPagingSnippet>
}
