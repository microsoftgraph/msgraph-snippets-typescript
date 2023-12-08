// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Client,
  BatchRequestStep,
  BatchRequestContent,
  BatchResponseContent,
  PageCollection,
} from '@microsoft/microsoft-graph-client';
import { User, Event } from '@microsoft/microsoft-graph-types';
import {
  startOfToday,
  endOfToday,
  setHours,
  setMinutes,
  format,
} from 'date-fns';

export default async function runBatchSamples(graphClient: Client) {
  await simpleBatch(graphClient);
  await dependentBatch(graphClient);
}

async function simpleBatch(graphClient: Client) {
  // <SimpleBatchSnippet>
  // Create a batch request step to GET /me
  // Request is from fetch polyfill, i.e. node-fetch
  const userRequestStep: BatchRequestStep = {
    id: '1',
    request: new Request('/me', {
      method: 'GET',
    }),
  };

  // startOfToday and endOfToday from date-fns
  const start = startOfToday().toISOString();
  const end = endOfToday().toISOString();

  // Create a batch request step to GET
  // /me/calendarView?startDateTime="start"&endDateTime="end"
  const calendarViewRequestStep: BatchRequestStep = {
    id: '2',
    request: new Request(
      `/me/calendarView?startDateTime=${start}&endDateTime=${end}`,
      {
        method: 'GET',
      },
    ),
  };

  // Create the batch request content with the steps created
  // above
  const batchRequestContent = new BatchRequestContent([
    userRequestStep,
    calendarViewRequestStep,
  ]);

  const content = await batchRequestContent.getContent();

  // POST the batch request content to the /$batch endpoint
  const batchResponse = await graphClient.api('/$batch').post(content);

  // Create a BatchResponseContent object to parse the response
  const batchResponseContent = new BatchResponseContent(batchResponse);

  // Get the user response using the id assigned to the request
  const userResponse = batchResponseContent.getResponseById('1');

  // For a single entity, the JSON payload can be deserialized
  // into the expected type
  // Types supplied by @microsoft/microsoft-graph-types
  if (userResponse.ok) {
    const user: User = (await userResponse.json()) as User;
    console.log(`Hello ${user.displayName}!`);
  } else {
    console.log(`Get user failed with status ${userResponse.status}`);
  }

  // Get the calendar view response by id
  const calendarResponse = batchResponseContent.getResponseById('2');

  // For a collection of entities, the "value" property of
  // the JSON payload can be deserialized into an array of
  // the expected type
  if (calendarResponse.ok) {
    const rawResponse = (await calendarResponse.json()) as PageCollection;
    const events: Event[] = rawResponse.value;
    console.log(`You have ${events.length} events on your calendar today.`);
  } else {
    console.log(
      `Get calendar view failed with status ${calendarResponse.status}`,
    );
  }
  // </SimpleBatchSnippet>
}

async function dependentBatch(graphClient: Client) {
  // <DependentBatchSnippet>
  // 5:00 PM today
  // startOfToday, endOfToday, setHours, setMinutes, format from date-fns
  const eventStart = setHours(startOfToday(), 17);
  const eventEnd = setMinutes(eventStart, 30);

  // Create a batch request step to add an event
  const newEvent: Event = {
    subject: 'File end-of-day report',
    start: {
      dateTime: format(eventStart, `yyyy-MM-dd'T'HH:mm:ss`),
      timeZone: 'Pacific Standard Time',
    },
    end: {
      // 5:30 PM
      dateTime: format(eventEnd, `yyyy-MM-dd'T'HH:mm:ss`),
      timeZone: 'Pacific Standard Time',
    },
  };

  // Request is from fetch polyfill, i.e. node-fetch
  const addEventRequestStep: BatchRequestStep = {
    id: '1',
    request: new Request('/me/events', {
      method: 'POST',
      body: JSON.stringify(newEvent),
      headers: {
        'Content-Type': 'application/json',
      },
    }),
  };

  const start = startOfToday().toISOString();
  const end = endOfToday().toISOString();

  // Create a batch request step to GET
  // /me/calendarView?startDateTime="start"&endDateTime="end"
  const calendarViewRequestStep: BatchRequestStep = {
    id: '2',
    // This step will happen after step 1
    dependsOn: ['1'],
    request: new Request(
      `/me/calendarView?startDateTime=${start}&endDateTime=${end}`,
      {
        method: 'GET',
      },
    ),
  };

  // Create the batch request content with the steps created
  // above
  const batchRequestContent = new BatchRequestContent([
    addEventRequestStep,
    calendarViewRequestStep,
  ]);

  const content = await batchRequestContent.getContent();

  // POST the batch request content to the /$batch endpoint
  const batchResponse = await graphClient.api('/$batch').post(content);

  // Create a BatchResponseContent object to parse the response
  const batchResponseContent = new BatchResponseContent(batchResponse);

  // Get the create event response by id
  const newEventResponse = batchResponseContent.getResponseById('1');
  if (newEventResponse.ok) {
    const event: Event = (await newEventResponse.json()) as Event;
    console.log(`New event created with ID: ${event.id}`);
  } else {
    console.log(`Create event failed with status ${newEventResponse.status}`);
  }

  // Get the calendar view response by id
  const calendarResponse = batchResponseContent.getResponseById('2');

  if (calendarResponse.ok) {
    // For a collection of entities, the "value" property of
    // the JSON payload can be deserialized into an array of
    // the expected type
    const rawResponse = (await calendarResponse.json()) as PageCollection;
    const events: Event[] = rawResponse.value;
    console.log(`You have ${events.length} events on your calendar today.`);
  } else {
    console.log(
      `Get calendar view failed with status ${calendarResponse.status}`,
    );
  }
  // </DependentBatchSnippet>
}
