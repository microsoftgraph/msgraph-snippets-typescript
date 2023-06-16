// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Client } from '@microsoft/microsoft-graph-client';
import {
  User,
  Message,
  Calendar,
  Team,
  Event,
} from '@microsoft/microsoft-graph-types';

export default async function runRequestSamples(
  graphClient: Client
): Promise<void> {
  // Create a new message
  const tempMessage: Message = await graphClient.api('/me/messages').post({
    subject: 'Temporary',
  });

  const messageId = tempMessage.id;
  if (!messageId) {
    throw new Error('Could not create a new message');
  }

  // Get a team to update
  const teams = await graphClient
    .api('/groups')
    .filter("resourceProvisioningOptions/Any(x:x eq 'Team')")
    .get();

  const teamId = teams.value[0]?.id;
  if (!teamId) {
    throw new Error('Could not get a team');
  }

  await makeReadRequest(graphClient);
  await makeSelectRequest(graphClient);
  await makeListRequest(graphClient);
  await makeItemByIdRequest(graphClient, messageId);
  await makeExpandRequest(graphClient, messageId);
  await makeDeleteRequest(graphClient, messageId);
  await makeCreateRequest(graphClient);
  await makeUpdateRequest(graphClient, teamId);
  await makeHeadersRequest(graphClient);
  await makeQueryParametersRequest(graphClient);
}

async function makeReadRequest(graphClient: Client): Promise<User> {
  // <ReadRequestSnippet>
  // GET https://graph.microsoft.com/v1.0/me
  const user = await graphClient.api('/me').get();
  // </ReadRequestSnippet>

  return user;
}

async function makeSelectRequest(graphClient: Client): Promise<User> {
  // <SelectRequestSnippet>
  // GET https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle
  const user = await graphClient
    .api('/me')
    .select(['displayName', 'jobTitle'])
    .get();
  // </SelectRequestSnippet>

  return user;
}

async function makeListRequest(graphClient: Client): Promise<Message[]> {
  // <ListRequestSnippet>
  // GET https://graph.microsoft.com/v1.0/me/messages?
  // $select=subject,sender&$filter=subject eq 'Hello world'
  const messages = await graphClient
    .api('/me/messages')
    .select(['subject', 'sender'])
    .filter(`subject eq 'Hello world'`)
    .get();
  // </ListRequestSnippet>

  return messages.value;
}

async function makeItemByIdRequest(
  graphClient: Client,
  messageId: string
): Promise<Message> {
  // <ItemByIdRequestSnippet>
  // GET https://graph.microsoft.com/v1.0/me/messages/{message-id}
  // messageId is a string containing the id property of the message
  const message = await graphClient.api(`/me/messages/${messageId}`).get();
  // </ItemByIdRequestSnippet>

  return message;
}

async function makeExpandRequest(
  graphClient: Client,
  messageId: string
): Promise<Message> {
  // <ExpandRequestSnippet>
  // <ExpandRequestSnippet>
  // GET https://graph.microsoft.com/v1.0/me/messages/{message-id}?$expand=attachments
  // messageId is a string containing the id property of the message
  const message = await graphClient
    .api(`/me/messages/${messageId}`)
    .expand('attachments')
    .get();
  // </ExpandRequestSnippet>

  return message;
}

async function makeDeleteRequest(
  graphClient: Client,
  messageId: string
): Promise<void> {
  // <DeleteRequestSnippet>
  // DELETE https://graph.microsoft.com/v1.0/me/messages/{message-id}
  // messageId is a string containing the id property of the message
  await graphClient.api(`/me/messages/${messageId}`).delete();
  // </DeleteRequestSnippet>
}

async function makeCreateRequest(graphClient: Client): Promise<Calendar> {
  // <CreateRequestSnippet>
  // POST https://graph.microsoft.com/v1.0/me/calendars
  const calendar: Calendar = {
    name: 'Volunteer',
  };

  const newCalendar = await graphClient.api('/me/calendars').post(calendar);
  // </CreateRequestSnippet>

  return newCalendar;
}

async function makeUpdateRequest(
  graphClient: Client,
  teamId: string
): Promise<void> {
  // <UpdateRequestSnippet>
  // PATCH https://graph.microsoft.com/v1.0/teams/{team-id}
  const team: Team = {
    funSettings: {
      allowGiphy: true,
      giphyContentRating: 'strict',
    },
  };

  // teamId is a string containing the id property of the team
  await graphClient.api(`/teams/${teamId}`).update(team);
  // </UpdateRequestSnippet>
}

async function makeHeadersRequest(graphClient: Client): Promise<Event[]> {
  // <HeadersRequestSnippet>
  // GET https://graph.microsoft.com/v1.0/me/events
  const events = await graphClient
    .api('/me/events')
    .header('Prefer', 'outlook.timezone="Pacific Standard Time"')
    .get();
  // </HeadersRequestSnippet>

  return events;
}

async function makeQueryParametersRequest(
  graphClient: Client
): Promise<Event[]> {
  // <QueryParametersRequestSnippet>
  // GET https://graph.microsoft.com/v1.0/me/calendarView?
  // startDateTime=2023-06-14T00:00:00Z&endDateTime=2023-06-15T00:00:00Z
  const events = await graphClient
    .api('me/calendar/calendarView')
    .query({
      startDateTime: '2023-06-14T00:00:00Z',
      endDateTime: '2023-06-15T00:00:00Z',
    })
    .get();
  // </QueryParametersRequestSnippet>

  return events;
}
