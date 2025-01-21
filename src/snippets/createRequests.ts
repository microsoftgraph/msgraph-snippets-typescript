// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { GraphServiceClient } from '@microsoft/msgraph-sdk';
import {
  Calendar,
  Event,
  Message,
  Team,
  User,
} from '@microsoft/msgraph-sdk/models';

export default async function runRequestSamples(
  graphClient: GraphServiceClient,
): Promise<void> {
  try {
    // Create a new message
    const tempMessage = await graphClient.me.messages.post({
      subject: 'Temporary',
    });

    const messageId = tempMessage?.id;
    if (!messageId) {
      throw new Error('Could not create a new message');
    }

    // Get a team to update
    const teams = await graphClient.groups.get({
      queryParameters: {
        filter: "resourceProvisioningOptions/Any(x:x eq 'Team')",
      },
    });

    const teamId = teams?.value?.at(0)?.id;
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
  } catch (err) {
    console.log(JSON.stringify(err, null, 2));
  }
}

async function makeReadRequest(
  graphClient: GraphServiceClient,
): Promise<User | undefined> {
  // <ReadRequestSnippet>
  // GET https://graph.microsoft.com/v1.0/me
  const user = await graphClient.me.get();
  // </ReadRequestSnippet>

  return user;
}

async function makeSelectRequest(
  graphClient: GraphServiceClient,
): Promise<User | undefined> {
  // <SelectRequestSnippet>
  // GET https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle
  const user = await graphClient.me.get({
    queryParameters: {
      select: ['displayName', 'jobTitle'],
    },
  });
  // </SelectRequestSnippet>

  return user;
}

async function makeListRequest(
  graphClient: GraphServiceClient,
): Promise<Message[]> {
  // <ListRequestSnippet>
  // GET https://graph.microsoft.com/v1.0/me/messages?
  // $select=subject,sender&$filter=subject eq 'Hello world'
  const messages = await graphClient.me.messages.get({
    queryParameters: {
      select: ['subject', 'sender'],
      filter: `subject eq 'Hello world'`,
    },
  });
  // </ListRequestSnippet>

  return messages?.value ?? [];
}

async function makeItemByIdRequest(
  graphClient: GraphServiceClient,
  messageId: string,
): Promise<Message | undefined> {
  // <ItemByIdRequestSnippet>
  // GET https://graph.microsoft.com/v1.0/me/messages/{message-id}
  // messageId is a string containing the id property of the message
  const message = await graphClient.me.messages.byMessageId(messageId).get();
  // </ItemByIdRequestSnippet>

  return message;
}

async function makeExpandRequest(
  graphClient: GraphServiceClient,
  messageId: string,
): Promise<Message | undefined> {
  // <ExpandRequestSnippet>
  // GET https://graph.microsoft.com/v1.0/me/messages/{message-id}?$expand=attachments
  // messageId is a string containing the id property of the message
  const message = await graphClient.me.messages.byMessageId(messageId).get({
    queryParameters: {
      expand: ['attachments'],
    },
  });
  // </ExpandRequestSnippet>

  return message;
}

async function makeDeleteRequest(
  graphClient: GraphServiceClient,
  messageId: string,
): Promise<void> {
  // <DeleteRequestSnippet>
  // DELETE https://graph.microsoft.com/v1.0/me/messages/{message-id}
  // messageId is a string containing the id property of the message
  await graphClient.me.messages.byMessageId(messageId).delete();
  // </DeleteRequestSnippet>
}

async function makeCreateRequest(
  graphClient: GraphServiceClient,
): Promise<Calendar | undefined> {
  // <CreateRequestSnippet>
  // POST https://graph.microsoft.com/v1.0/me/calendars
  const calendar: Calendar = {
    name: 'Volunteer',
  };

  const newCalendar = await graphClient.me.calendars.post(calendar);
  // </CreateRequestSnippet>

  return newCalendar;
}

async function makeUpdateRequest(
  graphClient: GraphServiceClient,
  teamId: string,
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
  await graphClient.teams.byTeamId(teamId).patch(team);
  // </UpdateRequestSnippet>
}

async function makeHeadersRequest(
  graphClient: GraphServiceClient,
): Promise<Event[]> {
  // <HeadersRequestSnippet>
  // GET https://graph.microsoft.com/v1.0/me/events
  const events = await graphClient.me.events.get({
    headers: {
      Prefer: 'outlook.timezone="Pacific Standard Time"',
    },
  });
  // </HeadersRequestSnippet>

  return events?.value ?? [];
}

async function makeQueryParametersRequest(
  graphClient: GraphServiceClient,
): Promise<Event[]> {
  // <QueryParametersRequestSnippet>
  // GET https://graph.microsoft.com/v1.0/me/calendarView?
  // startDateTime=2023-06-14T00:00:00Z&endDateTime=2023-06-15T00:00:00Z
  const events = await graphClient.me.calendar.calendarView.get({
    queryParameters: {
      startDateTime: '2023-06-14T00:00:00Z',
      endDateTime: '2023-06-15T00:00:00Z',
    },
  });
  // </QueryParametersRequestSnippet>

  return events?.value ?? [];
}
