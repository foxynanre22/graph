// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <graphServiceSnippet1>
import moment, { Moment } from 'moment';
import { Event } from 'microsoft-graph';
import { GraphRequestOptions, PageCollection, PageIterator } from '@microsoft/microsoft-graph-client';

var graph = require('@microsoft/microsoft-graph-client');

function getAuthenticatedClient(accessToken: string) {
  // Initialize Graph client
  const client = graph.Client.init({
    // Use the provided access token to authenticate
    // requests
    authProvider: (done: any) => {
      done(null, accessToken);
    }
  });

  return client;
}

export async function getUserDetails(accessToken: string) {
  const client = getAuthenticatedClient(accessToken);

  const user = await client
    .api('/me')
    .select('displayName,mail,mailboxSettings,userPrincipalName')
    .get();

  return user;
}
// </graphServiceSnippet1>

// <getUserWeekCalendarSnippet>
export async function getUserWeekCalendar(accessToken: string, timeZone: string, startDate: Moment): Promise<Event[]> {
  const client = getAuthenticatedClient(accessToken);

  // Generate startDateTime and endDateTime query params
  // to display a 7-day window
  var startDateTime = startDate.format();
  var endDateTime = moment(startDate).add(7, 'day').format();

  // GET /me/calendarview?startDateTime=''&endDateTime=''
  // &$select=subject,organizer,start,end
  // &$orderby=start/dateTime
  // &$top=50
  var response: PageCollection = await client
    .api('/me/calendarview')
    .header('Prefer', `outlook.timezone="${timeZone}"`)
    .query({ startDateTime: startDateTime, endDateTime: endDateTime })
    .select('subject,organizer,start,end')
    .orderby('start/dateTime')
    .top(25)
    .get();

  if (response["@odata.nextLink"]) {
    // Presence of the nextLink property indicates more results are available
    // Use a page iterator to get all results
    var events: Event[] = [];

    // Must include the time zone header in page
    // requests too
    var options: GraphRequestOptions = {
      headers: { 'Prefer': `outlook.timezone="${timeZone}"` }
    };

    var pageIterator = new PageIterator(client, response, (event) => {
      events.push(event);
      return true;
    }, options);

    await pageIterator.iterate();

    return events;
  } else {

    return response.value;
  }
}
// </getUserWeekCalendarSnippet>

// <createEventSnippet>
export async function createEvent(accessToken: string, newEvent: Event): Promise<Event> {
  const client = getAuthenticatedClient(accessToken);

  // POST /me/events
  // JSON representation of the new event is sent in the
  // request body
  return await client
    .api('/me/events')
    .post(newEvent);
}
// </createEventSnippet>

export const getUserId = async (accessToken: string, userEmail: string) => {
  const client = getAuthenticatedClient(accessToken);

  let resultGraph = await client.api(`/users/${userEmail}`).get();
  return resultGraph.id;
};

export const getCurrentUserId = async (accessToken: string) => {
  const client = getAuthenticatedClient(accessToken);

  let resultGraph = await client.api(`me`).get();
  return resultGraph.id;
};

export const createUsersChat = async (accessToken: string, requesterId: string, birthdayPersonId: string) => {
  let body: any = {
      "chatType": "oneOnOne",
      "members": [
          {
              "@odata.type": "#microsoft.graph.aadUserConversationMember",
              "roles": ["owner"],
              "user@odata.bind": `https://graph.microsoft.com/beta/users('${requesterId}')`
          },
          {
              "@odata.type": "#microsoft.graph.aadUserConversationMember",
              "roles": ["owner"],
              "user@odata.bind": `https://graph.microsoft.com/beta/users('${birthdayPersonId}')`
          }
      ]
  };

  const client = getAuthenticatedClient(accessToken);
  let resultGraph = await client.api(`chats`).version("beta").post(body);
  return resultGraph.id;
};

export const sendMessage = async (accessToken: string, chatId: string, chatMessage: string) => {
  let body = {
      "body": {
          "content": chatMessage
      }
  };

  const client = getAuthenticatedClient(accessToken);
  let resultGraph = await client.api(`chats/${chatId}/messages`).version("beta").post(body);
  return resultGraph;
};
