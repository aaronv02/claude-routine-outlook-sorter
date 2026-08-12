import type { DigestEvent, DigestMessage } from './types.js';

/**
 * Reading the mailbox and calendar for one digest.
 *
 * Separate from the analysis on purpose: everything in the other digest modules
 * is a pure function, and this is the only file here that touches the network.
 *
 * Three service limits shape it, the same ones the sorter contends with: 10,000
 * requests per 10 minutes per app+mailbox, at most 4 concurrent, and paging at
 * `$top=100`. A weekly digest is small enough that sequential paging is fine.
 */

const GRAPH = 'https://graph.microsoft.com/v1.0';

const MESSAGE_SELECT = [
  'id',
  'conversationId',
  'subject',
  'bodyPreview',
  'receivedDateTime',
  'sentDateTime',
  'isRead',
  'isDraft',
  'hasAttachments',
  'importance',
  'categories',
  'from',
  'sender',
  'toRecipients',
  'ccRecipients',
  'flag',
  'webLink',
].join(',');

const EVENT_SELECT = [
  'id',
  'subject',
  'bodyPreview',
  'start',
  'end',
  'isAllDay',
  'isCancelled',
  'isOrganizer',
  'responseStatus',
  'organizer',
  'attendees',
  'location',
  'showAs',
  'webLink',
].join(',');

/** Guard against paging forever on a mailbox with a very long history. */
const MAX_PAGES = 20;

async function pagedGet<T>(token: string, url: string, timeZone?: string): Promise<T[]> {
  const collected: T[] = [];
  let next: string | undefined = url;

  for (let page = 0; page < MAX_PAGES && next; page++) {
    const res = await fetch(next, {
      headers: {
        Authorization: `Bearer ${token}`,
        // Makes Graph return calendar times already converted to the mailbox's
        // zone. Without it they arrive in UTC and every event in the report is
        // offset by hours.
        ...(timeZone ? { Prefer: `outlook.timezone="${timeZone}"` } : {}),
      },
    });

    if (!res.ok) {
      throw new Error(`Graph GET failed (${res.status}): ${(await res.text()).slice(0, 300)}`);
    }

    const body = (await res.json()) as { value?: T[]; '@odata.nextLink'?: string };
    collected.push(...(body.value ?? []));
    next = body['@odata.nextLink'];
  }

  return collected;
}

export async function fetchInboxSince(token: string, since: Date): Promise<DigestMessage[]> {
  const url =
    `${GRAPH}/me/mailFolders/inbox/messages` +
    `?$select=${MESSAGE_SELECT}` +
    `&$filter=${encodeURIComponent(`receivedDateTime ge ${since.toISOString()}`)}` +
    `&$orderby=receivedDateTime%20desc&$top=100`;
  return pagedGet<DigestMessage>(token, url);
}

/**
 * Sent mail since a point in time.
 *
 * Called with an earlier `since` than the inbox fetch: a reply sent last week
 * still answers a message received last week, and if the sent window is too
 * narrow, answered mail comes back looking unanswered.
 */
export async function fetchSentSince(token: string, since: Date): Promise<DigestMessage[]> {
  const url =
    `${GRAPH}/me/mailFolders/sentitems/messages` +
    `?$select=${MESSAGE_SELECT}` +
    `&$filter=${encodeURIComponent(`sentDateTime ge ${since.toISOString()}`)}` +
    `&$orderby=sentDateTime%20desc&$top=100`;
  return pagedGet<DigestMessage>(token, url);
}

/**
 * Calendar entries between two instants.
 *
 * `calendarView` rather than `/events` because it expands recurring series into
 * their individual occurrences. Querying `/events` returns the recurrence master
 * instead, so a weekly standing meeting would appear once with the wrong date and
 * every actual occurrence would be missing.
 */
export async function fetchCalendar(
  token: string,
  start: Date,
  end: Date,
  timeZone: string,
): Promise<DigestEvent[]> {
  const url =
    `${GRAPH}/me/calendarView` +
    `?startDateTime=${encodeURIComponent(start.toISOString())}` +
    `&endDateTime=${encodeURIComponent(end.toISOString())}` +
    `&$select=${EVENT_SELECT}` +
    `&$orderby=start/dateTime&$top=100`;
  return pagedGet<DigestEvent>(token, url, timeZone);
}

/** The signed-in mailbox, for labelling the report and for the "is this me?" checks. */
export async function fetchMailbox(token: string): Promise<{ address: string; name: string }> {
  const res = await fetch(`${GRAPH}/me?$select=displayName,mail,userPrincipalName`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Could not read the signed-in account (${res.status}).`);

  const body = (await res.json()) as {
    displayName?: string;
    mail?: string;
    userPrincipalName?: string;
  };
  return {
    address: (body.mail ?? body.userPrincipalName ?? '').toLowerCase(),
    name: body.displayName ?? '',
  };
}
