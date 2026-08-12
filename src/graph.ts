import type { GraphMessage, MailSummary } from './types.js';

/**
 * Microsoft Graph client.
 *
 * Two service limits shape everything here:
 *  - 10,000 requests per 10 minutes per app+mailbox, and
 *  - a hard cap of 4 CONCURRENT requests per app+mailbox. Exceeding it returns
 *    MailboxConcurrency errors, so all fan-out goes through `mapLimited`.
 *
 * JSON batching ($batch, 20 sub-requests) cuts round trips, but note each
 * sub-request still counts against the 10k/10min budget - batching buys latency,
 * not quota.
 */

const GRAPH = 'https://graph.microsoft.com/v1.0';
const MAILBOX_CONCURRENCY = 4;
const BATCH_SIZE = 20;

/** Statuses worth trying again. 429 is throttling; 502/503/504 are transient. */
const RETRYABLE_STATUS = new Set([429, 502, 503, 504]);
const MAX_ATTEMPTS = 5;

/**
 * Fallback wait when Graph throttles without telling us how long.
 *
 * Microsoft's guidance is 30s. Backing off by only a second or two - the
 * intuitive choice - actively makes things worse: each premature retry counts
 * against the same budget that triggered the throttle and extends it.
 */
const THROTTLE_FALLBACK_MS = 30_000;

/** Longest we'll block a sort run waiting out a throttle before giving up. */
const MAX_WAIT_MS = 60_000;

/** Fields we need to classify. Kept minimal: less payload, faster paging. */
const MESSAGE_SELECT =
  'id,subject,bodyPreview,receivedDateTime,hasAttachments,categories,from';

/**
 * Named MAPI property used to stamp what we assigned to a message.
 *
 * This is how corrections made in Outlook itself are detected: we record our own
 * verdict on the message, then on a later pass compare it against the categories
 * actually present. A divergence is the user having overruled us, which is the
 * most valuable training signal available and costs no extra request - the stamp
 * rides along in the same PATCH that writes the categories.
 *
 * Office.js CustomProperties can't serve here: they only reach the currently
 * selected item, and this has to work across a whole inbox.
 */
const PROVENANCE_GUID = 'c11ff724-aa03-4555-9952-8fa248a11c3e';
export const PROVENANCE_PROP_ID = `String {${PROVENANCE_GUID}} Name InboxStewardAssignment`;

/** `$expand` clause that brings our stamp back with a message. */
const PROVENANCE_EXPAND =
  `singleValueExtendedProperties($filter=id eq '${PROVENANCE_PROP_ID}')`;

/**
 * Serialized stamp: category id, confidence, ISO timestamp, learning generation.
 *
 * Append-only. Stamps written before a field existed simply lack it, and
 * decoding defaults rather than discarding - these live on messages in the
 * mailbox and cannot be migrated.
 */
export function encodeProvenance(
  categoryId: string,
  confidence: number,
  generation: number,
): string {
  return `${categoryId}|${confidence.toFixed(2)}|${new Date().toISOString()}|${generation}`;
}

export function decodeProvenance(
  value: string,
): { categoryId: string; confidence: number; at: string; generation: number } | null {
  const [categoryId, confidence, at, generation] = value.split('|');
  if (!categoryId) return null;
  return {
    categoryId,
    confidence: Number(confidence ?? '0'),
    at: at ?? '',
    // Absent on stamps from before generations existed. Zero means "older than
    // anything we know", so such a message is reconsidered exactly once.
    generation: Number(generation ?? '0') || 0,
  };
}

export class GraphError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly body?: unknown,
  ) {
    super(message);
  }
}

/** Run tasks with a fixed concurrency ceiling, preserving input order. */
async function mapLimited<T, R>(
  items: T[],
  limit: number,
  fn: (item: T, index: number) => Promise<R>,
): Promise<R[]> {
  const results = new Array<R>(items.length);
  let cursor = 0;
  const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
    for (;;) {
      const i = cursor++;
      if (i >= items.length) return;
      results[i] = await fn(items[i] as T, i);
    }
  });
  await Promise.all(workers);
  return results;
}

/**
 * Shared throttle gate.
 *
 * Graph throttles per app+mailbox, not per request, so when one call is told to
 * back off every other in-flight call is already over the same limit. Without a
 * shared gate the other three workers keep firing during the backoff, each
 * earning its own 429 and pushing the throttle window further out. One mailbox,
 * one clock.
 */
let throttledUntil = 0;

async function awaitThrottleGate(): Promise<void> {
  const waitMs = throttledUntil - Date.now();
  if (waitMs > 0) await sleep(Math.min(waitMs, MAX_WAIT_MS));
}

function openThrottleGateIn(ms: number): void {
  throttledUntil = Math.max(throttledUntil, Date.now() + ms);
}

/**
 * Case-insensitive header lookup.
 *
 * Sub-responses inside a `$batch` payload are plain JSON, not a `Headers`
 * object, so the casing Graph happens to emit is the casing we get.
 */
function headerValue(
  headers: Record<string, string> | undefined,
  name: string,
): string | null {
  if (!headers) return null;
  const wanted = name.toLowerCase();
  for (const [key, value] of Object.entries(headers)) {
    if (key.toLowerCase() === wanted) return value;
  }
  return null;
}

/**
 * How long to wait before retrying, given a Retry-After header.
 *
 * The header is seconds in practice but the HTTP spec also permits a date, and
 * Graph does occasionally send one.
 */
function retryDelayMs(retryAfter: string | null, status: number, attempt: number): number {
  if (retryAfter) {
    const seconds = Number.parseInt(retryAfter, 10);
    if (Number.isFinite(seconds) && seconds > 0) return seconds * 1000;

    const at = new Date(retryAfter).getTime();
    if (Number.isFinite(at)) {
      const delta = at - Date.now();
      if (delta > 0) return delta;
    }
  }

  if (status === 429) return THROTTLE_FALLBACK_MS;
  // Transient server errors clear quickly: 1s, 2s, 4s, 8s.
  return Math.min(1000 * 2 ** attempt, 8000);
}

/**
 * A single Graph call, with throttle and transient-failure handling.
 *
 * Graph tells us exactly how long to wait via Retry-After; honouring it is both
 * required and cheaper than guessing. Anything else is surfaced immediately -
 * silently swallowing a 403 would look like "no mail to sort".
 */
async function graphFetch(
  token: string,
  path: string,
  init: RequestInit = {},
  attempt = 0,
): Promise<Response> {
  await awaitThrottleGate();

  const url = path.startsWith('http') ? path : `${GRAPH}${path}`;

  let res: Response;
  try {
    res = await fetch(url, {
      ...init,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        ...(init.headers ?? {}),
      },
    });
  } catch (err) {
    // Dropped wifi mid-sort is ordinary in an office, and losing one request
    // shouldn't abandon the run. Only the network layer throws here; an HTTP
    // error status arrives as a normal response below.
    if (attempt + 1 >= MAX_ATTEMPTS) throw err;
    await sleep(retryDelayMs(null, 0, attempt));
    return graphFetch(token, path, init, attempt + 1);
  }

  if (RETRYABLE_STATUS.has(res.status) && attempt + 1 < MAX_ATTEMPTS) {
    const waitMs = retryDelayMs(res.headers.get('Retry-After'), res.status, attempt);
    if (waitMs > MAX_WAIT_MS) {
      throw new GraphError(
        `Graph is throttling this mailbox for ${Math.ceil(waitMs / 1000)}s`,
        res.status,
      );
    }
    if (res.status === 429) openThrottleGateIn(waitMs);
    await sleep(waitMs);
    return graphFetch(token, path, init, attempt + 1);
  }

  if (!res.ok) {
    let body: unknown;
    try {
      body = await res.json();
    } catch {
      body = await res.text().catch(() => undefined);
    }
    throw new GraphError(`Graph ${init.method ?? 'GET'} ${path} failed (${res.status})`, res.status, body);
  }
  return res;
}

function sleep(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

async function graphJson<T>(token: string, path: string, init?: RequestInit): Promise<T> {
  const res = await graphFetch(token, path, init);
  return (await res.json()) as T;
}

// ---------------------------------------------------------------------------
// Reading mail
// ---------------------------------------------------------------------------

export function toSummary(m: GraphMessage, includePreview: boolean): MailSummary {
  const listHeader = m.internetMessageHeaders?.find(
    (h) => h.name.toLowerCase() === 'list-id' || h.name.toLowerCase() === 'list-unsubscribe',
  );
  const summary: MailSummary = {
    id: m.id,
    from: m.from?.emailAddress.address?.toLowerCase() ?? '',
    fromName: m.from?.emailAddress.name ?? '',
    subject: m.subject ?? '(no subject)',
    received: m.receivedDateTime,
    hasAttachments: m.hasAttachments,
    categories: m.categories ?? [],
  };
  if (listHeader) summary.listId = listHeader.value;

  const stamp = m.singleValueExtendedProperties?.find((p) => p.id === PROVENANCE_PROP_ID);
  if (stamp) {
    const decoded = decodeProvenance(stamp.value);
    if (decoded) summary.assigned = decoded;
  }

  if (includePreview && m.bodyPreview) {
    // Cap the preview. Long bodies burn tokens for no accuracy gain - the
    // signal in an email is nearly all in the first couple of sentences.
    summary.preview = m.bodyPreview.slice(0, 600);
  }
  return summary;
}

/**
 * Mail she has already categorized, from ANYWHERE in the mailbox.
 *
 * Two things separate this from `listRecentInbox`, and both matter on the one
 * run that uses it:
 *
 * Scope. Filed mail is the labelled training set, and filed mail is by
 * definition not in the inbox any more. Bootstrapping from the inbox alone
 * looks at the pile she hasn't dealt with and misses the years of decisions
 * she has.
 *
 * Depth. Every sender learned here is a sender the LLM never has to read, now
 * or ever. Free-tier quota is the binding constraint on this whole design, so
 * a one-off deeper scan on first run pays for itself immediately and forever.
 *
 * The payload stays small - no preview, no provenance - so several hundred
 * messages cost little. Uncategorized messages are dropped client-side because
 * Graph has no dependable filter for "has any category".
 */
export async function listCategorizedHistory(
  token: string,
  limit: number,
): Promise<MailSummary[]> {
  const collected: MailSummary[] = [];
  let scanned = 0;

  let path =
    `/me/messages` +
    `?$select=id,subject,receivedDateTime,hasAttachments,categories,from` +
    `&$orderby=receivedDateTime desc&$top=50`;

  // Bounded by messages examined, not messages kept: a mailbox where nothing is
  // categorized must not page to the end of time looking for one that is.
  while (scanned < limit) {
    const page = await graphJson<{ value: GraphMessage[]; '@odata.nextLink'?: string }>(token, path);
    scanned += page.value.length;

    for (const m of page.value) {
      if ((m.categories?.length ?? 0) > 0) collected.push(toSummary(m, false));
    }

    const next = page['@odata.nextLink'];
    if (!next || page.value.length === 0) break;
    path = next;
  }

  return collected;
}

/** Most recent inbox messages, newest first. */
export async function listRecentInbox(
  token: string,
  limit: number,
  includePreview: boolean,
): Promise<MailSummary[]> {
  const collected: GraphMessage[] = [];
  let path =
    `/me/mailFolders/inbox/messages` +
    `?$select=${MESSAGE_SELECT}&$expand=${PROVENANCE_EXPAND}` +
    `&$orderby=receivedDateTime desc&$top=${Math.min(limit, 50)}`;

  while (collected.length < limit) {
    const page = await graphJson<{ value: GraphMessage[]; '@odata.nextLink'?: string }>(token, path);
    collected.push(...page.value);
    const next = page['@odata.nextLink'];
    if (!next) break;
    path = next;
  }

  return collected.slice(0, limit).map((m) => toSummary(m, includePreview));
}

/**
 * Incremental read via delta query.
 *
 * Delta is the right primitive for a client-side add-in: unlike change
 * notifications it needs no public HTTPS endpoint and no subscription renewal
 * (mail subscriptions expire after ~2.9 days and must be actively renewed).
 * Pass the previous deltaLink to get only what changed since.
 */
export async function deltaInbox(
  token: string,
  deltaLink: string | null,
  includePreview: boolean,
  maxPages = 20,
): Promise<{ messages: MailSummary[]; deltaLink: string | null; removed: string[] }> {
  let path =
    deltaLink ??
    `/me/mailFolders/inbox/messages/delta` +
      `?$select=${MESSAGE_SELECT}&$expand=${PROVENANCE_EXPAND}&changeType=created,updated`;

  const messages: MailSummary[] = [];
  const removed: string[] = [];
  let nextDelta: string | null = null;

  for (let page = 0; page < maxPages; page++) {
    const body = await graphJson<{
      value: (GraphMessage & { '@removed'?: { reason: string } })[];
      '@odata.nextLink'?: string;
      '@odata.deltaLink'?: string;
    }>(token, path, { headers: { Prefer: 'odata.maxpagesize=50' } });

    for (const item of body.value) {
      if (item['@removed']) removed.push(item.id);
      else messages.push(toSummary(item, includePreview));
    }

    if (body['@odata.deltaLink']) {
      nextDelta = body['@odata.deltaLink'];
      break;
    }
    if (!body['@odata.nextLink']) break;
    path = body['@odata.nextLink'];
  }

  return { messages, deltaLink: nextDelta, removed };
}

// ---------------------------------------------------------------------------
// Writing categories
// ---------------------------------------------------------------------------

export interface CategoryUpdate {
  messageId: string;
  /** Full replacement set. Graph PATCH on `categories` overwrites, not merges. */
  categories: string[];
  /** Our verdict, stamped alongside so later corrections are detectable. */
  provenance?: { categoryId: string; confidence: number; generation: number };
}

/**
 * Apply categories to many messages.
 *
 * Returns the ids that failed rather than throwing, so one bad message never
 * aborts a whole sort run. Partial success is normal and worth reporting.
 */
export async function applyCategories(
  token: string,
  updates: CategoryUpdate[],
): Promise<{ succeeded: string[]; failed: { id: string; reason: string }[] }> {
  const succeeded: string[] = [];
  const failed: { id: string; reason: string }[] = [];

  const batches: CategoryUpdate[][] = [];
  for (let i = 0; i < updates.length; i += BATCH_SIZE) {
    batches.push(updates.slice(i, i + BATCH_SIZE));
  }

  function subRequest(u: CategoryUpdate, id: string) {
    return {
      id,
      method: 'PATCH',
      url: `/me/messages/${u.messageId}`,
      headers: { 'Content-Type': 'application/json' },
      body: {
        categories: u.categories,
        ...(u.provenance
          ? {
              singleValueExtendedProperties: [
                {
                  id: PROVENANCE_PROP_ID,
                  value: encodeProvenance(
                    u.provenance.categoryId,
                    u.provenance.confidence,
                    u.provenance.generation,
                  ),
                },
              ],
            }
          : {}),
      },
    };
  }

  await mapLimited(batches, MAILBOX_CONCURRENCY, async (batch) => {
    // A $batch call returns 200 even when individual sub-requests were
    // throttled - the envelope succeeded, its contents did not. So the retry
    // in graphFetch never sees these and they have to be handled here.
    // Treating a sub-request 429 as final is the difference between "sorted
    // 40 of 200" and sorting all 200: throttling is transient, and the messages
    // it hits are the ones that silently never get labelled.
    let pending = batch.map((u, i) => ({ update: u, id: String(i) }));

    for (let attempt = 0; attempt < MAX_ATTEMPTS && pending.length > 0; attempt++) {
      await awaitThrottleGate();

      let responses: { id: string; status: number; headers?: Record<string, string>; body?: unknown }[];
      try {
        const res = await graphJson<{
          responses: typeof responses;
        }>(token, '/$batch', {
          method: 'POST',
          body: JSON.stringify({ requests: pending.map((p) => subRequest(p.update, p.id)) }),
        });
        responses = res.responses;
      } catch (err) {
        for (const p of pending) {
          failed.push({
            id: p.update.messageId,
            reason: err instanceof Error ? err.message : 'unknown',
          });
        }
        return;
      }

      const byId = new Map(pending.map((p) => [p.id, p.update]));
      const retry: typeof pending = [];
      let longestWaitMs = 0;
      let throttled = false;

      for (const r of responses) {
        const update = byId.get(r.id);
        if (!update) continue;

        if (r.status >= 200 && r.status < 300) {
          succeeded.push(update.messageId);
        } else if (RETRYABLE_STATUS.has(r.status) && attempt + 1 < MAX_ATTEMPTS) {
          const waitMs = retryDelayMs(headerValue(r.headers, 'Retry-After'), r.status, attempt);
          longestWaitMs = Math.max(longestWaitMs, waitMs);
          throttled = throttled || r.status === 429;
          retry.push({ update, id: r.id });
        } else {
          failed.push({ id: update.messageId, reason: `status ${r.status}` });
        }
      }

      if (retry.length > 0 && longestWaitMs > MAX_WAIT_MS) {
        // Waiting this long would strand the user staring at a spinner. Report
        // it plainly instead - the next run picks these up, unlabelled mail is
        // still exactly where they left it. Deliberately without opening the
        // throttle gate: a gate longer than we're willing to wait would just
        // make every later call in this session sleep to no purpose.
        for (const p of retry) {
          failed.push({ id: p.update.messageId, reason: 'mailbox busy, try again shortly' });
        }
        return;
      }

      if (throttled) openThrottleGateIn(longestWaitMs);

      pending = retry;
      if (pending.length > 0) await sleep(longestWaitMs);
    }
  });

  return { succeeded, failed };
}

// ---------------------------------------------------------------------------
// Master category list
// ---------------------------------------------------------------------------

export interface OutlookCategory {
  id: string;
  displayName: string;
  color: string;
}

export async function getMasterCategories(token: string): Promise<OutlookCategory[]> {
  const body = await graphJson<{ value: OutlookCategory[] }>(token, '/me/outlook/masterCategories');
  return body.value;
}

/**
 * Create any of our categories that the mailbox doesn't have yet.
 *
 * This is not optional housekeeping: a category MUST exist in the mailbox
 * master list before it can be applied to a message. Skipping it makes every
 * PATCH silently useless.
 */
export async function ensureMasterCategories(
  token: string,
  wanted: { name: string; color: string }[],
): Promise<{ created: string[]; existing: string[] }> {
  const existingList = await getMasterCategories(token);
  const existingNames = new Set(existingList.map((c) => c.displayName.trim().toLowerCase()));

  const missing = wanted.filter((w) => !existingNames.has(w.name.trim().toLowerCase()));
  const created: string[] = [];

  // Sequential on purpose: the master list is a single small resource and
  // parallel creates on it invite conflicts for no meaningful speedup.
  for (const cat of missing) {
    try {
      await graphJson(token, '/me/outlook/masterCategories', {
        method: 'POST',
        body: JSON.stringify({ displayName: cat.name, color: cat.color }),
      });
      created.push(cat.name);
    } catch (err) {
      // A duplicate displayName is the likely cause and is harmless - another
      // client may have created it between our read and write.
      console.warn(`[graph] could not create category "${cat.name}"`, err);
    }
  }

  return { created, existing: wanted.filter((w) => existingNames.has(w.name.trim().toLowerCase())).map((w) => w.name) };
}

// ---------------------------------------------------------------------------
// Native Outlook rules
// ---------------------------------------------------------------------------

export interface MessageRule {
  id: string;
  displayName: string;
  sequence: number;
  isEnabled: boolean;
  conditions?: { senderContains?: string[]; fromAddresses?: { emailAddress: { address: string } }[] };
  actions?: { assignCategories?: string[]; stopProcessingRules?: boolean };
}

export async function listMessageRules(token: string): Promise<MessageRule[]> {
  const body = await graphJson<{ value: MessageRule[] }>(
    token,
    '/me/mailFolders/inbox/messageRules',
  );
  return body.value;
}

export async function createMessageRule(
  token: string,
  rule: Omit<MessageRule, 'id'>,
): Promise<MessageRule> {
  return graphJson<MessageRule>(token, '/me/mailFolders/inbox/messageRules', {
    method: 'POST',
    body: JSON.stringify(rule),
  });
}

export async function updateMessageRule(
  token: string,
  id: string,
  patch: Partial<Omit<MessageRule, 'id'>>,
): Promise<void> {
  await graphFetch(token, `/me/mailFolders/inbox/messageRules/${id}`, {
    method: 'PATCH',
    body: JSON.stringify(patch),
  });
}

export async function deleteMessageRule(token: string, id: string): Promise<void> {
  await graphFetch(token, `/me/mailFolders/inbox/messageRules/${id}`, { method: 'DELETE' });
}

/** Existing folder names, used to bootstrap sender rules from prior filing. */
export async function listMailFolders(token: string): Promise<{ id: string; displayName: string }[]> {
  const body = await graphJson<{ value: { id: string; displayName: string }[] }>(
    token,
    '/me/mailFolders?$top=100&$select=id,displayName',
  );
  return body.value;
}
