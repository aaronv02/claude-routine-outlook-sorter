import type { PersistedState } from '../src/types.js';
import { defaultState, normalizeTaxonomy, STATE_VERSION } from '../src/taxonomy.js';

/**
 * Where the routine's state lives.
 *
 * The add-in keeps state in `roamingSettings`, which is an Office.js surface and
 * therefore unreachable from a cron job. A scheduled cloud run also has no disk
 * that survives to the next run, so the state has to be somewhere durable that
 * both ends can see.
 *
 * It goes in the mailbox: a hidden mail folder holding one message whose body is
 * the JSON. That keeps the add-in's best property intact - nothing about this
 * mailbox is stored on anyone else's server, and deleting the folder deletes
 * every trace of the tool. It also needs no permission beyond the
 * `Mail.ReadWrite` the add-in already asks for.
 *
 * `isHidden` folders don't appear in any Outlook client's folder list, so this
 * doesn't clutter the mailbox she actually looks at.
 *
 * Set STEWARD_STATE_FILE to keep state in a local file instead. That's for
 * development against a test mailbox; a real routine should use the mailbox so
 * state follows the mailbox rather than the machine.
 */

const GRAPH = 'https://graph.microsoft.com/v1.0';
const FOLDER_NAME = 'Inbox Steward';
const STATE_SUBJECT = 'inbox-steward-state';

/** Routine-only fields, kept beside the state the add-in also understands. */
export interface RoutineState extends PersistedState {
  routine?: {
    /** Rotated on every token exchange. See the honesty note in README. */
    refreshToken?: string;
    lastRunAt?: string;
    /** Rolling tally of sorting sweeps. */
    runs?: number;
  };
  /** The weekly digest's own memory. Separate because the two run on different schedules. */
  weekly?: {
    /**
     * ISO week of the last digest produced, e.g. "2026-W33".
     *
     * This is what makes a late run report the week it missed instead of a week
     * that has barely started - and what stops it reporting the same week twice.
     */
    lastReportedWeek?: string;
    lastRunAt?: string;
    /** Remembered so the reporting window doesn't move if the env var is dropped. */
    timeZone?: string;
    /** Role aliases and shared mailboxes she is also reached at. */
    alsoAddressedAs?: string[];
  };
}

async function graph<T>(token: string, path: string, init?: RequestInit): Promise<T> {
  const res = await fetch(path.startsWith('http') ? path : `${GRAPH}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(init?.headers ?? {}),
    },
  });
  if (!res.ok) {
    throw new Error(`Graph ${init?.method ?? 'GET'} ${path} failed (${res.status}): ${await res.text()}`);
  }
  return res.status === 204 ? (undefined as T) : ((await res.json()) as T);
}

async function findFolder(token: string): Promise<string | null> {
  // Hidden folders are excluded from the default listing, so the query
  // parameter is load-bearing - without it we'd create a second folder on
  // every run and lose the state each time.
  const body = await graph<{ value: { id: string; displayName: string }[] }>(
    token,
    '/me/mailFolders?includeHiddenFolders=true&$top=200&$select=id,displayName',
  );
  return body.value.find((f) => f.displayName === FOLDER_NAME)?.id ?? null;
}

async function ensureFolder(token: string): Promise<string> {
  const existing = await findFolder(token);
  if (existing) return existing;
  const created = await graph<{ id: string }>(token, '/me/mailFolders', {
    method: 'POST',
    body: JSON.stringify({ displayName: FOLDER_NAME, isHidden: true }),
  });
  return created.id;
}

/**
 * The one message holding our state.
 *
 * Filtered by subject rather than "whatever is newest in the folder". Without the
 * filter, any other message that ever landed here would be read as state - and
 * worse, a `$top` window that missed the real one made `saveState` create a SECOND
 * state message, so every subsequent run wrote to a different one and the mailbox
 * accumulated state messages while appearing to forget everything.
 */
async function findStateMessage(token: string, folderId: string) {
  const body = await graph<{ value: { id: string; body?: { content?: string } }[] }>(
    token,
    `/me/mailFolders/${folderId}/messages` +
      `?$filter=subject eq '${STATE_SUBJECT}'` +
      `&$top=10&$select=id,body&$orderby=createdDateTime desc`,
  );
  return body.value.find((m) => m.body?.content !== undefined) ?? null;
}

function migrate(raw: unknown): RoutineState {
  const base = defaultState() as RoutineState;
  if (!raw || typeof raw !== 'object') return base;
  const parsed = raw as Partial<RoutineState>;

  // Same posture as the add-in's own state loader: a payload from a different
  // version is repaired field by field rather than discarded, because throwing
  // it away silently resets everything the tool has learned.
  return {
    ...base,
    ...parsed,
    version: STATE_VERSION,
    settings: { ...base.settings, ...(parsed.settings ?? {}) },
    taxonomy: normalizeTaxonomy(parsed.taxonomy?.length ? parsed.taxonomy : base.taxonomy),
    senderRules: parsed.senderRules ?? [],
    recentCorrections: parsed.recentCorrections ?? [],
    ...(parsed.routine ? { routine: parsed.routine } : {}),
    ...(parsed.weekly ? { weekly: parsed.weekly } : {}),
  };
}

export async function loadState(token: string): Promise<RoutineState> {
  const file = process.env.STEWARD_STATE_FILE;
  if (file) {
    const { readFile } = await import('node:fs/promises');
    try {
      return migrate(JSON.parse(await readFile(file, 'utf8')));
    } catch {
      return migrate(null);
    }
  }

  const folderId = await ensureFolder(token);
  const message = await findStateMessage(token, folderId);
  if (!message?.body?.content) return migrate(null);

  try {
    return migrate(JSON.parse(extractJson(message.body.content)));
  } catch {
    // Unparseable state is worse than none only if we act on it. Starting over
    // costs a bootstrap run; acting on half-read state can promote a wrong
    // sender into a permanent Outlook rule.
    console.warn('[store] state message could not be parsed; starting from defaults.');
    return migrate(null);
  }
}

/**
 * Marker prefixing the base64 payload, so an old plain-JSON state message is
 * still readable after an upgrade.
 */
const B64_MARKER = 'INBOX-STEWARD-B64:';

/**
 * Recover the JSON from a message body.
 *
 * Graph does not reliably store a body as the content type it was given - a
 * `text` body can come back wrapped in HTML with its characters escaped.
 * Round-tripping raw JSON through that is lossy: `&#39;` and `&nbsp;` are not in
 * any short unescape list, and a subject like `John <john@x.com>` - entirely
 * ordinary inside a correction record - is indistinguishable from markup once it
 * is in the body.
 *
 * So new state is written base64-encoded. That alphabet contains no `<`, `>`, `&`
 * or quotes, which means no amount of HTML mangling can alter it. The plain-JSON
 * path stays for state written before this change.
 */
export function extractJson(content: string): string {
  const stripped = content
    .replace(/<[^>]+>/g, '')
    .replace(/&quot;/g, '"')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&#(\d+);/g, (_, code: string) => String.fromCharCode(Number(code)))
    .replace(/&nbsp;/g, ' ')
    .trim();

  const marked = stripped.indexOf(B64_MARKER);
  if (marked !== -1) {
    // Everything after the marker that is still base64. Whitespace is dropped
    // because an HTML wrapper is free to insert line breaks anywhere.
    const payload = stripped.slice(marked + B64_MARKER.length).replace(/\s+/g, '');
    return Buffer.from(payload, 'base64').toString('utf8');
  }

  const trimmed = content.trim();
  return trimmed.startsWith('{') ? trimmed : stripped;
}

export async function saveState(token: string, state: RoutineState): Promise<void> {
  const json = JSON.stringify(state);
  // See extractJson: base64 so no HTML round-trip can corrupt it.
  const encoded = B64_MARKER + Buffer.from(json, 'utf8').toString('base64');

  const file = process.env.STEWARD_STATE_FILE;
  if (file) {
    const { writeFile } = await import('node:fs/promises');
    await writeFile(file, json, 'utf8');
    return;
  }

  const folderId = await ensureFolder(token);
  const existing = await findStateMessage(token, folderId);
  const body = { subject: STATE_SUBJECT, body: { contentType: 'text', content: encoded } };

  if (existing) {
    await graph(token, `/me/messages/${existing.id}`, {
      method: 'PATCH',
      body: JSON.stringify(body),
    });
    return;
  }
  await graph(token, `/me/mailFolders/${folderId}/messages`, {
    method: 'POST',
    body: JSON.stringify(body),
  });
}
