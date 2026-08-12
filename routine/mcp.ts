/**
 * MCP server, so she can ask about her mailbox in Claude instead of waiting for a
 * report. `npm run mcp` - though normally Claude Desktop launches it, not a human.
 *
 * The two scheduled routines are push: something fires, a report appears. This is
 * the pull side. It answers the questions a schedule cannot anticipate - "catch me
 * up, I was out Tuesday", "what did I tell the board about the audit?", "did I
 * ever reply to that bequest question?"
 *
 * All the analysis is imported. `src/digest/` and `src/classify/` are already pure
 * functions over fetched data, so nothing here re-decides anything: the same
 * "waiting on you" rules, the same category descriptions, the same confidence
 * gate.
 *
 * TRANSPORT DISCIPLINE: stdout carries JSON-RPC and nothing else. One stray
 * console.log to stdout corrupts the stream and the server dies in a way that
 * looks, from Claude, like the tools simply not existing. Every diagnostic here
 * goes to stderr, which Claude Desktop collects into its log file.
 */

import { createInterface } from 'node:readline';
import { stdin, stdout, stderr } from 'node:process';

import { loadEnv } from './env.js';
import { assertConfigured, redeemRefreshToken } from './auth.js';
import { loadState, saveState, type RoutineState } from './store.js';
import { mergeCategories } from './merge.js';

import { resolveWindow } from '../src/digest/window.js';
import { buildDigest } from '../src/digest/digest.js';
import { DEFAULT_IGNORED_SENDERS } from '../src/digest/senders.js';
import { addDays, describeMoment, assertValidZone } from '../src/digest/zone.js';
import {
  fetchCalendar,
  fetchInboxSince,
  fetchMailbox,
  fetchSentSince,
} from '../src/digest/fetch.js';
import {
  eventInstant,
  fromAddress,
  fromName,
  receivedAt,
  type DigestEvent,
  type DigestMessage,
} from '../src/digest/types.js';

import { applyCategories, listRecentInbox, type CategoryUpdate } from '../src/graph.js';
import { needsClassification } from '../src/engine.js';
import { upsertSenderRule } from '../src/classify/rules.js';
import { categoryById, selectableCategories, NEEDS_REVIEW_ID } from '../src/taxonomy.js';

const SERVER_NAME = 'outlook-sorter';
const SERVER_VERSION = '1.0.0';

/** Protocol version to fall back on when a client doesn't state one. */
const FALLBACK_PROTOCOL = '2025-06-18';

const DEFAULT_TIMEZONE = 'America/Denver';
const DEFAULT_GRACE_HOURS = 18;

function log(message: string): void {
  stderr.write(`[mcp] ${message}\n`);
}

// ---------------------------------------------------------------------------
// Message references
// ---------------------------------------------------------------------------

/**
 * Short handles for messages, so results can say "m3" and she can say "read m3".
 *
 * Graph message ids are 150-plus characters of base64. Putting them in tool output
 * burns context, invites transcription errors, and means the model is copying an
 * opaque blob between calls. A per-session counter is smaller and safer: a stale
 * ref fails to resolve rather than addressing some other message.
 */
class RefStore {
  private readonly byRef = new Map<string, DigestMessage>();
  private counter = 0;

  put(message: DigestMessage): string {
    const ref = `m${++this.counter}`;
    this.byRef.set(ref, message);
    return ref;
  }

  get(ref: string): DigestMessage | undefined {
    return this.byRef.get(ref.trim().toLowerCase());
  }
}

// ---------------------------------------------------------------------------
// Session: token, state, and the mailbox's own settings
// ---------------------------------------------------------------------------

interface Session {
  token: string;
  state: RoutineState;
  mailbox: { address: string; name: string };
  addresses: string[];
  timeZone: string;
  graceMs: number;
  ignoredPatterns: string[];
}

let cached: { session: Session; expiresAt: number } | null = null;

/**
 * A Graph access token lasts about an hour, and this process can stay alive for
 * days while Claude Desktop is open. Re-authenticating per call would be wasteful;
 * never re-authenticating would fail silently once the token aged out.
 */
async function session(): Promise<Session> {
  if (cached && Date.now() < cached.expiresAt) return cached.session;

  await loadEnv();
  assertConfigured();

  const refreshToken = process.env.STEWARD_REFRESH_TOKEN;
  if (!refreshToken) {
    throw new Error(
      'This mailbox has not been set up yet. Run `npm run setup` in the project folder.',
    );
  }

  const tokens = await redeemRefreshToken(refreshToken);
  const state = await loadState(tokens.accessToken);
  state.routine = { ...state.routine, refreshToken: tokens.refreshToken };

  const mailbox = await fetchMailbox(tokens.accessToken);
  const timeZone =
    process.env.STEWARD_TIMEZONE?.trim() || state.weekly?.timeZone || DEFAULT_TIMEZONE;
  assertValidZone(timeZone);

  const extra = (process.env.STEWARD_ALSO_ADDRESSED_AS ?? '')
    .split(',')
    .map((a) => a.trim().toLowerCase())
    .filter(Boolean);

  const built: Session = {
    token: tokens.accessToken,
    state,
    mailbox,
    addresses: [...new Set([mailbox.address, ...extra].filter(Boolean))],
    timeZone,
    graceMs:
      (Number(process.env.STEWARD_WAITING_GRACE_HOURS ?? '') || DEFAULT_GRACE_HOURS) * 3_600_000,
    ignoredPatterns: [
      ...DEFAULT_IGNORED_SENDERS,
      ...(process.env.STEWARD_IGNORED_SENDERS ?? '')
        .split(',')
        .map((p) => p.trim())
        .filter(Boolean),
    ],
  };

  // Refreshed well before the hour is up, so a long-lived session never trips
  // over an expiry mid-answer.
  cached = { session: built, expiresAt: Date.now() + 45 * 60_000 };
  return built;
}

// ---------------------------------------------------------------------------
// Shared data loading
// ---------------------------------------------------------------------------

const INBOX_LOOKBACK_DAYS = 45;
const SENT_LOOKBACK_DAYS = 21;

async function loadWeek(s: Session, now: Date) {
  const window = resolveWindow(now, s.timeZone, undefined);

  const [inbox, sent, events] = await Promise.all([
    fetchInboxSince(s.token, addDays(window.start, -INBOX_LOOKBACK_DAYS, s.timeZone)),
    fetchSentSince(s.token, addDays(window.start, -SENT_LOOKBACK_DAYS, s.timeZone)),
    fetchCalendar(s.token, window.start, window.nextEnd, s.timeZone),
  ]);

  const digest = buildDigest({ inbox, sent, events }, window, {
    mailbox: s.mailbox.address,
    addresses: s.addresses,
    timeZone: s.timeZone,
    now,
    graceMs: s.graceMs,
    ignoredPatterns: s.ignoredPatterns,
  });

  return { window, inbox, sent, events, digest };
}

// ---------------------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------------------

function describeEvent(event: DigestEvent, timeZone: string): string {
  const start = eventInstant(event.start, timeZone);
  const when = start ? describeMoment(start, timeZone) : 'time unknown';
  const who = event.organizer?.emailAddress?.name?.trim();
  return `${when} — ${event.subject?.trim() || '(no subject)'}${who ? ` (organizer: ${who})` : ''}`;
}

function messageLine(m: DigestMessage, ref: string, timeZone: string): string {
  return `[${ref}] ${fromName(m)} <${fromAddress(m)}> — "${m.subject?.trim() || '(no subject)'}" — ${describeMoment(receivedAt(m), timeZone)}`;
}

// ---------------------------------------------------------------------------
// Tools
// ---------------------------------------------------------------------------

interface ToolContext {
  refs: RefStore;
}

type ToolHandler = (args: Record<string, unknown>, ctx: ToolContext) => Promise<string>;

function toolDefinitions() {
  const object = (properties: Record<string, unknown>, required: string[] = []) => ({
    type: 'object',
    properties,
    ...(required.length > 0 ? { required } : {}),
  });

  return [
    {
      name: 'whats_waiting',
      description:
        'Emails still waiting on a reply: sent directly to her (not CC\'d), from a person rather ' +
        'than a mailing list, with nothing sent back in that thread since. Oldest first. Use this ' +
        'for "who is waiting on me", "what am I forgetting", "anything I have left hanging".',
      inputSchema: object({
        min_days_old: {
          type: 'integer',
          description: 'Only show items older than this many days. Default 2.',
        },
      }),
    },
    {
      name: 'weekly_summary',
      description:
        'The full weekly picture: what is waiting, what she flagged, what went unread, meetings ' +
        'and mail sent, unanswered invitations, and next week. Use for "how was my week" or ' +
        '"catch me up".',
      inputSchema: object({}),
    },
    {
      name: 'whats_next_week',
      description:
        'The coming week\'s calendar plus any meeting invitations she has not responded to, and ' +
        'the hours already committed.',
      inputSchema: object({}),
    },
    {
      name: 'find_mail',
      description:
        'Search her mail by sender, subject text, or age. Searches the inbox, and sent mail too ' +
        'when include_sent is set — useful for "what did I say to them?".',
      inputSchema: object({
        from: {
          type: 'string',
          description: 'Match sender name or address, partial and case-insensitive.',
        },
        subject: {
          type: 'string',
          description: 'Match subject text, partial and case-insensitive.',
        },
        days: { type: 'integer', description: 'How far back to look. Default 30, maximum 90.' },
        unread_only: { type: 'boolean', description: 'Only messages she has not opened.' },
        include_sent: { type: 'boolean', description: 'Include mail she sent.' },
      }),
    },
    {
      name: 'read_message',
      description:
        'The full text of one message, using a reference like "m3" from an earlier result. Use ' +
        'before drafting a reply, or when the subject line is not enough to answer her question.',
      inputSchema: object(
        { ref: { type: 'string', description: 'A message reference such as "m3".' } },
        ['ref'],
      ),
    },
    {
      name: 'list_categories',
      description:
        'The labels available for sorting, with a description of what belongs in each. Read this ' +
        'before suggesting labels so the descriptions guide the choice.',
      inputSchema: object({}),
    },
    {
      name: 'suggest_labels',
      description:
        'Fetches recent unlabelled mail so you can propose a category for each. THIS CHANGES ' +
        'NOTHING — it only reads. Decide a label for each message using the category ' +
        'descriptions, show her the proposals, and only call apply_labels once she agrees. If a ' +
        'message genuinely fits no category, or two fit equally well, propose "needs-review" ' +
        'rather than guessing.',
      inputSchema: object({
        count: { type: 'integer', description: 'How many messages to fetch. Default 25, maximum 50.' },
      }),
    },
    {
      name: 'apply_labels',
      description:
        'Writes labels to specific messages. ONLY call this after showing her the proposed labels ' +
        'and getting explicit agreement. Never call it in the same turn as suggest_labels. Labels ' +
        'are reversible and nothing is moved or deleted, but she should always see what will ' +
        'change first.',
      inputSchema: object(
        {
          labels: {
            type: 'array',
            description: 'The messages to label and the category for each.',
            items: {
              type: 'object',
              properties: {
                ref: { type: 'string', description: 'Message reference such as "m3".' },
                category: { type: 'string', description: 'Category id, e.g. "grants" or "donors".' },
              },
              required: ['ref', 'category'],
            },
          },
        },
        ['labels'],
      ),
    },
  ];
}

const handlers: Record<string, ToolHandler> = {
  async whats_waiting(args, ctx) {
    const s = await session();
    const minDays = clampInt(args.min_days_old, 2, 0, 365);
    const { digest } = await loadWeek(s, new Date());

    const items = digest.waiting.filter((w) => w.ageDays >= minDays);
    if (items.length === 0) {
      return minDays > 0
        ? `Nothing has been waiting more than ${minDays} day(s). Her inbox is clear on that front.`
        : 'Nothing is waiting on a reply.';
    }

    const lines = items.map((w) => {
      const ref = ctx.refs.put(w.message);
      const age = w.ageDays === 1 ? '1 day' : `${w.ageDays} days`;
      const preview = w.message.bodyPreview?.replace(/\s+/g, ' ').slice(0, 200) ?? '';
      return `[${ref}] ${fromName(w.message)} <${fromAddress(w.message)}>\n  "${w.message.subject?.trim() || '(no subject)'}"\n  waiting ${age} (since ${describeMoment(receivedAt(w.message), s.timeZone)})\n  ${preview}`;
    });

    return `${items.length} waiting on a reply, oldest first:\n\n${lines.join('\n\n')}`;
  },

  async weekly_summary(_args, ctx) {
    const s = await session();
    const { digest, window, events } = await loadWeek(s, new Date());
    const out: string[] = [`Week of ${window.description} — ${digest.headline}`];

    if (digest.waiting.length > 0) {
      out.push(
        `\nWAITING ON A REPLY (${digest.waiting.length})\n` +
          digest.waiting
            .map((w) => `  ${messageLine(w.message, ctx.refs.put(w.message), s.timeZone)} — ${w.ageDays}d`)
            .join('\n'),
      );
    }

    if (digest.flagged.length > 0) {
      out.push(
        `\nSHE FLAGGED THESE (${digest.flagged.length})\n` +
          digest.flagged
            .map((m) => `  ${messageLine(m, ctx.refs.put(m), s.timeZone)}`)
            .join('\n'),
      );
    }

    if (digest.calendar.nextWeekUnanswered.length > 0) {
      out.push(
        `\nINVITATIONS NEEDING AN RSVP (${digest.calendar.nextWeekUnanswered.length})\n` +
          digest.calendar.nextWeekUnanswered
            .map((e) => `  ${describeEvent(e, s.timeZone)}`)
            .join('\n'),
      );
    }

    out.push(
      `\nUNREAD\n  ${digest.unread.total} total, ${digest.unread.staleCount} older than a week` +
        (digest.unread.people.length > 0
          ? `\n  From people: ${digest.unread.people
              .slice(0, 6)
              .map((g) => `${g.name} (${g.count})`)
              .join(', ')}`
          : '') +
        (digest.unread.automated.length > 0
          ? `\n  Bulk: ${digest.unread.automated.reduce((n, g) => n + g.count, 0)} from ${digest.unread.automated.length} sender(s)`
          : ''),
    );

    const r = digest.review;
    out.push(
      `\nACTIVITY (what Outlook can see — not what was achieved, and attendance is only an RSVP)\n` +
        `  ${r.meetingsHeld} meetings, ${r.meetingHours}h, ${r.meetingsOrganized} organized by her\n` +
        `  ${r.emailsSent} emails sent to ${r.peopleWrittenTo} people across ${r.threadsAdvanced} threads` +
        (r.busiestDay ? `\n  Busiest day: ${r.busiestDay}` : ''),
    );

    out.push(
      `\nNEXT WEEK\n  ${digest.calendar.nextWeekHours}h committed across ${digest.calendar.nextWeek.length} entries`,
    );

    if (events.length === 0) {
      out.push(
        '\nNote: the calendar came back empty. That more often means a missing Calendars.Read permission than a week with no meetings.',
      );
    }

    return out.join('\n');
  },

  async whats_next_week() {
    const s = await session();
    const { digest } = await loadWeek(s, new Date());
    const c = digest.calendar;

    const parts = [
      `Next week: ${c.nextWeekHours}h committed across ${c.nextWeek.length} calendar entries.`,
    ];

    if (c.nextWeekUnanswered.length > 0) {
      parts.push(
        `\nNot yet answered (${c.nextWeekUnanswered.length}):\n` +
          c.nextWeekUnanswered.map((e) => `  ${describeEvent(e, s.timeZone)}`).join('\n'),
      );
    } else {
      parts.push('\nNo outstanding invitations.');
    }

    if (c.nextWeek.length > 0) {
      parts.push(
        `\nThe week:\n${c.nextWeek.map((e) => `  ${describeEvent(e, s.timeZone)}`).join('\n')}`,
      );
    }

    return parts.join('\n');
  },

  async find_mail(args, ctx) {
    const s = await session();
    const days = clampInt(args.days, 30, 1, 90);
    const from = String(args.from ?? '').trim().toLowerCase();
    const subject = String(args.subject ?? '').trim().toLowerCase();
    const unreadOnly = args.unread_only === true;
    const includeSent = args.include_sent === true;

    if (!from && !subject && !unreadOnly) {
      return 'Give at least one of: from, subject, or unread_only. An unfiltered search would just dump her inbox.';
    }

    const since = addDays(new Date(), -days, s.timeZone);
    const [inbox, sent] = await Promise.all([
      fetchInboxSince(s.token, since),
      includeSent ? fetchSentSince(s.token, since) : Promise.resolve([]),
    ]);

    const matches = (m: DigestMessage, outbound: boolean): boolean => {
      if (unreadOnly && (m.isRead || outbound)) return false;
      if (from) {
        // On sent mail, "from: Tessa" plainly means the recipient - matching the
        // sender there would only ever return her own address.
        const haystack = outbound
          ? (m.toRecipients ?? [])
              .map((r) => `${r.emailAddress?.name ?? ''} ${r.emailAddress?.address ?? ''}`)
              .join(' ')
              .toLowerCase()
          : `${fromName(m)} ${fromAddress(m)}`.toLowerCase();
        if (!haystack.includes(from)) return false;
      }
      if (subject && !(m.subject ?? '').toLowerCase().includes(subject)) return false;
      return true;
    };

    const found = [
      ...inbox.filter((m) => matches(m, false)).map((m) => ({ m, outbound: false })),
      ...sent.filter((m) => matches(m, true)).map((m) => ({ m, outbound: true })),
    ]
      .sort((a, b) => receivedAt(b.m).getTime() - receivedAt(a.m).getTime())
      .slice(0, 40);

    if (found.length === 0) return `Nothing matched in the last ${days} days.`;

    const lines = found.map(({ m, outbound }) => {
      const ref = ctx.refs.put(m);
      const direction = outbound ? 'sent' : 'received';
      const who = outbound
        ? (m.toRecipients ?? [])
            .map((r) => r.emailAddress?.name || r.emailAddress?.address)
            .filter(Boolean)
            .join(', ')
        : `${fromName(m)} <${fromAddress(m)}>`;
      return `[${ref}] ${direction} — ${who} — "${m.subject?.trim() || '(no subject)'}" — ${describeMoment(receivedAt(m), s.timeZone)}`;
    });

    return `${found.length} match(es) in the last ${days} days:\n\n${lines.join('\n')}`;
  },

  async read_message(args, ctx) {
    const s = await session();
    const ref = String(args.ref ?? '').trim();
    const message = ctx.refs.get(ref);

    if (!message) {
      return `No message is known by the reference "${ref}". References come from an earlier result in this conversation and are not stable across restarts - run the search again.`;
    }

    // Fetched fresh rather than served from the cached summary: the search results
    // carry only a preview, and this tool exists precisely because the preview was
    // not enough.
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${message.id}?$select=subject,from,toRecipients,ccRecipients,receivedDateTime,body,categories`,
      { headers: { Authorization: `Bearer ${s.token}` } },
    );
    if (!res.ok) return `Could not read that message (Graph returned ${res.status}).`;

    const full = (await res.json()) as DigestMessage & { body?: { content?: string } };
    const text = (full.body?.content ?? '')
      .replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/\s+/g, ' ')
      .trim();

    const to = (full.toRecipients ?? [])
      .map((r) => r.emailAddress?.address)
      .filter(Boolean)
      .join(', ');

    return [
      `From: ${fromName(full)} <${fromAddress(full)}>`,
      `To: ${to}`,
      `Date: ${describeMoment(receivedAt(full), s.timeZone)}`,
      `Subject: ${full.subject?.trim() || '(no subject)'}`,
      (full.categories ?? []).length > 0 ? `Labels: ${(full.categories ?? []).join(', ')}` : '',
      '',
      text.slice(0, 6000) || '(no readable body)',
    ]
      .filter((line) => line !== '')
      .join('\n');
  },

  async list_categories() {
    const s = await session();
    const lines = selectableCategories(s.state.taxonomy).map(
      (c) => `${c.id} ("${c.name}")\n  ${c.description}`,
    );
    return `${lines.length} categories. Use the id when calling apply_labels.\n\nAlso available: needs-review — for anything that genuinely fits nothing, or fits two equally well.\n\n${lines.join('\n\n')}`;
  },

  async suggest_labels(args, ctx) {
    const s = await session();
    const count = clampInt(args.count, 25, 1, 50);

    const recent = await listRecentInbox(s.token, Math.min(count * 3, 120), true);
    const unlabelled = needsClassification(
      recent,
      s.state.taxonomy,
      s.state.settings.generation,
    ).slice(0, count);

    if (unlabelled.length === 0) return 'Everything recent is already labelled. Nothing to propose.';

    const lines = unlabelled.map((m) => {
      // MailSummary and DigestMessage are different shapes; the ref store holds the
      // latter, so this adapts rather than duplicating the store.
      const ref = ctx.refs.put({
        id: m.id,
        subject: m.subject,
        receivedDateTime: m.received,
        from: { emailAddress: { address: m.from, name: m.fromName } },
        categories: m.categories,
      });
      return `[${ref}] ${m.fromName} <${m.from}> — "${m.subject}"${m.preview ? `\n  ${m.preview.replace(/\s+/g, ' ').slice(0, 250)}` : ''}`;
    });

    return `${unlabelled.length} unlabelled message(s). Nothing has been changed. Propose a category id for each, show her, and call apply_labels only once she agrees.\n\n${lines.join('\n\n')}`;
  },

  async apply_labels(args, ctx) {
    const s = await session();
    const requested = Array.isArray(args.labels) ? args.labels : [];
    if (requested.length === 0) return 'No labels were given, so nothing was changed.';

    const updates: CategoryUpdate[] = [];
    const applied: { ref: string; messageId: string; sender: string; categoryId: string; name: string }[] = [];
    const problems: string[] = [];

    for (const raw of requested) {
      const entry = raw as { ref?: string; category?: string };
      const ref = String(entry.ref ?? '').trim();
      const categoryId = String(entry.category ?? '').trim();

      const message = ctx.refs.get(ref);
      if (!message) {
        problems.push(`${ref}: unknown reference`);
        continue;
      }

      const category = categoryById(s.state.taxonomy, categoryId);
      if (!category) {
        problems.push(`${ref}: "${categoryId}" is not a category id`);
        continue;
      }

      updates.push({
        messageId: message.id,
        categories: mergeCategories(message.categories ?? [], s.state.taxonomy, category.name),
        provenance: {
          categoryId,
          // She approved this one explicitly, which is the strongest signal
          // available - stronger than anything the classifier infers.
          confidence: 1,
          generation: s.state.settings.generation,
        },
      });
      applied.push({
        ref,
        messageId: message.id,
        sender: fromAddress(message),
        categoryId,
        name: category.name,
      });
    }

    if (updates.length === 0) {
      return `Nothing was changed. Problems:\n${problems.map((p) => `  ${p}`).join('\n')}`;
    }

    const result = await applyCategories(s.token, updates);
    const succeeded = new Set(result.succeeded);

    // Her explicit approval counts as a confirmation, which is what earns a sender
    // a native Outlook rule later. Promotion itself is left to the scheduled sweep
    // rather than done here: writing a permanent mailbox rule as a side effect of
    // a chat message is more than she agreed to.
    let learned = 0;
    for (const entry of applied) {
      if (!succeeded.has(entry.messageId)) continue;
      if (!entry.sender || entry.categoryId === NEEDS_REVIEW_ID) continue;
      s.state.senderRules = upsertSenderRule(
        s.state.senderRules,
        entry.sender,
        entry.categoryId,
        true,
      );
      learned++;
    }
    if (learned > 0) s.state.settings.generation++;
    await saveState(s.token, s.state);

    const lines = applied
      .filter((entry) => succeeded.has(entry.messageId))
      .map((entry) => `  ${entry.ref} → ${entry.name}`);

    return [
      `Labelled ${result.succeeded.length} message(s):`,
      ...lines,
      learned > 0 ? `\nLearned ${learned} sender(s), which will be applied automatically from now on.` : '',
      result.failed.length > 0 ? `\n${result.failed.length} failed to update.` : '',
      problems.length > 0 ? `\nSkipped:\n${problems.map((p) => `  ${p}`).join('\n')}` : '',
      '\nAll of this is reversible - the labels are ordinary Outlook categories.',
    ]
      .filter(Boolean)
      .join('\n');
  },
};

function clampInt(value: unknown, fallback: number, min: number, max: number): number {
  const n = Math.floor(Number(value));
  if (!Number.isFinite(n)) return fallback;
  return Math.max(min, Math.min(max, n));
}

// ---------------------------------------------------------------------------
// JSON-RPC over stdio
// ---------------------------------------------------------------------------

interface RpcRequest {
  jsonrpc?: string;
  id?: string | number | null;
  method?: string;
  params?: Record<string, unknown>;
}

const refs = new RefStore();

function send(payload: Record<string, unknown>): void {
  stdout.write(`${JSON.stringify(payload)}\n`);
}

function reply(id: string | number, result: unknown): void {
  send({ jsonrpc: '2.0', id, result });
}

function replyError(id: string | number, code: number, message: string): void {
  send({ jsonrpc: '2.0', id, error: { code, message } });
}

async function handle(request: RpcRequest): Promise<void> {
  const { id, method } = request;

  // A notification has no id and must never be answered.
  if (id === undefined || id === null) {
    if (method === 'notifications/initialized') log('client initialized');
    return;
  }

  switch (method) {
    case 'initialize': {
      const requested = (request.params?.protocolVersion as string) || FALLBACK_PROTOCOL;
      log(`initialize (protocol ${requested})`);
      reply(id, {
        protocolVersion: requested,
        capabilities: { tools: {} },
        serverInfo: { name: SERVER_NAME, version: SERVER_VERSION },
      });
      return;
    }

    case 'ping':
      reply(id, {});
      return;

    case 'tools/list':
      reply(id, { tools: toolDefinitions() });
      return;

    case 'tools/call': {
      const name = String(request.params?.name ?? '');
      const handler = handlers[name];
      if (!handler) {
        replyError(id, -32602, `No tool named "${name}".`);
        return;
      }

      const args = (request.params?.arguments as Record<string, unknown>) ?? {};
      log(`tools/call ${name}`);

      try {
        const text = await handler(args, { refs });
        reply(id, { content: [{ type: 'text', text }] });
      } catch (err) {
        // Returned as a tool error rather than a protocol error: the model should
        // see what went wrong and be able to tell her, instead of the call
        // vanishing.
        const message = err instanceof Error ? err.message : String(err);
        log(`tool ${name} failed: ${message}`);
        reply(id, { content: [{ type: 'text', text: `That failed: ${message}` }], isError: true });
      }
      return;
    }

    default:
      replyError(id, -32601, `Unsupported method "${method}".`);
  }
}

log(`${SERVER_NAME} ${SERVER_VERSION} ready on stdio`);

const lines = createInterface({ input: stdin });

// Serialized rather than concurrent: two tool calls sharing one refs store and one
// state object could interleave a read of senderRules with a write of it, and the
// second save would silently drop the first one's learning.
let queue: Promise<void> = Promise.resolve();

lines.on('line', (line) => {
  const trimmed = line.trim();
  if (trimmed === '') return;

  let request: RpcRequest;
  try {
    request = JSON.parse(trimmed) as RpcRequest;
  } catch {
    log(`ignored unparseable line (${trimmed.length} chars)`);
    return;
  }

  queue = queue.then(() =>
    handle(request).catch((err) => {
      log(`handler crashed: ${err instanceof Error ? err.message : String(err)}`);
    }),
  );
});

lines.on('close', () => {
  // Drain before exiting. Calling process.exit() here directly abandons whatever
  // is still in flight, so the last request of a session gets no response at all -
  // which from the client looks like the server crashing mid-answer.
  log('stdin closed, finishing in-flight work');
  queue.then(() => {
    log('exiting');
    process.exit(0);
  });
});
