/**
 * Weekly digest driver. `npm run weekly`.
 *
 * The same division of labour as the sorter: this does the deterministic work -
 * resolving which week to report on, fetching the mailbox and calendar, and
 * reducing all of it to counts and short lists - and writes the result to a file.
 * The routine prompt then turns that into prose a person actually reads.
 *
 * Nothing here writes to the mailbox. A digest is read-only by design: it reports
 * on the week, it does not act on it.
 */

import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { loadEnv } from './env.js';
import { assertConfigured, redeemRefreshToken } from './auth.js';
import { loadState, saveState, type RoutineState } from './store.js';
import { resolveWindow } from '../src/digest/window.js';
import { buildDigest } from '../src/digest/digest.js';
import { DEFAULT_IGNORED_SENDERS } from '../src/digest/senders.js';
import { assertValidZone, addDays, describeMoment } from '../src/digest/zone.js';
import {
  fetchCalendar,
  fetchInboxSince,
  fetchMailbox,
  fetchSentSince,
} from '../src/digest/fetch.js';
import {
  eventInstant,
  fromName,
  receivedAt,
  type DigestEvent,
} from '../src/digest/types.js';

// fileURLToPath, not URL.pathname: on Windows the latter yields "/C:/...",
// which every path built from it then fails to resolve.
const HERE = dirname(fileURLToPath(import.meta.url));
const DIGEST_PATH = resolve(HERE, '.local/digest.json');

/**
 * Default zone.
 *
 * The foundation is in Durango, Colorado. Getting this wrong shifts the whole
 * reporting window, so it is a stated default rather than the runner's local
 * time - a cloud runner is almost certainly on UTC.
 */
const DEFAULT_TIMEZONE = 'America/Denver';

/** Hours a message may sit before it counts as waiting. This morning's mail is not a failure. */
const DEFAULT_GRACE_HOURS = 18;

/**
 * How far back to read sent mail, relative to the start of the reporting window.
 *
 * Wider than the window on purpose: a reply sent on Monday answers a message that
 * arrived the previous Friday, and if the sent history is too shallow that
 * message comes back looking unanswered.
 */
const SENT_LOOKBACK_DAYS = 21;

/** How far back to read the inbox. Waiting mail can be older than the reporting week. */
const INBOX_LOOKBACK_DAYS = 45;

interface WeeklyConfig {
  timeZone: string;
  graceMs: number;
  addresses: string[];
  ignoredPatterns: string[];
}

function readConfig(mailbox: string, state: RoutineState): WeeklyConfig {
  const timeZone = process.env.STEWARD_TIMEZONE?.trim() || state.weekly?.timeZone || DEFAULT_TIMEZONE;
  assertValidZone(timeZone);

  const graceHours = Number(process.env.STEWARD_WAITING_GRACE_HOURS ?? '') || DEFAULT_GRACE_HOURS;

  // Aliases are additive to the signed-in address, never a replacement: dropping
  // the primary would make every message fail the "was she asked?" test.
  const extra = (process.env.STEWARD_ALSO_ADDRESSED_AS ?? state.weekly?.alsoAddressedAs?.join(',') ?? '')
    .split(',')
    .map((a) => a.trim().toLowerCase())
    .filter(Boolean);

  const ignoredExtra = (process.env.STEWARD_IGNORED_SENDERS ?? '')
    .split(',')
    .map((p) => p.trim())
    .filter(Boolean);

  return {
    timeZone,
    graceMs: graceHours * 3_600_000,
    addresses: [...new Set([mailbox, ...extra].filter(Boolean))],
    ignoredPatterns: [...DEFAULT_IGNORED_SENDERS, ...ignoredExtra],
  };
}

async function main(): Promise<void> {
  await loadEnv();
  assertConfigured();

  const refreshToken = process.env.STEWARD_REFRESH_TOKEN ?? (await readBootstrapToken());
  if (!refreshToken) {
    throw new Error('No refresh token available. Run `npm run setup` once.');
  }

  const tokens = await redeemRefreshToken(refreshToken);
  const state = await loadState(tokens.accessToken);
  state.routine = { ...state.routine, refreshToken: tokens.refreshToken };

  const mailbox = await fetchMailbox(tokens.accessToken);
  const config = readConfig(mailbox.address, state);

  const now = new Date();
  const window = resolveWindow(now, config.timeZone, state.weekly?.lastReportedWeek);

  const inboxSince = addDays(window.start, -INBOX_LOOKBACK_DAYS, config.timeZone);
  const sentSince = addDays(window.start, -SENT_LOOKBACK_DAYS, config.timeZone);

  const [inbox, sent, events] = await Promise.all([
    fetchInboxSince(tokens.accessToken, inboxSince),
    fetchSentSince(tokens.accessToken, sentSince),
    fetchCalendar(tokens.accessToken, window.start, window.nextEnd, config.timeZone),
  ]);

  const digest = buildDigest({ inbox, sent, events }, window, {
    mailbox: mailbox.address,
    addresses: config.addresses,
    timeZone: config.timeZone,
    now,
    graceMs: config.graceMs,
    ignoredPatterns: config.ignoredPatterns,
  });

  // Rendered into a shape meant to be read by a model and by a person looking at
  // the file: absolute timestamps become "Tue 12 Aug, 14:30" here rather than in
  // the prompt, because the zone is known here and guessing it there would be a
  // silent source of wrong times.
  const output = {
    week: window.label,
    covering: window.description,
    catchUp: window.catchUp,
    mailbox: `${mailbox.name} <${mailbox.address}>`.trim(),
    timeZone: config.timeZone,
    generatedAt: describeMoment(now, config.timeZone),
    headline: digest.headline,
    hasAnythingToSay: digest.hasAnythingToSay,

    waiting: digest.waiting.map((w) => ({
      from: `${fromName(w.message)} <${w.message.from?.emailAddress?.address ?? ''}>`,
      subject: w.message.subject?.trim() || '(no subject)',
      received: describeMoment(receivedAt(w.message), config.timeZone),
      daysWaiting: w.ageDays,
      preview: w.message.bodyPreview?.replace(/\s+/g, ' ').slice(0, 300) ?? '',
      link: w.message.webLink ?? '',
    })),

    flagged: digest.flagged.map((m) => ({
      from: fromName(m),
      subject: m.subject?.trim() || '(no subject)',
      received: describeMoment(receivedAt(m), config.timeZone),
      link: m.webLink ?? '',
    })),

    unread: {
      total: digest.unread.total,
      olderThanAWeek: digest.unread.staleCount,
      fromPeople: digest.unread.people.map((g) => ({
        who: `${g.name} <${g.address}>`,
        count: g.count,
        oldest: describeMoment(new Date(g.oldest), config.timeZone),
        subjects: g.subjects,
      })),
      bulk: digest.unread.automated.map((g) => ({ who: g.name || g.address, count: g.count })),
    },

    activity: {
      meetingsHeld: digest.review.meetingsHeld,
      meetingHours: digest.review.meetingHours,
      meetingsSheOrganized: digest.review.meetingsOrganized,
      emailsSent: digest.review.emailsSent,
      peopleWrittenTo: digest.review.peopleWrittenTo,
      threadsAdvanced: digest.review.threadsAdvanced,
      busiestDay: digest.review.busiestDay,
      biggestMeetings: digest.review.highlights.map((e) => describeEvent(e, config.timeZone)),
    },

    calendar: {
      invitationsNeverAnswered: digest.calendar.unanswered.map((e) =>
        describeEvent(e, config.timeZone),
      ),
      declined: digest.calendar.declined.map((e) => describeEvent(e, config.timeZone)),
      nextWeekHoursCommitted: digest.calendar.nextWeekHours,
      nextWeek: digest.calendar.nextWeek.map((e) => describeEvent(e, config.timeZone)),
      nextWeekNeedsRsvp: digest.calendar.nextWeekUnanswered.map((e) =>
        describeEvent(e, config.timeZone),
      ),
    },

    counts: {
      inboxMessagesRead: inbox.length,
      sentMessagesRead: sent.length,
      calendarEntriesRead: events.length,
    },
  };

  await mkdir(dirname(DIGEST_PATH), { recursive: true });
  await writeFile(DIGEST_PATH, JSON.stringify(output, null, 2), 'utf8');

  // Recorded only once the digest exists, so a failed run does not burn the week
  // and suppress the catch-up that is meant to cover it.
  state.weekly = {
    ...state.weekly,
    lastReportedWeek: window.label,
    lastRunAt: new Date().toISOString(),
    timeZone: config.timeZone,
  };
  await saveState(tokens.accessToken, state);

  console.log(`week ${window.label} (${window.description})${window.catchUp ? ' [catch-up]' : ''}`);
  console.log(`headline: ${digest.headline}`);
  console.log(
    `waiting ${digest.waiting.length}, flagged ${digest.flagged.length}, unread ${digest.unread.total}, RSVPs due ${digest.calendar.nextWeekUnanswered.length}`,
  );
  console.log(`digest written to ${DIGEST_PATH}`);
}

interface EventLine {
  what: string;
  when: string;
  organizer: string;
  attendees?: number;
}

function describeEvent(event: DigestEvent, timeZone: string): EventLine {
  const start = eventInstant(event.start, timeZone);
  // Meeting rooms and equipment are attendees as far as Graph is concerned, and
  // counting them makes a two-person meeting look like a three-person one.
  const humans = (event.attendees ?? []).filter((a) => a.type !== 'resource').length;

  return {
    what: event.subject?.trim() || '(no subject)',
    when: start ? describeMoment(start, timeZone) : 'unknown',
    organizer: event.organizer?.emailAddress?.name?.trim() || '',
    ...(humans > 0 ? { attendees: humans } : {}),
  };
}

async function readBootstrapToken(): Promise<string | null> {
  try {
    return (await readFile(resolve(HERE, '.local/refresh-token'), 'utf8')).trim() || null;
  } catch {
    return null;
  }
}

main().catch((err) => {
  console.error(err instanceof Error ? err.message : String(err));
  process.exit(1);
});
