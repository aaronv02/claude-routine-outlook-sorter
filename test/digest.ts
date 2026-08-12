/**
 * Offline checks for the weekly digest. Imported by test/run.ts.
 *
 * The digest has more ways to be quietly wrong than the sorter does, and none of
 * them announce themselves: a timezone slip moves the week by hours, a To/Cc
 * mistake fills the waiting list with mail nobody asked her about, and a
 * too-shallow sent-mail window makes answered threads look ignored. All three
 * produce a plausible-looking report, which is why they are tested rather than
 * eyeballed.
 */

import { isoLabel, isoWeekStart, instantOf, wallOf, describeRange } from '../src/digest/zone.js';
import { resolveWindow } from '../src/digest/window.js';
import { isAutomatedSender, DEFAULT_IGNORED_SENDERS } from '../src/digest/senders.js';
import { findOpenFlags, findWaiting } from '../src/digest/waiting.js';
import { buildUnread } from '../src/digest/unread.js';
import { buildCalendarGaps } from '../src/digest/calendar.js';
import { buildReview } from '../src/digest/review.js';
import { buildDigest } from '../src/digest/digest.js';
import type { DigestEvent, DigestMessage } from '../src/digest/types.js';

const TZ = 'America/Denver';
const HER = 'director@example.org';
const ALIAS = 'grants@example.org';

type Check = (name: string, fn: () => void) => void;
type Assert = (condition: unknown, message: string) => void;

interface MessageOverrides {
  id?: string;
  from?: string;
  fromName?: string;
  to?: string[];
  cc?: string[];
  subject?: string;
  received?: string;
  sent?: string;
  conversationId?: string;
  isRead?: boolean;
  isDraft?: boolean;
  flagged?: boolean;
}

function msg(over: MessageOverrides = {}): DigestMessage {
  const address = over.from ?? 'someone@outside.org';
  return {
    id: over.id ?? `m-${Math.abs(hash(JSON.stringify(over)))}`,
    conversationId: over.conversationId ?? `c-${address}`,
    subject: over.subject ?? 'A subject',
    bodyPreview: 'Body preview text.',
    receivedDateTime: over.received ?? '2026-08-03T15:00:00Z',
    ...(over.sent ? { sentDateTime: over.sent } : {}),
    isRead: over.isRead ?? true,
    isDraft: over.isDraft ?? false,
    from: { emailAddress: { address, name: over.fromName ?? 'Some One' } },
    toRecipients: (over.to ?? [HER]).map((a) => ({ emailAddress: { address: a } })),
    ccRecipients: (over.cc ?? []).map((a) => ({ emailAddress: { address: a } })),
    flag: { flagStatus: over.flagged ? 'flagged' : 'notFlagged' },
  };
}

/** Stable ids without Math.random, so a failure is reproducible. */
function hash(s: string): number {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0;
  return h;
}

function evt(over: Partial<DigestEvent> & { startLocal?: string; endLocal?: string }): DigestEvent {
  const { startLocal, endLocal, ...rest } = over;
  return {
    id: `e-${Math.abs(hash(JSON.stringify(over)))}`,
    subject: 'A meeting',
    start: { dateTime: startLocal ?? '2026-08-04T09:00:00.0000000', timeZone: TZ },
    end: { dateTime: endLocal ?? '2026-08-04T10:00:00.0000000', timeZone: TZ },
    isAllDay: false,
    isCancelled: false,
    isOrganizer: false,
    responseStatus: { response: 'accepted' },
    attendees: [],
    ...rest,
  };
}

const HOUR = 3_600_000;
const DAY = 24 * HOUR;

export function runDigestChecks(check: Check, assert: Assert): void {
  // -------------------------------------------------------------------------
  // Timezone arithmetic - everything else is built on this
  // -------------------------------------------------------------------------

  check('a wall-clock time resolves to the right instant', () => {
    // 2026-08-04 00:00 in Denver is UTC-6 (MDT), so 06:00Z.
    const at = instantOf({ year: 2026, month: 8, day: 4 }, TZ);
    assert(at.toISOString() === '2026-08-04T06:00:00.000Z', `got ${at.toISOString()}`);
  });

  check('winter and summer offsets are both handled', () => {
    // January is MST (UTC-7), so midnight local is 07:00Z.
    const winter = instantOf({ year: 2026, month: 1, day: 15 }, TZ);
    assert(winter.toISOString() === '2026-01-15T07:00:00.000Z', `got ${winter.toISOString()}`);
  });

  check('the day after a DST transition still starts at local midnight', () => {
    // US DST began 2026-03-08. A single-pass offset guess lands an hour out here,
    // which would move a Monday boundary back into Sunday.
    const after = instantOf({ year: 2026, month: 3, day: 9 }, TZ);
    const w = wallOf(after, TZ);
    assert(w.hour === 0 && w.day === 9, `got day ${w.day} hour ${w.hour}`);
  });

  check('the week starts on Monday local, not Sunday and not UTC', () => {
    // A Sunday evening in Denver is already Monday in UTC. The week must not roll.
    const sundayEvening = new Date('2026-08-10T02:00:00Z'); // Sun 9 Aug, 20:00 MDT
    const start = isoWeekStart(sundayEvening, TZ);
    const w = wallOf(start, TZ);
    assert(w.isoWeekday === 1, `week started on ISO weekday ${w.isoWeekday}`);
    assert(w.day === 3 && w.month === 8, `week started ${w.day}/${w.month}, expected 3/8`);
  });

  check('ISO week labels follow the Thursday rule', () => {
    // 2026-01-01 is a Thursday, so its week is 2026-W01 and starts in December.
    const label = isoLabel(isoWeekStart(new Date('2026-01-01T12:00:00Z'), TZ), TZ);
    assert(label === '2026-W01', `got ${label}`);
  });

  check('a range is described without repeating a shared month', () => {
    const start = instantOf({ year: 2026, month: 8, day: 3 }, TZ);
    const end = instantOf({ year: 2026, month: 8, day: 8 }, TZ);
    const text = describeRange(start, end, TZ);
    assert(text === 'Mon 3 – Fri 7 Aug', `got "${text}"`);
  });

  // -------------------------------------------------------------------------
  // Which week gets reported - the fix for a late run
  // -------------------------------------------------------------------------

  check('a Friday run reports the current week', () => {
    const friday = new Date('2026-08-07T22:00:00Z'); // Fri 7 Aug, 16:00 MDT
    const w = resolveWindow(friday, TZ, undefined);
    assert(w.label === '2026-W32', `got ${w.label}`);
    assert(!w.catchUp, 'called a normal Friday run a catch-up');
  });

  check('a Monday run reports the week that was missed', () => {
    const monday = new Date('2026-08-10T16:00:00Z');
    const w = resolveWindow(monday, TZ, undefined);
    assert(w.catchUp, 'failed to notice the missed week');
    assert(w.label === '2026-W32', `reported ${w.label} instead of the missed week`);
  });

  check('a Monday run does not resend a week already reported', () => {
    const monday = new Date('2026-08-10T16:00:00Z');
    const w = resolveWindow(monday, TZ, '2026-W32');
    assert(!w.catchUp, 'reported a catch-up for a week already sent');
    assert(w.label === '2026-W33', `got ${w.label}`);
  });

  check('the window never claims to cover the future', () => {
    const midweek = new Date('2026-08-12T18:00:00Z'); // Wednesday
    const w = resolveWindow(midweek, TZ, '2026-W32');
    assert(w.end.getTime() <= midweek.getTime(), 'reported time that has not happened');
  });

  // -------------------------------------------------------------------------
  // Automated sender detection - anchored, not substring
  // -------------------------------------------------------------------------

  check('robot addresses are recognized', () => {
    for (const address of [
      'no-reply@example.org',
      'noreply-service@example.org',
      'newsletter@foundation.org',
      'someone@mailchimp.com',
      'bounces-1234@example.org',
      'webinars@candid.org',
      'invitations@example.org',
    ]) {
      assert(
        isAutomatedSender(address, DEFAULT_IGNORED_SENDERS),
        `missed automated sender ${address}`,
      );
    }
  });

  check('a person is never mistaken for a robot by substring', () => {
    // The whole reason matching is anchored: "news@" must not swallow this, and
    // dropping a real person from the waiting list is the worst failure here.
    for (const address of [
      'goodnews@apersonsdomain.org',
      'newsome@example.org',
      'alerta@example.org',
      'automatedteller@example.org',
    ]) {
      assert(
        !isAutomatedSender(address, DEFAULT_IGNORED_SENDERS),
        `treated ${address} as automated`,
      );
    }
  });

  check('info@ and events@ are left alone', () => {
    // At a small nonprofit these are usually staffed by a person who does reply.
    assert(!isAutomatedSender('info@example.org', DEFAULT_IGNORED_SENDERS), 'ignored info@');
    assert(!isAutomatedSender('events@example.org', DEFAULT_IGNORED_SENDERS), 'ignored events@');
  });

  // -------------------------------------------------------------------------
  // Waiting on you - the section with the most ways to be wrong
  // -------------------------------------------------------------------------

  const now = new Date('2026-08-07T22:00:00Z');
  const threeDaysAgo = new Date(now.getTime() - 3 * DAY).toISOString();
  const grace = 18 * HOUR;

  const waitingOf = (inbox: DigestMessage[], sent: DigestMessage[] = []) =>
    findWaiting(inbox, sent, [HER, ALIAS], grace, now, DEFAULT_IGNORED_SENDERS);

  check('unanswered mail addressed to her is waiting', () => {
    const out = waitingOf([msg({ received: threeDaysAgo })]);
    assert(out.length === 1, `expected 1, got ${out.length}`);
    assert(out[0]?.ageDays === 3, `age ${out[0]?.ageDays}`);
  });

  check('being CC\'d is not being asked', () => {
    const out = waitingOf([msg({ received: threeDaysAgo, to: ['someone@else.org'], cc: [HER] })]);
    assert(out.length === 0, 'listed a message she was only copied on');
  });

  check('a role alias counts as being asked', () => {
    const out = waitingOf([msg({ received: threeDaysAgo, to: [ALIAS] })]);
    assert(out.length === 1, 'ignored mail sent to her grants@ alias');
  });

  check('a reply removes it from the list', () => {
    const inbox = [msg({ id: 'in', conversationId: 'thread-1', received: threeDaysAgo })];
    const sent = [
      msg({
        id: 'out',
        conversationId: 'thread-1',
        from: HER,
        sent: new Date(now.getTime() - 2 * DAY).toISOString(),
      }),
    ];
    assert(waitingOf(inbox, sent).length === 0, 'listed a thread she already answered');
  });

  check('a reply sent BEFORE the message does not count as an answer', () => {
    const inbox = [msg({ id: 'in', conversationId: 'thread-2', received: threeDaysAgo })];
    const sent = [
      msg({
        id: 'out',
        conversationId: 'thread-2',
        from: HER,
        sent: new Date(now.getTime() - 5 * DAY).toISOString(),
      }),
    ];
    assert(waitingOf(inbox, sent).length === 1, 'treated an earlier message as a reply');
  });

  check('this morning\'s mail is not a failure', () => {
    const recent = new Date(now.getTime() - 2 * HOUR).toISOString();
    assert(waitingOf([msg({ received: recent })]).length === 0, 'nagged about mail from 2h ago');
  });

  check('a thread collapses to one entry, dated from the first ask', () => {
    const oldest = new Date(now.getTime() - 6 * DAY).toISOString();
    const inbox = [
      msg({ id: 'b', conversationId: 'thread-3', received: threeDaysAgo, subject: 'Following up' }),
      msg({ id: 'a', conversationId: 'thread-3', received: oldest, subject: 'Original ask' }),
    ];
    const out = waitingOf(inbox);
    assert(out.length === 1, `a 2-message thread produced ${out.length} items`);
    assert(out[0]?.ageDays === 6, `measured from the nudge, not the ask: ${out[0]?.ageDays} days`);
  });

  check('newsletters never appear as waiting', () => {
    const out = waitingOf([msg({ from: 'newsletter@council.org', received: threeDaysAgo })]);
    assert(out.length === 0, 'put a newsletter on the nag list');
  });

  check('her own mail to herself is not waiting', () => {
    assert(waitingOf([msg({ from: HER, received: threeDaysAgo })]).length === 0, 'listed her own mail');
  });

  check('drafts are not waiting', () => {
    assert(
      waitingOf([msg({ received: threeDaysAgo, isDraft: true })]).length === 0,
      'listed a draft',
    );
  });

  check('the list is oldest first', () => {
    const out = waitingOf([
      msg({ id: 'newer', conversationId: 'c1', from: 'a@x.org', received: threeDaysAgo }),
      msg({
        id: 'older',
        conversationId: 'c2',
        from: 'b@x.org',
        received: new Date(now.getTime() - 9 * DAY).toISOString(),
      }),
    ]);
    assert(out[0]?.message.id === 'older', 'did not put the longest-waiting item first');
  });

  // -------------------------------------------------------------------------
  // Flags
  // -------------------------------------------------------------------------

  check('a flagged message is reported regardless of To line or sender', () => {
    const flagged = msg({ id: 'f', from: 'newsletter@x.org', to: ['other@x.org'], flagged: true });
    assert(findOpenFlags([flagged], []).length === 1, 'dropped her own explicit flag');
  });

  check('an item already listed as waiting is not repeated as a flag', () => {
    const both = msg({ id: 'dup', conversationId: 'thread-4', received: threeDaysAgo, flagged: true });
    const waiting = waitingOf([both]);
    assert(waiting.length === 1, 'setup: expected it to be waiting');
    assert(findOpenFlags([both], waiting).length === 0, 'reported the same item twice');
  });

  // -------------------------------------------------------------------------
  // Unread
  // -------------------------------------------------------------------------

  check('unread is grouped by sender and split from bulk', () => {
    const report = buildUnread(
      [
        msg({ id: 'u1', from: 'tessa@partner.org', isRead: false }),
        msg({ id: 'u2', from: 'tessa@partner.org', isRead: false }),
        msg({ id: 'u3', from: 'newsletter@council.org', isRead: false }),
        msg({ id: 'r1', from: 'tessa@partner.org', isRead: true }),
      ],
      [HER],
      now,
      DEFAULT_IGNORED_SENDERS,
    );
    assert(report.total === 3, `total ${report.total}`);
    assert(report.people.length === 1 && report.people[0]?.count === 2, 'mis-grouped people');
    assert(report.automated.length === 1, 'did not separate bulk mail');
  });

  check('unread older than a week is counted separately', () => {
    const report = buildUnread(
      [msg({ id: 'old', isRead: false, received: new Date(now.getTime() - 10 * DAY).toISOString() })],
      [HER],
      now,
      DEFAULT_IGNORED_SENDERS,
    );
    assert(report.staleCount === 1, `staleCount ${report.staleCount}`);
  });

  // -------------------------------------------------------------------------
  // Calendar
  // -------------------------------------------------------------------------

  const window = resolveWindow(now, TZ, undefined);

  check('an unanswered invitation is reported, an accepted one is not', () => {
    const gaps = buildCalendarGaps(
      [
        evt({ startLocal: '2026-08-05T09:00:00.0000000', responseStatus: { response: 'notResponded' } }),
        evt({ startLocal: '2026-08-05T14:00:00.0000000', responseStatus: { response: 'accepted' } }),
      ],
      TZ,
      window,
    );
    assert(gaps.unanswered.length === 1, `unanswered ${gaps.unanswered.length}`);
  });

  check('her own appointments are not unanswered invitations', () => {
    // An event she created for herself has no invitation semantics.
    const gaps = buildCalendarGaps(
      [
        evt({
          startLocal: '2026-08-05T09:00:00.0000000',
          isOrganizer: true,
          responseStatus: { response: 'notResponded' },
        }),
      ],
      TZ,
      window,
    );
    assert(gaps.unanswered.length === 0, 'asked her to RSVP to her own appointment');
  });

  check('next week is separated and its hours are totalled', () => {
    const gaps = buildCalendarGaps(
      [
        evt({
          startLocal: '2026-08-11T09:00:00.0000000',
          endLocal: '2026-08-11T11:00:00.0000000',
          responseStatus: { response: 'notResponded' },
        }),
      ],
      TZ,
      window,
    );
    assert(gaps.nextWeek.length === 1, 'lost a next-week event');
    assert(gaps.nextWeekUnanswered.length === 1, 'missed an actionable RSVP');
    assert(gaps.nextWeekHours === 2, `hours ${gaps.nextWeekHours}`);
  });

  check('declined meetings do not consume next week\'s hours', () => {
    const gaps = buildCalendarGaps(
      [
        evt({
          startLocal: '2026-08-11T09:00:00.0000000',
          endLocal: '2026-08-11T11:00:00.0000000',
          responseStatus: { response: 'declined' },
        }),
      ],
      TZ,
      window,
    );
    assert(gaps.nextWeekHours === 0, `counted declined hours: ${gaps.nextWeekHours}`);
  });

  check('cancelled events are ignored entirely', () => {
    const gaps = buildCalendarGaps(
      [evt({ startLocal: '2026-08-05T09:00:00.0000000', isCancelled: true, responseStatus: { response: 'notResponded' } })],
      TZ,
      window,
    );
    assert(gaps.unanswered.length === 0, 'reported a cancelled meeting');
  });

  // -------------------------------------------------------------------------
  // Activity review
  // -------------------------------------------------------------------------

  check('meetings and hours are counted, all-day entries excluded', () => {
    const review = buildReview(
      [
        evt({ startLocal: '2026-08-04T09:00:00.0000000', endLocal: '2026-08-04T10:30:00.0000000' }),
        // A holiday or out-of-office marker. Counting its hours would distort
        // the total wildly.
        evt({
          startLocal: '2026-08-05T00:00:00.0000000',
          endLocal: '2026-08-06T00:00:00.0000000',
          isAllDay: true,
        }),
      ],
      [],
      TZ,
      window,
    );
    assert(review.meetingsHeld === 1, `meetings ${review.meetingsHeld}`);
    assert(review.meetingHours === 1.5, `hours ${review.meetingHours}`);
  });

  check('declined meetings did not happen for her', () => {
    const review = buildReview(
      [evt({ startLocal: '2026-08-04T09:00:00.0000000', responseStatus: { response: 'declined' } })],
      [],
      TZ,
      window,
    );
    assert(review.meetingsHeld === 0, 'counted a meeting she declined');
  });

  check('sent mail is counted with distinct recipients and threads', () => {
    const inWeek = '2026-08-05T16:00:00Z';
    const review = buildReview(
      [],
      [
        msg({ id: 's1', from: HER, to: ['a@x.org'], conversationId: 't1', sent: inWeek }),
        msg({ id: 's2', from: HER, to: ['a@x.org'], conversationId: 't1', sent: inWeek }),
        msg({ id: 's3', from: HER, to: ['b@x.org'], conversationId: 't2', sent: inWeek }),
      ],
      TZ,
      window,
    );
    assert(review.emailsSent === 3, `sent ${review.emailsSent}`);
    assert(review.peopleWrittenTo === 2, `people ${review.peopleWrittenTo}`);
    assert(review.threadsAdvanced === 2, `threads ${review.threadsAdvanced}`);
  });

  check('activity outside the window is not counted', () => {
    const review = buildReview(
      [],
      [msg({ id: 'old', from: HER, sent: '2026-07-20T16:00:00Z' })],
      TZ,
      window,
    );
    assert(review.emailsSent === 0, 'counted mail from another week');
  });

  // -------------------------------------------------------------------------
  // Assembly
  // -------------------------------------------------------------------------

  check('a quiet week says so instead of reporting zeros', () => {
    const digest = buildDigest({ inbox: [], sent: [], events: [] }, window, {
      mailbox: HER,
      addresses: [HER],
      timeZone: TZ,
      now,
      graceMs: grace,
      ignoredPatterns: DEFAULT_IGNORED_SENDERS,
    });
    assert(!digest.hasAnythingToSay, 'claimed an empty week had something to say');
    assert(digest.headline === 'nothing waiting - inbox is clear', `headline "${digest.headline}"`);
  });

  check('the headline leads with what is waiting', () => {
    const digest = buildDigest(
      {
        inbox: [
          msg({ id: 'w1', from: 'a@x.org', conversationId: 'c1', received: threeDaysAgo }),
          msg({ id: 'w2', from: 'b@x.org', conversationId: 'c2', received: threeDaysAgo }),
        ],
        sent: [],
        events: [],
      },
      window,
      {
        mailbox: HER,
        addresses: [HER],
        timeZone: TZ,
        now,
        graceMs: grace,
        ignoredPatterns: DEFAULT_IGNORED_SENDERS,
      },
    );
    assert(digest.headline === '2 emails are waiting on you', `headline "${digest.headline}"`);
  });
}

// ---------------------------------------------------------------------------
// Claude Desktop registration
//
// Writing this config means editing a file she may already be using for other
// MCP servers. Clobbering those to install ours would be invisible until
// something else stopped working.
// ---------------------------------------------------------------------------

export async function runDesktopChecks(
  checkAsync: (name: string, fn: () => Promise<void>) => Promise<void>,
  assert: Assert,
): Promise<void> {
  const { mkdtemp, readFile, writeFile } = await import('node:fs/promises');
  const { join } = await import('node:path');
  const { tmpdir } = await import('node:os');
  const { connectToDesktop, disconnectFromDesktop } = await import('../routine/desktop.js');

  const dir = await mkdtemp(join(tmpdir(), 'sorter-desktop-'));
  const configFile = join(dir, 'claude_desktop_config.json');
  const read = async () => JSON.parse(await readFile(configFile, 'utf8'));

  await checkAsync('connect creates a config when none exists', async () => {
    const result = await connectToDesktop(configFile);
    assert(!result.replaced, 'claimed to replace an entry in a new file');
    const config = await read();
    assert(config.mcpServers['outlook-sorter'], 'did not add our server');
    assert(
      Array.isArray(config.mcpServers['outlook-sorter'].args),
      'entry has no args to launch',
    );
  });

  await checkAsync('connect preserves other MCP servers and unrelated keys', async () => {
    await writeFile(
      configFile,
      JSON.stringify({
        mcpServers: { filesystem: { command: 'other' }, github: { command: 'other' } },
        globalShortcut: 'Cmd+Shift+Space',
      }),
      'utf8',
    );

    const result = await connectToDesktop(configFile);
    assert(result.preserved.length === 2, `preserved ${result.preserved.length}, expected 2`);

    const config = await read();
    assert(config.mcpServers.filesystem?.command === 'other', 'destroyed another server');
    assert(config.mcpServers.github?.command === 'other', 'destroyed another server');
    assert(config.globalShortcut === 'Cmd+Shift+Space', 'destroyed an unrelated setting');
    assert(config.mcpServers['outlook-sorter'], 'did not add ours');
  });

  await checkAsync('connect twice replaces rather than duplicates', async () => {
    const result = await connectToDesktop(configFile);
    assert(result.replaced, 'did not notice our entry was already there');
    const config = await read();
    assert(Object.keys(config.mcpServers).length === 3, 'entry count changed unexpectedly');
  });

  await checkAsync('disconnect removes only our entry', async () => {
    const result = await disconnectFromDesktop(configFile);
    assert(result.removed, 'reported nothing removed');
    const config = await read();
    assert(!config.mcpServers['outlook-sorter'], 'left our entry behind');
    assert(config.mcpServers.filesystem, 'removed someone else\'s server');
  });

  await checkAsync('disconnect on a missing config is not an error', async () => {
    const result = await disconnectFromDesktop(join(dir, 'nope.json'));
    assert(!result.removed, 'claimed to remove something from a missing file');
  });

  await checkAsync('a malformed config is replaced rather than fatal', async () => {
    await writeFile(configFile, '{ not json', 'utf8');
    await connectToDesktop(configFile);
    const config = await read();
    assert(config.mcpServers['outlook-sorter'], 'gave up on an unparseable config');
  });
}

// ---------------------------------------------------------------------------
// State encoding
//
// State round-trips through a mail message body, and Graph does not reliably
// store a body as the content type it was given. Corruption here silently resets
// everything the tool has learned.
// ---------------------------------------------------------------------------

export async function runStateChecks(check: Check, assert: Assert): Promise<void> {
  const { extractJson } = await import('../routine/store.js');
  const MARKER = 'INBOX-STEWARD-B64:';
  const encode = (value: unknown) =>
    MARKER + Buffer.from(JSON.stringify(value), 'utf8').toString('base64');

  check('base64 state survives an HTML wrapper', () => {
    const state = { senderRules: [{ pattern: 'a@b.org' }] };
    const wrapped = `<html><head></head><body><div>${encode(state)}</div></body></html>`;
    assert(
      JSON.parse(extractJson(wrapped)).senderRules[0].pattern === 'a@b.org',
      'lost state inside an HTML body',
    );
  });

  check('base64 state survives line breaks injected into the body', () => {
    // An HTML wrapper is free to break lines anywhere, including mid-payload.
    const raw = encode({ ok: true });
    const split = `${raw.slice(0, 20)}\n   ${raw.slice(20)}`;
    assert(JSON.parse(extractJson(split)).ok === true, 'line breaks broke the payload');
  });

  check('a correction containing angle brackets round-trips intact', () => {
    // This is the case that defeated the old unescape approach: a subject like
    // "John <john@x.com>" is indistinguishable from markup once it is in a body.
    const state = { recentCorrections: [{ subject: 'Re: John <john@x.com> asked', sender: 'a@b.org' }] };
    const recovered = JSON.parse(extractJson(encode(state)));
    assert(
      recovered.recentCorrections[0].subject === 'Re: John <john@x.com> asked',
      `mangled to: ${recovered.recentCorrections[0].subject}`,
    );
  });

  check('plain JSON written before base64 is still readable', () => {
    const recovered = JSON.parse(extractJson('{"version":1,"senderRules":[]}'));
    assert(recovered.version === 1, 'dropped pre-upgrade state');
  });

  check('an HTML-escaped plain-JSON body is still recovered', () => {
    const recovered = JSON.parse(extractJson('<div>{&quot;version&quot;:1}</div>'));
    assert(recovered.version === 1, 'could not unescape legacy state');
  });
}

// ---------------------------------------------------------------------------
// Prompt extraction
//
// The prompt documents open with an explanation aimed at a person, including a
// sentence telling the reader to paste the text below. Feeding the whole file to a
// model hands it that commentary as instructions.
// ---------------------------------------------------------------------------

export async function runPromptFileChecks(check: Check, assert: Assert): Promise<void> {
  const { extractPrompt, PROMPT_MARKER } = await import('../routine/prompt-file.js');
  const { readFile } = await import('node:fs/promises');
  const { fileURLToPath } = await import('node:url');
  const { dirname, resolve } = await import('node:path');

  const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

  check('the explanation above the marker is dropped', () => {
    const doc = `# Title\n\nPaste the text below into a scheduled task.\n\n${PROMPT_MARKER}\n\nYou are running a sweep.`;
    const out = extractPrompt(doc);
    assert(!out.includes('Paste the text below'), 'kept the instructions-to-the-human');
    assert(out.startsWith('You are running'), `unexpected start: ${out.slice(0, 40)}`);
  });

  check('a horizontal rule under the marker is not treated as content', () => {
    const out = extractPrompt(`x\n${PROMPT_MARKER}\n\n---\n\nYou are running a sweep.`);
    assert(out.startsWith('You are running'), `unexpected start: ${out.slice(0, 40)}`);
  });

  check('a document with no marker fails loudly', () => {
    let threw = false;
    try {
      extractPrompt('# Just a document\n\nNo marker here.', 'FAKE.md');
    } catch (err) {
      threw = true;
      assert(String(err).includes('FAKE.md'), 'error did not name the file');
    }
    assert(threw, 'silently returned something for a markerless document');
  });

  for (const name of ['PROMPT.md', 'PROMPT-WEEKLY.md']) {
    await (async () => {
      const doc = await readFile(resolve(root, 'routine', name), 'utf8');
      check(`${name} yields a usable prompt`, () => {
        const out = extractPrompt(doc, name);
        assert(out.length > 400, `suspiciously short: ${out.length} chars`);
        assert(
          !out.includes('copy from here down'),
          'the marker itself leaked into the prompt',
        );
        assert(
          !out.toLowerCase().includes('paste it as the prompt'),
          'the human-facing preamble leaked into the prompt',
        );
        assert(out.includes('npm run'), 'the prompt lost its commands');
      });
    })();
  }
}
