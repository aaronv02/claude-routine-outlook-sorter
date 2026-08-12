/**
 * A fake week, run through the real digest. `npm run scenario`
 *
 * The sorting side can be scored against labelled fixtures. The digest cannot:
 * there is no "correct" weekly summary to diff against. What can be checked is
 * whether it reaches the right conclusions about a week whose facts are known -
 * so this builds one deliberately, containing every trap the analysis is supposed
 * to avoid, and prints what the routine would report.
 *
 * The mailbox below contains, on purpose:
 *
 *   - mail she was CC'd on but not asked                  (must not be "waiting")
 *   - a thread she already answered                        (must not be "waiting")
 *   - a five-message thread she never answered             (must count once, aged
 *                                                          from the first ask)
 *   - a newsletter that went unanswered                    (unread, never waiting)
 *   - a message from an address that merely looks like bulk (must still be waiting)
 *   - a flagged newsletter                                 (flags outrank sender)
 *   - mail to a role alias                                 (waiting, via the alias)
 *   - a meeting she declined                                (not in her activity)
 *   - an all-day out-of-office marker                       (not counted as hours)
 *   - an unanswered invitation for next week                (actionable RSVP)
 *
 * Run it after changing anything in src/digest/ and read the output. Numbers that
 * look plausible are the failure mode here, so the assertions at the end state what
 * the week actually contains.
 */

import { resolveWindow } from '../src/digest/window.js';
import { buildDigest } from '../src/digest/digest.js';
import { DEFAULT_IGNORED_SENDERS } from '../src/digest/senders.js';
import { describeMoment } from '../src/digest/zone.js';
import { eventInstant, type DigestEvent, type DigestMessage } from '../src/digest/types.js';

const TZ = 'America/Denver';
const HER = 'director@swcommunityfoundation.org';
const ALIAS = 'grants@swcommunityfoundation.org';

/** Friday 7 August 2026, 16:00 in Denver - the moment the digest would fire. */
const NOW = new Date('2026-08-07T22:00:00Z');

const bold = (s: string) => `\x1b[1m${s}\x1b[0m`;
const dim = (s: string) => `\x1b[2m${s}\x1b[0m`;
const green = (s: string) => `\x1b[32m${s}\x1b[0m`;
const red = (s: string) => `\x1b[31m${s}\x1b[0m`;

let id = 0;

function mail(o: {
  from: string;
  name: string;
  subject: string;
  daysAgo: number;
  to?: string[];
  cc?: string[];
  thread?: string;
  read?: boolean;
  flagged?: boolean;
}): DigestMessage {
  return {
    id: `m${++id}`,
    conversationId: o.thread ?? `t${id}`,
    subject: o.subject,
    bodyPreview: `Regarding ${o.subject.toLowerCase()}.`,
    receivedDateTime: new Date(NOW.getTime() - o.daysAgo * 86_400_000).toISOString(),
    isRead: o.read ?? true,
    isDraft: false,
    from: { emailAddress: { address: o.from, name: o.name } },
    toRecipients: (o.to ?? [HER]).map((a) => ({ emailAddress: { address: a } })),
    ccRecipients: (o.cc ?? []).map((a) => ({ emailAddress: { address: a } })),
    flag: { flagStatus: o.flagged ? 'flagged' : 'notFlagged' },
  };
}

function sent(o: { to: string[]; subject: string; daysAgo: number; thread?: string }): DigestMessage {
  const at = new Date(NOW.getTime() - o.daysAgo * 86_400_000).toISOString();
  return {
    id: `s${++id}`,
    conversationId: o.thread ?? `t${id}`,
    subject: o.subject,
    receivedDateTime: at,
    sentDateTime: at,
    isRead: true,
    from: { emailAddress: { address: HER, name: 'The Director' } },
    toRecipients: o.to.map((a) => ({ emailAddress: { address: a } })),
    flag: { flagStatus: 'notFlagged' },
  };
}

function meeting(o: {
  subject: string;
  day: string;
  from: string;
  to: string;
  attendees?: number;
  response?: string;
  organizer?: boolean;
  allDay?: boolean;
  cancelled?: boolean;
}): DigestEvent {
  return {
    id: `e${++id}`,
    subject: o.subject,
    start: { dateTime: `${o.day}T${o.from}:00.0000000`, timeZone: TZ },
    end: { dateTime: `${o.day}T${o.to}:00.0000000`, timeZone: TZ },
    isAllDay: o.allDay ?? false,
    isCancelled: o.cancelled ?? false,
    isOrganizer: o.organizer ?? false,
    responseStatus: { response: o.response ?? 'accepted' },
    organizer: { emailAddress: { address: 'someone@example.org', name: 'Organizer' } },
    attendees: Array.from({ length: o.attendees ?? 2 }, (_, i) => ({
      type: 'required',
      emailAddress: { address: `a${i}@example.org` },
    })),
  };
}

// ---------------------------------------------------------------------------
// The week
// ---------------------------------------------------------------------------

const inbox: DigestMessage[] = [
  // Genuinely waiting.
  mail({ from: 'ed@mancosvalleyresources.org', name: 'Dana Ruiz', subject: 'Following up on our letter of inquiry', daysAgo: 9 }),
  mail({ from: 'dcastellano@fourcornerslaw.com', name: 'Diane Castellano', subject: 'Client bequest - can we talk Thursday?', daysAgo: 4 }),
  // Sent to the grants@ alias, not her personal address. Waiting via the alias.
  mail({ from: 'grants@sanjuanmountainsassoc.org', name: 'Will Trent', subject: 'LAUNCH budget template question', daysAgo: 3, to: [ALIAS] }),
  // Looks like bulk but is a person. Must NOT be filtered out.
  mail({ from: 'goodnews@pagosayouthcenter.org', name: 'Marisol Vega', subject: 'Merger conversation - your advice?', daysAgo: 6 }),

  // A thread she never answered, five messages deep. One entry, aged from the first.
  mail({ from: 'cwexler@wexlercpa.com', name: 'Carl Wexler', subject: '990 - grant schedule', daysAgo: 12, thread: 'audit' }),
  mail({ from: 'cwexler@wexlercpa.com', name: 'Carl Wexler', subject: 'Re: 990 - grant schedule', daysAgo: 10, thread: 'audit' }),
  mail({ from: 'cwexler@wexlercpa.com', name: 'Carl Wexler', subject: 'Re: 990 - grant schedule', daysAgo: 8, thread: 'audit' }),
  mail({ from: 'cwexler@wexlercpa.com', name: 'Carl Wexler', subject: 'Re: 990 - still need this', daysAgo: 5, thread: 'audit' }),
  mail({ from: 'cwexler@wexlercpa.com', name: 'Carl Wexler', subject: 'Re: 990 - checking in again', daysAgo: 2, thread: 'audit' }),

  // CC'd, not asked. Must not appear.
  mail({ from: 'boardchair@swcommunityfoundation.org', name: 'Board Chair', subject: 'FYI - retreat logistics', daysAgo: 5, to: ['staff@swcommunityfoundation.org'], cc: [HER] }),

  // She replied to this one. Must not appear.
  mail({ from: 'jromero@durangoherald.com', name: 'Jess Romero', subject: 'Interview request', daysAgo: 6, thread: 'press' }),

  // Too recent to count as a failure.
  mail({ from: 'elena.marquez2027@student.durangoschools.org', name: 'Elena Marquez', subject: 'Transcript deadline', daysAgo: 0 }),

  // Bulk: unread, never waiting.
  mail({ from: 'news@cof.org', name: 'Council on Foundations', subject: 'This week in philanthropy', daysAgo: 3, read: false }),
  mail({ from: 'updates@coloradononprofits.org', name: 'CNA', subject: 'Legislative update', daysAgo: 4, read: false }),
  mail({ from: 'noreply@philanthropy.com', name: 'Chronicle', subject: 'Philanthropy 400', daysAgo: 11, read: false }),

  // Unread from a person, twice.
  mail({ from: 'tbrennan@bayfieldschools.net', name: 'Tom Brennan', subject: 'Recommendation letter', daysAgo: 8, read: false, to: [HER], thread: 'rec1' }),
  mail({ from: 'tbrennan@bayfieldschools.net', name: 'Tom Brennan', subject: 'Re: Recommendation letter', daysAgo: 7, read: false, to: [HER], thread: 'rec1' }),

  // Flagged newsletter: her own flag outranks the bulk sender.
  mail({ from: 'webinars@candid.org', name: 'Candid', subject: 'Impact story webinar', daysAgo: 4, flagged: true }),
];

const sentItems: DigestMessage[] = [
  sent({ to: ['jromero@durangoherald.com'], subject: 'Re: Interview request', daysAgo: 5, thread: 'press' }),
  sent({ to: ['boardchair@swcommunityfoundation.org'], subject: 'Retreat agenda', daysAgo: 3 }),
  sent({ to: ['kbaptiste@pinerivervalley.org', 'ed@archuletahousingcoalition.org'], subject: 'Grant decisions', daysAgo: 2 }),
  sent({ to: ['statements@vanguardcharitable.org'], subject: 'Q2 question', daysAgo: 1 }),
];

const events: DigestEvent[] = [
  // This week.
  meeting({ subject: 'Board meeting', day: '2026-08-04', from: '09:00', to: '11:30', attendees: 9 }),
  meeting({ subject: 'Scholarship committee', day: '2026-08-05', from: '13:00', to: '14:00', attendees: 4 }),
  meeting({ subject: 'Donor coffee - Holloway', day: '2026-08-06', from: '10:00', to: '11:00', attendees: 2, organizer: true }),
  meeting({ subject: 'Declined: vendor pitch', day: '2026-08-05', from: '15:00', to: '16:00', response: 'declined' }),
  meeting({ subject: 'Out of office', day: '2026-08-03', from: '00:00', to: '23:59', allDay: true }),
  meeting({ subject: 'Cancelled: site visit', day: '2026-08-06', from: '14:00', to: '15:00', cancelled: true }),
  meeting({ subject: 'Regional funders call', day: '2026-08-05', from: '16:00', to: '17:00', response: 'notResponded' }),

  // Next week.
  meeting({ subject: 'Wine Experience walkthrough', day: '2026-08-11', from: '09:00', to: '11:00', attendees: 5 }),
  meeting({ subject: 'Audit fieldwork kickoff', day: '2026-08-12', from: '13:00', to: '15:00', response: 'notResponded' }),
  meeting({ subject: 'Staff one-to-one', day: '2026-08-13', from: '10:00', to: '10:30', organizer: true }),
];

// ---------------------------------------------------------------------------

const window = resolveWindow(NOW, TZ, '2026-W31');
const digest = buildDigest({ inbox, sent: sentItems, events }, window, {
  mailbox: HER,
  addresses: [HER, ALIAS],
  timeZone: TZ,
  now: NOW,
  graceMs: 18 * 3_600_000,
  ignoredPatterns: DEFAULT_IGNORED_SENDERS,
});

console.log(`\n${bold('What the routine would report')}`);
console.log(dim(`week ${window.label}, covering ${window.description}\n`));
console.log(`${bold('Headline:')} ${digest.headline}\n`);

console.log(bold(`WAITING ON A REPLY (${digest.waiting.length})`));
for (const w of digest.waiting) {
  console.log(
    `  ${w.message.from?.emailAddress?.name} <${w.message.from?.emailAddress?.address}>`,
  );
  console.log(dim(`    "${w.message.subject}" — ${w.ageDays} days`));
}

console.log(`\n${bold(`SHE FLAGGED (${digest.flagged.length})`)}`);
for (const m of digest.flagged) console.log(`  ${m.from?.emailAddress?.name} — "${m.subject}"`);

console.log(`\n${bold('UNREAD')}`);
console.log(`  ${digest.unread.total} total, ${digest.unread.staleCount} older than a week`);
for (const g of digest.unread.people) console.log(`  person: ${g.name} (${g.count})`);
console.log(`  bulk: ${digest.unread.automated.reduce((n, g) => n + g.count, 0)} from ${digest.unread.automated.length} senders`);

const r = digest.review;
console.log(`\n${bold('HER WEEK')}`);
console.log(`  ${r.meetingsHeld} meetings, ${r.meetingHours}h, ${r.meetingsOrganized} she organized`);
console.log(`  ${r.emailsSent} emails to ${r.peopleWrittenTo} people, ${r.threadsAdvanced} threads`);
console.log(`  busiest: ${r.busiestDay}`);

console.log(`\n${bold('CALENDAR')}`);
console.log(`  never answered this week: ${digest.calendar.unanswered.length}`);
console.log(`  next week: ${digest.calendar.nextWeek.length} entries, ${digest.calendar.nextWeekHours}h`);
for (const e of digest.calendar.nextWeekUnanswered) {
  const at = eventInstant(e.start, TZ);
  console.log(`  needs RSVP: "${e.subject}" ${at ? describeMoment(at, TZ) : ''}`);
}

// ---------------------------------------------------------------------------
// What the week actually contains
// ---------------------------------------------------------------------------

const senders = digest.waiting.map((w) => w.message.from?.emailAddress?.address ?? '');
const checks: [string, boolean][] = [
  // Six, not nine: the five-message 990 thread and the two-message
  // recommendation thread each collapse to one item.
  ['6 items waiting, threads collapsed', digest.waiting.length === 6],
  ['the 990 thread is aged from the first ask (12d)', digest.waiting[0]?.ageDays === 12],
  ['oldest first', digest.waiting[0]?.message.conversationId === 'audit'],
  ['mail she was only CC\'d on is absent', !senders.includes('boardchair@swcommunityfoundation.org')],
  ['the thread she answered is absent', !senders.includes('jromero@durangoherald.com')],
  ['today\'s mail is not a failure', !senders.includes('elena.marquez2027@student.durangoschools.org')],
  ['newsletters never appear as waiting', !senders.some((s) => s.startsWith('news@') || s.startsWith('updates@'))],
  ['mail to the grants@ alias counts', senders.includes('grants@sanjuanmountainsassoc.org')],
  ['a person at goodnews@ is not mistaken for bulk', senders.includes('goodnews@pagosayouthcenter.org')],
  ['her flag outranks a bulk sender', digest.flagged.some((m) => m.from?.emailAddress?.address === 'webinars@candid.org')],
  ['unread splits 2 people / 3 bulk', digest.unread.people[0]?.count === 2 && digest.unread.automated.length === 3],
  // Four: board, scholarship committee, donor coffee, and the funders call she
  // never RSVP'd to - not responding is not the same as declining.
  ['declined and cancelled meetings excluded (4 held)', r.meetingsHeld === 4],
  ['all-day marker adds no hours (5.5h)', r.meetingHours === 5.5],
  ['next week totals 4.5h', digest.calendar.nextWeekHours === 4.5],
  ['next week RSVP surfaced', digest.calendar.nextWeekUnanswered.length === 1],
  ['this week\'s unanswered invite surfaced', digest.calendar.unanswered.length === 1],
];

console.log(`\n${bold('Does it match the week we built?')}\n`);
let failed = 0;
for (const [label, ok] of checks) {
  console.log(`  ${ok ? green('✓') : red('✗')} ${label}`);
  if (!ok) failed++;
}
console.log('');
if (failed > 0) {
  console.log(red(`${failed} of ${checks.length} wrong.\n`));
  process.exit(1);
}
console.log(green(`All ${checks.length} correct.\n`));
