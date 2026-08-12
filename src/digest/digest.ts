import { buildCalendarGaps, type CalendarGaps } from './calendar.js';
import { buildReview, type Review } from './review.js';
import { buildUnread, type UnreadReport } from './unread.js';
import { findOpenFlags, findWaiting, type WaitingItem } from './waiting.js';
import type { Window } from './window.js';
import type { DigestEvent, DigestMessage } from './types.js';

/**
 * Assembling the weekly digest.
 *
 * Every function reachable from here is a pure function of its inputs: no
 * network, no clock beyond the `now` that gets passed in. That is what makes the
 * whole report testable on a machine that has never seen the target mailbox, and
 * it is why the Graph fetching lives in a separate module.
 */

export interface DigestInput {
  inbox: DigestMessage[];
  sent: DigestMessage[];
  events: DigestEvent[];
}

export interface DigestOptions {
  mailbox: string;
  /**
   * Every address she may be reached at, including `mailbox`. Role aliases and
   * shared mailboxes belong here, or mail sent to them never registers as having
   * been addressed to her.
   */
  addresses: string[];
  timeZone: string;
  now: Date;
  graceMs: number;
  ignoredPatterns: string[];
}

export interface Digest {
  window: Window;
  mailbox: string;
  generatedAt: string;
  timeZone: string;

  waiting: WaitingItem[];
  unread: UnreadReport;
  review: Review;
  calendar: CalendarGaps;
  /** Messages carrying an unresolved follow-up flag, whenever they arrived. */
  flagged: DigestMessage[];

  /** True when there is anything actionable at all. */
  hasAnythingToSay: boolean;
  /** One-sentence summary, suitable as a subject line. */
  headline: string;
}

export function buildDigest(
  input: DigestInput,
  window: Window,
  options: DigestOptions,
): Digest {
  const addresses = options.addresses.length > 0 ? options.addresses : [options.mailbox];

  const waiting = findWaiting(
    input.inbox,
    input.sent,
    addresses,
    options.graceMs,
    options.now,
    options.ignoredPatterns,
  );

  const unread = buildUnread(input.inbox, addresses, options.now, options.ignoredPatterns);
  const calendar = buildCalendarGaps(input.events, options.timeZone, window);
  const flagged = findOpenFlags(input.inbox, waiting);

  const digest: Digest = {
    window,
    mailbox: options.mailbox,
    generatedAt: options.now.toISOString(),
    timeZone: options.timeZone,
    waiting,
    unread,
    review: buildReview(input.events, input.sent, options.timeZone, window),
    calendar,
    flagged,
    hasAnythingToSay:
      waiting.length > 0 ||
      flagged.length > 0 ||
      unread.total > 0 ||
      calendar.nextWeekUnanswered.length > 0,
    headline: headlineFor(waiting, unread),
  };

  return digest;
}

/**
 * The one-line summary.
 *
 * Ordered by what actually demands attention. On a genuinely quiet week it says
 * so plainly rather than announcing "0 items waiting" as though that were a
 * finding.
 */
function headlineFor(waiting: WaitingItem[], unread: UnreadReport): string {
  if (waiting.length === 1) return '1 email is waiting on you';
  if (waiting.length > 1) return `${waiting.length} emails are waiting on you`;
  if (unread.staleCount > 0) return `${unread.staleCount} unread from last week or earlier`;
  if (unread.total > 0) return `${unread.total} unread`;
  return 'nothing waiting - inbox is clear';
}
