import type { Window } from './window.js';
import { wallOf } from './zone.js';
import {
  eventDurationHours,
  eventInstant,
  eventResponse,
  normalizeAddress,
  roundTo,
  sentAt,
  type DigestEvent,
  type DigestMessage,
} from './types.js';

/** What happened in the mailbox and calendar this week. */
export interface Review {
  meetingsHeld: number;
  meetingHours: number;
  meetingsOrganized: number;
  emailsSent: number;
  peopleWrittenTo: number;
  /** Conversations she sent the last message in: threads she moved along. */
  threadsAdvanced: number;
  busiestDay: string;
  /** A few of the week's larger commitments, for context. */
  highlights: DigestEvent[];
}

const DAY_NAMES = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
const MAX_HIGHLIGHTS = 4;

/**
 * Summarize the week's activity.
 *
 * A deliberate note on what this is not: Outlook cannot see accomplishment. It
 * records that a meeting existed and that mail was sent, not what was decided or
 * achieved, and it cannot confirm attendance - only the RSVP. Site visits, phone
 * calls, grant deliberations, and anything done in another system are invisible
 * here. The report is framed as *activity* for that reason, and the prompt that
 * renders it is told to say so rather than implying more than the data supports.
 */
export function buildReview(
  events: DigestEvent[],
  sent: DigestMessage[],
  timeZone: string,
  window: Window,
): Review {
  const review: Review = {
    meetingsHeld: 0,
    meetingHours: 0,
    meetingsOrganized: 0,
    emailsSent: 0,
    peopleWrittenTo: 0,
    threadsAdvanced: 0,
    busiestDay: '',
    highlights: [],
  };

  // Indexed 0 = Monday, matching ISO.
  const perDay = new Array<number>(7).fill(0);
  let anyActivity = false;

  for (const event of events) {
    if (event.isCancelled) continue;
    // All-day entries are usually holidays, travel, or out-of-office markers
    // rather than meetings, and counting their hours would wildly distort the
    // total.
    if (event.isAllDay) continue;
    // Events she declined did not happen for her.
    if (eventResponse(event) === 'declined') continue;

    const start = eventInstant(event.start, timeZone);
    if (!start) continue;
    const at = start.getTime();
    if (at < window.start.getTime() || at >= window.end.getTime()) continue;

    review.meetingsHeld++;
    review.meetingHours += eventDurationHours(event, timeZone);
    if (event.isOrganizer) review.meetingsOrganized++;

    perDay[wallOf(start, timeZone).isoWeekday - 1]!++;
    anyActivity = true;
  }

  const correspondents = new Set<string>();
  const threads = new Set<string>();

  for (const message of sent) {
    const when = sentAt(message);
    const at = when.getTime();
    if (Number.isNaN(at)) continue;
    if (at < window.start.getTime() || at >= window.end.getTime()) continue;

    review.emailsSent++;
    for (const recipient of message.toRecipients ?? []) {
      const address = normalizeAddress(recipient.emailAddress?.address);
      if (address) correspondents.add(address);
    }
    if (message.conversationId) threads.add(message.conversationId);

    perDay[wallOf(when, timeZone).isoWeekday - 1]!++;
    anyActivity = true;
  }

  review.peopleWrittenTo = correspondents.size;
  review.threadsAdvanced = threads.size;
  review.meetingHours = roundTo(review.meetingHours, 1);

  if (anyActivity) {
    // Scanned in weekday order so a tie resolves to the earlier day rather than
    // arbitrarily.
    let busiest = 0;
    for (let day = 1; day < 7; day++) {
      if ((perDay[day] as number) > (perDay[busiest] as number)) busiest = day;
    }
    review.busiestDay = DAY_NAMES[busiest] as string;
  }

  review.highlights = pickHighlights(events, timeZone, window);
  return review;
}

/** How much of the week a meeting represents: duration plus a nod to how many people were in it. */
function weight(event: DigestEvent, timeZone: string): number {
  const humans = (event.attendees ?? []).filter((a) => a.type !== 'resource').length;
  return eventDurationHours(event, timeZone) + humans / 2;
}

/**
 * The week's most substantial meetings.
 *
 * Weighted by attendees and duration: a two-hour board meeting with nine people is
 * more of the week than a fifteen-minute one-to-one.
 */
function pickHighlights(
  events: DigestEvent[],
  timeZone: string,
  window: Window,
): DigestEvent[] {
  return events
    .filter((e) => !e.isCancelled && !e.isAllDay && eventResponse(e) !== 'declined')
    .filter((e) => {
      const start = eventInstant(e.start, timeZone);
      if (!start) return false;
      const at = start.getTime();
      return at >= window.start.getTime() && at < window.end.getTime();
    })
    .sort((a, b) => weight(b, timeZone) - weight(a, timeZone))
    .slice(0, MAX_HIGHLIGHTS);
}
