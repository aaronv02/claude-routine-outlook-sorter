import type { Window } from './window.js';
import {
  eventDurationHours,
  eventInstant,
  eventResponse,
  roundTo,
  type DigestEvent,
} from './types.js';

/** RSVPs left hanging, and what is coming next. */
export interface CalendarGaps {
  /** Invitations in the reporting week she never responded to. */
  unanswered: DigestEvent[];
  /** Meetings she declined, for the record. */
  declined: DigestEvent[];
  /** Next week's schedule. */
  nextWeek: DigestEvent[];
  /** Next week's invitations still awaiting an RSVP - the actionable ones. */
  nextWeekUnanswered: DigestEvent[];
  /** Hours already committed next week. */
  nextWeekHours: number;
}

/**
 * RSVP debt and the week ahead.
 *
 * Unanswered invitations are the calendar equivalent of unanswered mail: an
 * organizer is waiting on her, and unlike email there is no thread to remind
 * anyone. Next week's are separated out because those are the ones she can still
 * do something about.
 */
export function buildCalendarGaps(
  events: DigestEvent[],
  timeZone: string,
  window: Window,
): CalendarGaps {
  const gaps: CalendarGaps = {
    unanswered: [],
    declined: [],
    nextWeek: [],
    nextWeekUnanswered: [],
    nextWeekHours: 0,
  };

  for (const event of events) {
    if (event.isCancelled) continue;

    const start = eventInstant(event.start, timeZone);
    if (!start) continue;

    const at = start.getTime();
    const inWindow = at >= window.start.getTime() && at < window.end.getTime();
    const inNext = at >= window.nextStart.getTime() && at < window.nextEnd.getTime();
    const response = eventResponse(event);

    if (inWindow) {
      // "none" appears on entries with no invitation semantics, such as
      // appointments she made for herself; only a real unanswered invite counts.
      if (response === 'notResponded' && !event.isOrganizer) gaps.unanswered.push(event);
      if (response === 'declined') gaps.declined.push(event);
      continue;
    }

    if (inNext) {
      gaps.nextWeek.push(event);
      if (!event.isAllDay && response !== 'declined') {
        gaps.nextWeekHours += eventDurationHours(event, timeZone);
      }
      if (response === 'notResponded' && !event.isOrganizer) gaps.nextWeekUnanswered.push(event);
    }
  }

  const byStart = (list: DigestEvent[]) =>
    list.sort(
      (a, b) =>
        (eventInstant(a.start, timeZone)?.getTime() ?? 0) -
        (eventInstant(b.start, timeZone)?.getTime() ?? 0),
    );

  byStart(gaps.unanswered);
  byStart(gaps.declined);
  byStart(gaps.nextWeek);
  byStart(gaps.nextWeekUnanswered);

  gaps.nextWeekHours = roundTo(gaps.nextWeekHours, 1);
  return gaps;
}
