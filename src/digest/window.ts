import { addDays, describeRange, isoLabel, isoWeekStart, wallOf } from './zone.js';

/** The stretch of time a digest reports on. */
export interface Window {
  /** ISO label such as "2026-W33". The idempotency key: a catch-up run must not resend a week. */
  label: string;
  /** Inclusive start, exclusive end. */
  start: Date;
  end: Date;
  /** The following week, for the look-ahead section. */
  nextStart: Date;
  nextEnd: Date;
  /** True when this run is covering a week whose scheduled slot was missed. */
  catchUp: boolean;
  /** "Mon 27 – Fri 31 Jul", for the report. */
  description: string;
}

/**
 * Decide which week to report on.
 *
 * This is the fix for the central weakness of anything scheduled: the run can be
 * late. A cloud runner might be throttled, a laptop might be asleep, a token
 * might have needed re-auth. If the slot is Friday afternoon and the sweep
 * actually fires on Monday, naively reporting "this week" produces an almost
 * empty digest for a week that has barely started - and the week that actually
 * mattered is never reported at all.
 *
 * So: Friday through Sunday reports the current week. Monday through Thursday
 * reports the *previous* week, but only if that week was never sent - otherwise
 * it reports the current week as usual.
 */
export function resolveWindow(
  now: Date,
  timeZone: string,
  lastSentWeek: string | undefined,
): Window {
  const currentStart = isoWeekStart(now, timeZone);
  const weekday = wallOf(now, timeZone).isoWeekday;

  let target = currentStart;
  let catchUp = false;

  if (weekday >= 1 && weekday <= 4) {
    const previousStart = addDays(currentStart, -7, timeZone);
    if (isoLabel(previousStart, timeZone) !== lastSentWeek) {
      target = previousStart;
      catchUp = true;
    }
  }

  let end = addDays(target, 7, timeZone);
  // Never claim to cover time that hasn't happened yet.
  if (now.getTime() < end.getTime()) end = now;

  const nextStart = addDays(target, 7, timeZone);

  return {
    label: isoLabel(target, timeZone),
    start: target,
    end,
    nextStart,
    nextEnd: addDays(nextStart, 7, timeZone),
    catchUp,
    description: describeRange(target, end, timeZone),
  };
}
