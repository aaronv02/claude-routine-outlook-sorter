/**
 * Timezone arithmetic, without a dependency.
 *
 * The digest reports on a *week*, and a week has boundaries: Monday 00:00 local
 * to the following Monday 00:00 local. Getting the zone wrong shifts the whole
 * reporting window, so this is explicit rather than inferred from whatever
 * machine happens to run the sweep - a cloud runner is almost certainly on UTC
 * while the mailbox owner is not.
 *
 * Node has no built-in way to turn "Monday 00:00 in America/Denver" into an
 * instant. `Intl.DateTimeFormat` can go the other way - instant to wall clock in
 * a named zone - so that is what everything here is built out of.
 */

/** Wall-clock fields of an instant, as seen in a named zone. */
export interface Wall {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
  /** 1 = Monday ... 7 = Sunday, matching ISO rather than JS. */
  isoWeekday: number;
}

const WEEKDAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/**
 * Throws on an unknown zone rather than silently falling back to UTC, which
 * would shift every boundary in the report by hours with nothing to indicate it.
 */
export function assertValidZone(timeZone: string): void {
  try {
    new Intl.DateTimeFormat('en-US', { timeZone });
  } catch {
    throw new Error(
      `"${timeZone}" is not a timezone this system recognizes. Use an IANA name such as America/Denver.`,
    );
  }
}

const partsFormatter = new Map<string, Intl.DateTimeFormat>();

function formatter(timeZone: string): Intl.DateTimeFormat {
  let existing = partsFormatter.get(timeZone);
  if (!existing) {
    existing = new Intl.DateTimeFormat('en-US', {
      timeZone,
      hourCycle: 'h23',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      weekday: 'short',
    });
    partsFormatter.set(timeZone, existing);
  }
  return existing;
}

export function wallOf(instant: Date, timeZone: string): Wall {
  const parts = new Map(
    formatter(timeZone)
      .formatToParts(instant)
      .map((p) => [p.type, p.value] as const),
  );
  const weekdayIndex = WEEKDAYS.indexOf(parts.get('weekday') ?? 'Mon');

  return {
    year: Number(parts.get('year')),
    month: Number(parts.get('month')),
    day: Number(parts.get('day')),
    hour: Number(parts.get('hour')),
    minute: Number(parts.get('minute')),
    second: Number(parts.get('second')),
    isoWeekday: (weekdayIndex === -1 ? 0 : weekdayIndex) + 1,
  };
}

/** The zone's offset from UTC at a given instant, in milliseconds. */
function offsetAt(instant: Date, timeZone: string): number {
  const w = wallOf(instant, timeZone);
  const asUtc = Date.UTC(w.year, w.month - 1, w.day, w.hour, w.minute, w.second);
  // Discard sub-second precision, which formatToParts does not report.
  return asUtc - Math.floor(instant.getTime() / 1000) * 1000;
}

/**
 * The instant at which the given wall-clock time occurs in a zone.
 *
 * Solved by iteration because the offset depends on the answer: guess that the
 * wall time is UTC, measure the zone's offset near that guess, correct, then
 * confirm. The second pass is what makes the days either side of a DST
 * transition come out right - a single pass is off by an hour there, which is
 * enough to move a Monday-morning boundary into Sunday.
 *
 * In the spring-forward gap the requested wall time does not exist; this settles
 * on the instant just after the jump, which is the sane reading of "midnight" on
 * a day that has no midnight.
 */
export function instantOf(
  wall: { year: number; month: number; day: number; hour?: number; minute?: number },
  timeZone: string,
): Date {
  const target = Date.UTC(wall.year, wall.month - 1, wall.day, wall.hour ?? 0, wall.minute ?? 0);

  let guess = new Date(target);
  for (let pass = 0; pass < 2; pass++) {
    const corrected = new Date(target - offsetAt(guess, timeZone));
    if (corrected.getTime() === guess.getTime()) break;
    guess = corrected;
  }
  return guess;
}

/** Midnight local on the Monday of the ISO week containing `instant`. */
export function isoWeekStart(instant: Date, timeZone: string): Date {
  const w = wallOf(instant, timeZone);
  const monday = new Date(Date.UTC(w.year, w.month - 1, w.day) - (w.isoWeekday - 1) * 86_400_000);
  return instantOf(
    {
      year: monday.getUTCFullYear(),
      month: monday.getUTCMonth() + 1,
      day: monday.getUTCDate(),
    },
    timeZone,
  );
}

export function addDays(instant: Date, days: number, timeZone: string): Date {
  const w = wallOf(instant, timeZone);
  const shifted = new Date(Date.UTC(w.year, w.month - 1, w.day + days));
  return instantOf(
    {
      year: shifted.getUTCFullYear(),
      month: shifted.getUTCMonth() + 1,
      day: shifted.getUTCDate(),
      hour: w.hour,
      minute: w.minute,
    },
    timeZone,
  );
}

/**
 * ISO week label, e.g. "2026-W33".
 *
 * Used as the idempotency key for "did we already report this week?", so it has
 * to follow the ISO rule that a week belongs to the year containing its Thursday
 * - otherwise the last days of December and first of January disagree about
 * which year they are in, and one week gets reported twice or not at all.
 */
export function isoLabel(weekStart: Date, timeZone: string): string {
  const w = wallOf(weekStart, timeZone);
  const thursday = new Date(Date.UTC(w.year, w.month - 1, w.day + 3));

  const year = thursday.getUTCFullYear();
  const jan1 = Date.UTC(year, 0, 1);
  const week = Math.floor((thursday.getTime() - jan1) / 86_400_000 / 7) + 1;

  return `${year}-W${String(week).padStart(2, '0')}`;
}

/** "Mon 27 – Fri 31 Jul", collapsing the month when both dates share one. */
export function describeRange(start: Date, endExclusive: Date, timeZone: string): string {
  const lastInstant = new Date(Math.max(start.getTime(), endExclusive.getTime() - 1000));
  const a = wallOf(start, timeZone);
  const b = wallOf(lastInstant, timeZone);

  const left = `${WEEKDAYS[a.isoWeekday - 1]} ${a.day}`;
  const right = `${WEEKDAYS[b.isoWeekday - 1]} ${b.day} ${MONTHS[b.month - 1]}`;

  return a.month === b.month
    ? `${left} – ${right}`
    : `${left} ${MONTHS[a.month - 1]} – ${right}`;
}

/** "Tue 12 Aug, 14:30", for listing individual items. */
export function describeMoment(instant: Date, timeZone: string): string {
  const w = wallOf(instant, timeZone);
  return `${WEEKDAYS[w.isoWeekday - 1]} ${w.day} ${MONTHS[w.month - 1]}, ${String(w.hour).padStart(2, '0')}:${String(w.minute).padStart(2, '0')}`;
}
