import { instantOf } from './zone.js';

/**
 * The subset of a mailbox and calendar the digest reasons about, plus the small
 * accessors that carry actual judgement.
 *
 * Kept dependency-free and free of any network so the whole report can be tested
 * on a machine that has never seen the target mailbox.
 */

export interface EmailAddress {
  name?: string;
  address?: string;
}

export interface Recipient {
  emailAddress: EmailAddress;
}

export interface DigestMessage {
  id: string;
  conversationId?: string;
  subject?: string | null;
  bodyPreview?: string | null;
  receivedDateTime: string;
  sentDateTime?: string;
  isRead?: boolean;
  isDraft?: boolean;
  hasAttachments?: boolean;
  importance?: string;
  categories?: string[] | null;
  from?: Recipient | null;
  sender?: Recipient | null;
  toRecipients?: Recipient[] | null;
  ccRecipients?: Recipient[] | null;
  /** "notFlagged" | "complete" | "flagged" */
  flag?: { flagStatus?: string } | null;
  webLink?: string;
}

/** Graph's local-time representation for calendar items: wall clock plus a zone name. */
export interface DateTimeTimeZone {
  dateTime: string;
  timeZone?: string;
}

export interface DigestEvent {
  id: string;
  subject?: string | null;
  bodyPreview?: string | null;
  start: DateTimeTimeZone;
  end: DateTimeTimeZone;
  isAllDay?: boolean;
  isCancelled?: boolean;
  isOrganizer?: boolean;
  /** "none" | "organizer" | "tentativelyAccepted" | "accepted" | "declined" | "notResponded" */
  responseStatus?: { response?: string };
  organizer?: Recipient | null;
  attendees?: { type?: string; emailAddress?: EmailAddress }[] | null;
  location?: { displayName?: string } | null;
  showAs?: string;
  webLink?: string;
}

export function normalizeAddress(value: string | undefined | null): string {
  return (value ?? '').trim().toLowerCase();
}

/** The sender address, lowercased, or "" when absent. */
export function fromAddress(m: DigestMessage): string {
  return normalizeAddress(m.from?.emailAddress?.address ?? m.sender?.emailAddress?.address);
}

/** The sender's display name, falling back to the address. */
export function fromName(m: DigestMessage): string {
  return m.from?.emailAddress?.name?.trim() || fromAddress(m);
}

/**
 * Whether any of her addresses is on the To line specifically.
 *
 * The To/Cc distinction is the crux of the "waiting on you" section: being CC'd
 * is being kept informed, being on To is being asked. Treating them the same
 * produces a nag list nobody trusts.
 *
 * Multiple addresses matter more than it first appears. At a small nonprofit one
 * person often receives mail at several - a personal address, a role alias like
 * grants@, sometimes a shared mailbox. If the digest only knows one of them,
 * everything sent to the others fails the "was she actually asked?" test, and the
 * failure is silent: a suspiciously short list rather than an error.
 */
export function addressedToAny(m: DigestMessage, addresses: string[]): boolean {
  if (addresses.length === 0) return false;
  const wanted = new Set(addresses.map(normalizeAddress).filter(Boolean));

  return (m.toRecipients ?? []).some((r) => {
    const got = normalizeAddress(r.emailAddress?.address);
    return got !== '' && wanted.has(got);
  });
}

/** An active follow-up flag - her own explicit "come back to this". */
export function isFlagged(m: DigestMessage): boolean {
  return m.flag?.flagStatus === 'flagged';
}

export function receivedAt(m: DigestMessage): Date {
  return new Date(m.receivedDateTime);
}

/**
 * When something was sent, falling back to when it was received.
 *
 * Sent-items messages carry both; the fallback covers the odd item where Graph
 * reports only one, which would otherwise sort as the epoch and be treated as a
 * reply that predates every message it could answer.
 */
export function sentAt(m: DigestMessage): Date {
  const raw = m.sentDateTime ?? m.receivedDateTime;
  return new Date(raw);
}

/**
 * Parse one of Graph's zone-less calendar times as wall clock in `timeZone`.
 *
 * Requesting `Prefer: outlook.timezone` makes Graph return these already
 * converted to the mailbox's zone, so interpreting them in that zone is correct.
 * They arrive without an offset - appending "Z" here, which is the tempting
 * shortcut, would silently shift every event by the zone's offset.
 */
export function eventInstant(value: DateTimeTimeZone | undefined, timeZone: string): Date | null {
  const raw = value?.dateTime?.trim();
  if (!raw) return null;

  const match = /^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})(?::(\d{2}))?/.exec(raw);
  if (!match) return null;

  const [, year, month, day, hour, minute] = match;
  return instantOf(
    {
      year: Number(year),
      month: Number(month),
      day: Number(day),
      hour: Number(hour),
      minute: Number(minute),
    },
    timeZone,
  );
}

export function eventDurationHours(e: DigestEvent, timeZone: string): number {
  const start = eventInstant(e.start, timeZone);
  const end = eventInstant(e.end, timeZone);
  if (!start || !end) return 0;
  const hours = (end.getTime() - start.getTime()) / 3_600_000;
  return hours > 0 ? hours : 0;
}

export function eventResponse(e: DigestEvent): string {
  return e.responseStatus?.response ?? 'none';
}

export function roundTo(value: number, places: number): number {
  const factor = 10 ** places;
  return Math.round(value * factor) / factor;
}
