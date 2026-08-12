/**
 * Telling bulk and robot mail apart from a person.
 *
 * This is the load-bearing filter for "waiting on you": a nag list that includes
 * newsletters and no-reply addresses gets ignored inside two weeks. Every rule
 * here protects precision at the cost of recall, because silently dropping a real
 * person from the waiting list is the worst failure this tool can have.
 */

/**
 * The usual automated traffic.
 *
 * Only excluded from "waiting on you" - these still appear under unread, because
 * "I never opened it" is true regardless of who sent it.
 *
 * Notably absent: `info@` and `events@`. At a small nonprofit those are often
 * staffed by an actual person who does expect a reply.
 */
export const DEFAULT_IGNORED_SENDERS = [
  'no-reply@',
  'noreply@',
  'donotreply@',
  'do-not-reply@',
  'no_reply@',
  'notifications@',
  'notification@',
  'mailer-daemon@',
  'postmaster@',
  'automated@',
  'auto@',
  'alerts@',
  'alert@',
  'newsletter@',
  'news@',
  'updates@',
  'update@',
  'bulletin@',
  'digest@',
  // Found by test/scenario.ts, where a Candid webinar invitation surfaced as
  // "waiting on a reply". Nobody at webinars@ is waiting to hear back.
  'webinars@',
  'webinar@',
  'events-noreply@',
  'invitations@',
  'announcements@',
  'announce@',
  'marketing@',
  'campaigns@',
  'bounce',
  'unsubscribe',
  '@mailchimp',
  '@sendgrid',
  '@constantcontact',
  '@salsalabs',
  '@mailgun',
  '@sparkpostmail',
  '@amazonses',
];

/**
 * Whether the local part is exactly `base`, or `base` followed by a separator.
 *
 * So "noreply" matches "noreply-service" and "noreply.2", but not
 * "noreplyingtoyou" - which is a contrived example standing in for a real class
 * of false positive.
 */
function localPartMatches(local: string, base: string): boolean {
  if (base === '') return false;
  if (local === base) return true;
  if (!local.startsWith(base)) return false;
  return '-._+'.includes(local[base.length] as string);
}

/**
 * Whether an address looks like bulk or robot mail.
 *
 * Matching is anchored rather than a naive substring test, because a substring
 * test is quietly dangerous here: the pattern "news@" would also swallow
 * "goodnews@apersonsdomain.org".
 *
 * Three pattern shapes:
 *   - "@example.org" matches anywhere in the domain
 *   - "news@"        matches the local part, anchored at its start
 *   - "bounce"       falls back to a plain substring match
 */
export function isAutomatedSender(address: string, patterns: string[]): boolean {
  const normalized = address.trim().toLowerCase();
  if (normalized === '') return false;

  const at = normalized.lastIndexOf('@');
  const local = at >= 0 ? normalized.slice(0, at) : normalized;
  const domain = at >= 0 ? normalized.slice(at) : '';

  for (const raw of patterns) {
    const pattern = raw.trim().toLowerCase();
    if (pattern === '') continue;

    if (pattern.startsWith('@')) {
      if (domain.includes(pattern)) return true;
    } else if (pattern.endsWith('@')) {
      if (localPartMatches(local, pattern.slice(0, -1))) return true;
    } else if (normalized.includes(pattern)) {
      return true;
    }
  }
  return false;
}
