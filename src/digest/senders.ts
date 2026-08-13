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
  // Seen in a real inbox audit, all unambiguously bulk.
  'informational@',
  'insiderdeals@',
  'offers@',
  'deals@',
  'survey@',
  'surveys@',
  'promotions@',
  'promo@',
  // NOT bare 'noreply' - that falls through to the substring branch and would
  // swallow "noreplyingtoyou@". The suffix half of localPartMatches already
  // catches "googleone-noreply@" via the 'noreply@' pattern above.
  '@mailchimp',
  '@sendgrid',
  '@constantcontact',
  '@salsalabs',
  '@mailgun',
  '@sparkpostmail',
  '@amazonses',
];

/**
 * Whether the local part is exactly `base`, or `base` adjoining a separator on
 * either side.
 *
 * So "noreply" matches "noreply-service" and "noreply.2", and also
 * "googleone-noreply" - a real address from a real inbox that the
 * prefix-only version missed. It still does not match "noreplyingtoyou", which
 * stands in for a real class of false positive.
 */
function localPartMatches(local: string, base: string): boolean {
  if (base === '') return false;
  if (local === base) return true;

  const separators = '-._+';
  if (local.startsWith(base) && separators.includes(local[base.length] as string)) return true;
  if (local.endsWith(base) && separators.includes(local[local.length - base.length - 1] as string)) {
    return true;
  }
  return false;
}

/**
 * Leftmost domain labels that mean "this came out of a bulk sending platform".
 *
 * Auditing a real 60-day inbox showed the local-part patterns above catching only
 * about a third of the bulk senders actually present. The reason is a shift in how
 * marketing mail is addressed: the local part is now the *brand* -
 * `venmo@`, `ebay@`, `samsung@`, `HarborFreight@` - and the sending platform shows
 * up in a subdomain instead. `em.oakley.com`, `reply.ebay.com`,
 * `promos.discounttire.com`, `notify.wellsfargo.com`, `engage.canva.com`.
 *
 * Deliberately excluded, despite appearing in that inbox: `mail`, `email`, `e`,
 * `m`, `t`, `my`, `go`, `service` and `communication`. Those are short or generic
 * enough that real people do send from them - `mail.some-university.edu` is a real
 * shape - and dropping a real person from the waiting list is the worst failure
 * this tool can have. Better to miss some bulk than to hide one person.
 */
const BULK_SUBDOMAINS = new Set([
  'em', 'em1', 'em2', 'em3', 'em4', 'em5',
  'mail1', 'mail2', 'mail3', 'mail4', 'mail5',
  'mailer', 'mailers', 'mailing', 'mailings',
  'reply', 'replies', 'noreply', 'no-reply', 'donotreply',
  'notify', 'notification', 'notifications', 'notifs',
  'infomails', 'infomail',
  'promo', 'promos', 'offers', 'deals',
  'marketing', 'mktg', 'mkt', 'engage', 'engagement',
  'campaign', 'campaigns',
  'click', 'clicks', 'links', 'trk', 'track',
  'sendgrid', 'sparkpost', 'mailgun',
  'messaging', 'transactional',
  'updates', 'update', 'alerts', 'alert',
  'members', 'voices', 'customerfeedback',
]);

/**
 * Whether the domain is a bulk sending subdomain, e.g. `em.oakley.com`.
 *
 * Requires at least three labels, so a bare `oakley.com` is never caught this way -
 * only a dedicated sending subdomain underneath it.
 */
function isBulkSendingDomain(domain: string): boolean {
  const labels = domain.replace(/^@/, '').split('.');
  if (labels.length < 3) return false;
  return BULK_SUBDOMAINS.has(labels[0] as string);
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

  if (isBulkSendingDomain(domain)) return true;

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
