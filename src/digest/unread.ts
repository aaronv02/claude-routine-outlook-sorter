import { isAutomatedSender } from './senders.js';
import { fromAddress, fromName, normalizeAddress, receivedAt, type DigestMessage } from './types.js';

/** Unread mail from one sender. */
export interface SenderGroup {
  address: string;
  name: string;
  count: number;
  oldest: string;
  newest: string;
  /** A few subjects for context, not the whole list. */
  subjects: string[];
  automated: boolean;
}

export interface UnreadReport {
  people: SenderGroup[];
  automated: SenderGroup[];
  total: number;
  /** Unread older than a week - the number actually worth worrying about. */
  staleCount: number;
}

const STALE_MS = 7 * 86_400_000;
const SUBJECTS_PER_SENDER = 3;

/**
 * Everything she never opened, grouped by sender and split into people versus bulk.
 *
 * Grouping is the entire point. "37 unread" is a number that produces guilt and
 * no action; "Tessa Nunn sent 3 you haven't opened, plus 22 newsletters" is
 * something you can do something about in thirty seconds.
 */
export function buildUnread(
  inbox: DigestMessage[],
  addresses: string[],
  now: Date,
  ignoredPatterns: string[],
): UnreadReport {
  const own = new Set(addresses.map(normalizeAddress).filter(Boolean));
  const groups = new Map<string, SenderGroup>();

  let total = 0;
  let staleCount = 0;

  for (const m of inbox) {
    if (m.isRead || m.isDraft) continue;

    const from = fromAddress(m);
    if (from === '' || own.has(from)) continue;

    const received = receivedAt(m);
    if (Number.isNaN(received.getTime())) continue;

    total++;
    if (now.getTime() - received.getTime() > STALE_MS) staleCount++;

    let group = groups.get(from);
    if (!group) {
      group = {
        address: from,
        name: fromName(m),
        count: 0,
        oldest: m.receivedDateTime,
        newest: m.receivedDateTime,
        subjects: [],
        automated: isAutomatedSender(from, ignoredPatterns),
      };
      groups.set(from, group);
    }

    group.count++;
    if (received.getTime() < new Date(group.oldest).getTime()) group.oldest = m.receivedDateTime;
    if (received.getTime() > new Date(group.newest).getTime()) group.newest = m.receivedDateTime;
    if (group.subjects.length < SUBJECTS_PER_SENDER) {
      group.subjects.push(m.subject?.trim() || '(no subject)');
    }
  }

  const all = [...groups.values()];

  return {
    // People by volume, then by age. Bulk purely by volume, since it is really
    // just a cleanup list.
    people: all
      .filter((g) => !g.automated)
      .sort((a, b) => b.count - a.count || new Date(a.oldest).getTime() - new Date(b.oldest).getTime()),
    automated: all.filter((g) => g.automated).sort((a, b) => b.count - a.count),
    total,
    staleCount,
  };
}
