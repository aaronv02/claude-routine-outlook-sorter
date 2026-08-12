import { isAutomatedSender } from './senders.js';
import {
  addressedToAny,
  fromAddress,
  normalizeAddress,
  receivedAt,
  sentAt,
  type DigestMessage,
} from './types.js';

/** One message that appears to still need a reply. */
export interface WaitingItem {
  message: DigestMessage;
  /** Whole days it has gone unanswered. */
  ageDays: number;
}

/**
 * Mail that was addressed to her and never answered.
 *
 * This is the section with the most value and the most ways to be wrong. A nag
 * list that includes newsletters, CCs, and things already handled gets ignored
 * within two weeks, so every filter here exists to protect precision at the cost
 * of recall. Better to under-report than to cry wolf.
 *
 * A message is waiting when all of these hold:
 *   - she is on the To line, not merely CC'd (being copied is not being asked)
 *   - the sender looks like a person, not a mailing list or a robot
 *   - nothing was sent from her mailbox in that conversation after it arrived
 *   - it is older than the grace period, so this morning's mail is not a failure
 *
 * `sent` should reach further back than `inbox`: a reply sent last week still
 * answers a message received last week.
 *
 * `addresses` is every address she might legitimately be reached at - her primary
 * plus any role aliases or shared mailboxes.
 */
export function findWaiting(
  inbox: DigestMessage[],
  sent: DigestMessage[],
  addresses: string[],
  graceMs: number,
  now: Date,
  ignoredPatterns: string[],
): WaitingItem[] {
  const own = new Set(addresses.map(normalizeAddress).filter(Boolean));

  // Latest outbound activity per conversation.
  const repliedAt = new Map<string, number>();
  for (const m of sent) {
    const conversation = m.conversationId;
    if (!conversation) continue;
    const when = sentAt(m).getTime();
    const existing = repliedAt.get(conversation);
    if (existing === undefined || when > existing) repliedAt.set(conversation, when);
  }

  const items: WaitingItem[] = [];
  // Conversation id -> index into items, so a thread collapses to one entry.
  const byConversation = new Map<string, number>();

  for (const m of inbox) {
    if (m.isDraft) continue;

    const from = fromAddress(m);
    if (from === '' || own.has(from)) continue;
    if (!addressedToAny(m, addresses)) continue;
    if (isAutomatedSender(from, ignoredPatterns)) continue;

    const received = receivedAt(m).getTime();
    if (Number.isNaN(received)) continue;

    const age = now.getTime() - received;
    if (age < graceMs) continue;

    const replied = repliedAt.get(m.conversationId ?? '');
    if (replied !== undefined && replied > received) continue;

    const item: WaitingItem = { message: m, ageDays: Math.floor(age / 86_400_000) };

    // One entry per thread: a five-message thread she never answered is one thing
    // to do, not five. The oldest message is kept as the representative, because
    // the honest answer to "how long have they been waiting?" is measured from
    // the first unanswered ask, not the latest nudge.
    const conversation = m.conversationId;
    if (conversation) {
      const existingIndex = byConversation.get(conversation);
      if (existingIndex !== undefined) {
        const incumbent = items[existingIndex] as WaitingItem;
        if (received < receivedAt(incumbent.message).getTime()) items[existingIndex] = item;
        continue;
      }
      byConversation.set(conversation, items.length);
    }

    items.push(item);
  }

  // Oldest first. The thing that has been sitting longest is the thing most
  // likely to have become a problem.
  return items.sort(
    (a, b) => receivedAt(a.message).getTime() - receivedAt(b.message).getTime(),
  );
}

/**
 * Messages she flagged and never cleared.
 *
 * A flag is her own explicit "come back to this", which makes it a stronger
 * signal than anything this tool infers - so these are reported regardless of
 * age, sender, or whether she was on the To line.
 *
 * Anything already listed as waiting is omitted. The same item appearing twice in
 * one report reads as a bug, and pads a list whose whole value is that it is short
 * and true.
 */
export function findOpenFlags(
  inbox: DigestMessage[],
  waiting: WaitingItem[],
): DigestMessage[] {
  const listedIds = new Set(waiting.map((w) => w.message.id));
  const listedConversations = new Set(
    waiting.map((w) => w.message.conversationId).filter((c): c is string => Boolean(c)),
  );

  return inbox
    .filter((m) => m.flag?.flagStatus === 'flagged')
    .filter((m) => !listedIds.has(m.id))
    .filter((m) => !(m.conversationId && listedConversations.has(m.conversationId)))
    .sort((a, b) => receivedAt(a).getTime() - receivedAt(b).getTime());
}
