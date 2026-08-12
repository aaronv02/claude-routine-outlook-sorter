import type { Decision, Suggestion } from '../types.js';
import { NEEDS_REVIEW_ID } from '../taxonomy.js';

/**
 * Layer 3: the honesty gate.
 *
 * A confident wrong label is worse than an admitted unknown - it teaches the
 * user the tool can't be trusted, and they stop looking at the labels at all.
 * So anything that doesn't clear the threshold becomes Needs Review.
 */

/** Number of alternatives kept per message. */
export const RANKED_LIMIT = 3;

/**
 * Turn raw suggestions into a decision.
 *
 * Ranked alternatives are preserved even when gated. Only the top one is ever
 * written to the mailbox, but the runners-up are what the digest reports when a
 * message lands in Needs Review, so she can see what it was torn between.
 */
export function gate(messageId: string, ranked: Suggestion[], threshold: number): Decision {
  const cleaned = ranked
    .filter((s) => s.categoryId !== NEEDS_REVIEW_ID)
    .sort((a, b) => b.confidence - a.confidence)
    .slice(0, RANKED_LIMIT);

  if (cleaned.length === 0) {
    return {
      messageId,
      ranked: [
        { categoryId: NEEDS_REVIEW_ID, confidence: 0, reason: 'No category could be determined.' },
      ],
      source: 'unresolved',
      gated: true,
    };
  }

  const top = cleaned[0] as Suggestion;
  const runnerUp = cleaned[1];

  // Two separate ways to be unsure, and both matter. An absolute low score means
  // nothing fits; a narrow margin between the top two means something fits but
  // we can't tell which. Either one should defer to the user.
  const belowThreshold = top.confidence < threshold;
  const tooClose = runnerUp !== undefined && top.confidence - runnerUp.confidence < 0.1;

  return {
    messageId,
    ranked: cleaned,
    source: 'llm',
    gated: belowThreshold || tooClose,
  };
}

/** The category to actually write, honouring the gate. */
export function categoryToApply(decision: Decision): string {
  if (decision.gated) return NEEDS_REVIEW_ID;
  return (decision.ranked[0] as Suggestion).categoryId;
}
