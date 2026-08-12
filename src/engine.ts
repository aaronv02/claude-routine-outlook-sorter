import type { Category, Correction, Decision, MailSummary, Suggestion } from './types.js';
import { NEEDS_REVIEW_ID } from './taxonomy.js';

/**
 * The parts of the sort loop that don't need a model.
 *
 * Three layers decide where mail goes, cheapest first, and the ordering is the
 * whole economics of the tool:
 *
 *  1. Sender rules - free, instant, deterministic. Absorbs the repetitive
 *     majority. Applied in `routine/sweep.ts`.
 *  2. Claude - only for senders no rule covers. This is the layer the routine
 *     prompt performs; there is no code for it here, which is the point.
 *  3. A confidence gate - see `classify/confidence.ts`. Anything that doesn't
 *     clear the bar becomes ⚠ Needs Review instead of a confident wrong guess.
 *
 * Over time layer 1 grows and layer 2 shrinks toward nothing. What remains in
 * this file is the feedback machinery: noticing when she disagrees, and deciding
 * what is worth asking about again.
 */

/**
 * Find corrections the user made in Outlook itself.
 *
 * She changes a category the normal way, we notice on the next pass. This is the
 * feedback path that matters, because it costs her nothing to use - there is no
 * widget to discover and no habit to form. It works by comparing our provenance
 * stamp against the categories actually on the message.
 *
 * A removed category is not treated as a correction: it says "not this" without
 * saying "but that", and inventing a target from silence would poison the rules.
 */
export function detectCorrections(mail: MailSummary[], taxonomy: Category[]): Correction[] {
  const nameToId = new Map(taxonomy.map((c) => [c.name.trim().toLowerCase(), c.id] as const));
  const corrections: Correction[] = [];

  for (const item of mail) {
    if (!item.assigned) continue;

    const presentIds = item.categories
      .map((name) => nameToId.get(name.trim().toLowerCase()))
      .filter((id): id is string => Boolean(id));

    const realIds = presentIds.filter((id) => id !== NEEDS_REVIEW_ID);
    if (realIds.length === 0) continue;

    // Still carrying what we assigned - no disagreement.
    if (realIds.includes(item.assigned.categoryId)) continue;

    // Exactly one replacement is an unambiguous correction. Several of our
    // categories at once means she's using them as multi-labels, which is a
    // legitimate choice but not a signal we can learn a single mapping from.
    if (realIds.length !== 1) continue;

    corrections.push({
      sender: item.from,
      subject: item.subject,
      fromCategoryId: item.assigned.categoryId,
      toCategoryId: realIds[0] as string,
      at: new Date().toISOString(),
    });
  }

  return corrections;
}

/**
 * Messages worth spending a classification on: unlabelled, or previously gated
 * and something has been learned since.
 *
 * That second clause is the whole point. Retrying every Needs Review on every
 * run sounds harmless and is not: two sweeps in a row would re-read the same
 * unresolvable messages and pay to reach the same answer. A retry is only worth
 * anything if the inputs changed - a new sender rule, a correction, an edited
 * category description - and `generation` is exactly the record of that having
 * happened.
 */
export function needsClassification(
  mail: MailSummary[],
  taxonomy: Category[],
  generation: number,
): MailSummary[] {
  const ourNames = new Set(taxonomy.map((c) => c.name.trim().toLowerCase()));
  const reviewName = taxonomy
    .find((c) => c.id === NEEDS_REVIEW_ID)
    ?.name.trim()
    .toLowerCase();

  return mail.filter((item) => {
    const ours = item.categories.filter((n) => ourNames.has(n.trim().toLowerCase()));
    if (ours.length === 0) return true;

    const onlyNeedsReview = ours.every((n) => n.trim().toLowerCase() === reviewName);
    if (!onlyNeedsReview) return false;

    // No stamp means it was parked by something other than this tool, or by a
    // build from before generations existed. Reconsider it once; the stamp we
    // write then carries the current generation and settles it.
    if (!item.assigned) return true;

    return item.assigned.generation < generation;
  });
}

/** The top suggestion, or null when unresolved. */
export function topSuggestion(decision: Decision): Suggestion | null {
  const first = decision.ranked[0];
  if (!first || first.categoryId === NEEDS_REVIEW_ID) return null;
  return first;
}
