/**
 * Offline checks. No network, no mailbox, no API key.
 *
 *   npm test
 *
 * The target mailbox isn't available to us, so everything that decides what gets
 * labelled, what gets learned, and what gets promoted has to be verifiable
 * without an inbox. What's covered here is deliberately the destructive and the
 * irreversible: writes that could drop one of her own categories, and rules that
 * keep applying themselves in Outlook long after a sweep ends.
 *
 * The classification itself is not tested here - it's a model, performed by the
 * routine prompt. Its safeguard is the confidence gate, which is.
 */

import type { Correction, MailSummary, SenderRule } from '../src/types.js';
import {
  NEEDS_REVIEW_ID,
  SEED_TAXONOMY,
  matchExistingCategory,
  normalizeTaxonomy,
  sanitizeCategoryName,
} from '../src/taxonomy.js';
import {
  bootstrapSenderRules,
  classifyByRule,
  learnFromCorrection,
  promotionCandidates,
  upsertSenderRule,
} from '../src/classify/rules.js';
import { categoryToApply, gate } from '../src/classify/confidence.js';
import { detectCorrections, needsClassification } from '../src/engine.js';
import { mergeCategories } from '../routine/merge.js';
import { runDigestChecks, runDesktopChecks, runStateChecks } from './digest.js';

let passed = 0;
const failures: string[] = [];

function check(name: string, fn: () => void): void {
  try {
    fn();
    passed++;
  } catch (err) {
    failures.push(`${name}: ${err instanceof Error ? err.message : String(err)}`);
  }
}

async function checkAsync(name: string, fn: () => Promise<void>): Promise<void> {
  try {
    await fn();
    passed++;
  } catch (err) {
    failures.push(`${name}: ${err instanceof Error ? err.message : String(err)}`);
  }
}

function assert(condition: unknown, message: string): void {
  if (!condition) throw new Error(message);
}

const REVIEW_NAME = SEED_TAXONOMY.find((c) => c.id === NEEDS_REVIEW_ID)?.name as string;

function mail(over: Partial<MailSummary> = {}): MailSummary {
  return {
    id: 'msg-1',
    from: 'someone@example.org',
    fromName: 'Someone',
    subject: 'A subject',
    received: '2026-01-01T00:00:00Z',
    hasAttachments: false,
    categories: [],
    ...over,
  };
}

// ---------------------------------------------------------------------------
// Writing categories without destroying hers
//
// Graph PATCH replaces the whole categories array, so this is the one place that
// can silently delete something she did by hand.
// ---------------------------------------------------------------------------

check('keeps categories that are not ours', () => {
  const out = mergeCategories(['Smith Family Fund', 'Grants'], SEED_TAXONOMY, 'Donors & Gifts');
  assert(out.includes('Smith Family Fund'), 'dropped a category of hers');
});

check('replaces our previous label rather than accumulating', () => {
  const out = mergeCategories(['Grants'], SEED_TAXONOMY, 'Donors & Gifts');
  assert(!out.includes('Grants'), 'kept a stale label of ours');
  assert(out.length === 1 && out[0] === 'Donors & Gifts', `unexpected: ${out.join(', ')}`);
});

check('does not duplicate a label she already applied herself', () => {
  // Case differs only. Outlook compares names case-insensitively, so appending
  // would leave the mailbox holding two spellings of one label.
  const out = mergeCategories(['events'], [], 'Events');
  assert(out.length === 1, `expected one category, got: ${out.join(', ')}`);
});

check('clears Needs Review when a message is resolved, keeps it when re-gated', () => {
  const resolved = mergeCategories([REVIEW_NAME], SEED_TAXONOMY, 'Grants');
  assert(!resolved.includes(REVIEW_NAME), 'left Needs Review on a resolved message');
  const still = mergeCategories([REVIEW_NAME], SEED_TAXONOMY, REVIEW_NAME);
  assert(still.length === 1 && still[0] === REVIEW_NAME, 'lost Needs Review on a re-gated message');
});

// ---------------------------------------------------------------------------
// The confidence gate - the only thing between a guess and a label
// ---------------------------------------------------------------------------

check('a confident, clear winner is written', () => {
  const decision = gate('m', [{ categoryId: 'grants', confidence: 0.9, reason: 'r' }], 0.65);
  assert(!decision.gated, 'gated a confident verdict');
  assert(categoryToApply(decision) === 'grants', 'wrote the wrong category');
});

check('below the threshold becomes Needs Review', () => {
  const decision = gate('m', [{ categoryId: 'grants', confidence: 0.4, reason: 'r' }], 0.65);
  assert(decision.gated, 'let a low-confidence guess through');
  assert(categoryToApply(decision) === NEEDS_REVIEW_ID, 'did not fall back to Needs Review');
});

check('a near-tie becomes Needs Review even when confident', () => {
  // Two ways to be unsure and both matter: a low score means nothing fits, a
  // narrow margin means something fits but we cannot tell which.
  const decision = gate(
    'm',
    [
      { categoryId: 'grants', confidence: 0.8, reason: 'r' },
      { categoryId: 'partners', confidence: 0.75, reason: 'r' },
    ],
    0.65,
  );
  assert(decision.gated, 'let a coin-flip through as a confident label');
});

check('the gate never writes Needs Review as a ranked choice', () => {
  const decision = gate('m', [{ categoryId: NEEDS_REVIEW_ID, confidence: 0.99, reason: 'r' }], 0.65);
  assert(decision.gated, 'accepted Needs Review as a real verdict');
  assert(decision.ranked.every((r, i) => i > 0 || r.categoryId === NEEDS_REVIEW_ID), 'unexpected ranking');
});

// ---------------------------------------------------------------------------
// Noticing corrections she made in Outlook
// ---------------------------------------------------------------------------

const stamped = (categoryId: string, generation = 1) => ({
  categoryId,
  confidence: 0.8,
  at: '2026-01-01T00:00:00Z',
  generation,
});

check('a swapped category is read as a correction', () => {
  const found = detectCorrections(
    [mail({ categories: ['Donors & Gifts'], assigned: stamped('grants') })],
    SEED_TAXONOMY,
  );
  assert(found.length === 1, `expected one correction, got ${found.length}`);
  assert(found[0]?.toCategoryId === 'donors', `learned the wrong target: ${found[0]?.toCategoryId}`);
});

check('an unchanged label is not a correction', () => {
  const found = detectCorrections(
    [mail({ categories: ['Grants'], assigned: stamped('grants') })],
    SEED_TAXONOMY,
  );
  assert(found.length === 0, 'invented a correction from agreement');
});

check('a removed label is not a correction', () => {
  // It says "not this" without saying "but that". Inventing a target from
  // silence would poison the sender rules.
  const found = detectCorrections([mail({ categories: [], assigned: stamped('grants') })], SEED_TAXONOMY);
  assert(found.length === 0, 'invented a correction from a deletion');
});

check('two of our categories at once is multi-labelling, not a correction', () => {
  const found = detectCorrections(
    [mail({ categories: ['Donors & Gifts', 'Events'], assigned: stamped('grants') })],
    SEED_TAXONOMY,
  );
  assert(found.length === 0, 'read a multi-label as a single correction');
});

check('a category of her own alongside ours is not a correction', () => {
  const found = detectCorrections(
    [mail({ categories: ['Grants', 'Smith Family Fund'], assigned: stamped('grants') })],
    SEED_TAXONOMY,
  );
  assert(found.length === 0, 'read her own category as a disagreement');
});

// ---------------------------------------------------------------------------
// What is worth reconsidering on a later sweep
// ---------------------------------------------------------------------------

check('unlabelled mail is always considered', () => {
  const out = needsClassification([mail()], SEED_TAXONOMY, 5);
  assert(out.length === 1, 'skipped an unlabelled message');
});

check('mail we already labelled is left alone', () => {
  const out = needsClassification([mail({ categories: ['Grants'] })], SEED_TAXONOMY, 5);
  assert(out.length === 0, 're-read a message that was already sorted');
});

check('Needs Review is retried only once something has been learned', () => {
  const parked = mail({ categories: [REVIEW_NAME], assigned: stamped(NEEDS_REVIEW_ID, 3) });
  assert(
    needsClassification([parked], SEED_TAXONOMY, 3).length === 0,
    'paid to re-read an unresolvable message with nothing new learned',
  );
  assert(
    needsClassification([parked], SEED_TAXONOMY, 4).length === 1,
    'failed to reconsider after learning something',
  );
});

// ---------------------------------------------------------------------------
// Sender rules and promotion into native Outlook rules
//
// Promotion is the irreversible one: a native rule keeps applying itself in the
// mailbox whether or not a sweep ever runs again.
// ---------------------------------------------------------------------------

check('a known sender is resolved without a model', () => {
  const rules = upsertSenderRule([], 'board@example.org', 'board', false);
  const hit = classifyByRule(rules, mail({ from: 'board@example.org' }));
  assert(hit?.categoryId === 'board', 'failed to match a known sender');
  assert((hit?.confidence ?? 0) < 1, 'treated a sender rule as a certainty');
});

check('only confirmed senders are promoted', () => {
  let rules: SenderRule[] = [];
  // Our own inferences, three times over. Never promoted: a native rule that
  // bakes in a guess is a mistake that keeps applying itself invisibly.
  for (let i = 0; i < 3; i++) rules = upsertSenderRule(rules, 'guess@example.org', 'grants', false);
  assert(promotionCandidates(rules, 3).length === 0, 'promoted a sender we only guessed at');

  const correction: Correction = {
    sender: 'real@example.org',
    subject: 's',
    fromCategoryId: 'grants',
    toCategoryId: 'donors',
    at: '2026-01-01T00:00:00Z',
  };
  for (let i = 0; i < 3; i++) rules = learnFromCorrection(rules, correction);
  const candidates = promotionCandidates(rules, 3);
  assert(candidates.length === 1, `expected one candidate, got ${candidates.length}`);
  assert(candidates[0]?.pattern === 'real@example.org', 'promoted the wrong sender');
});

check('changing her mind about a sender resets its evidence and un-promotes it', () => {
  let rules = upsertSenderRule([], 'x@example.org', 'grants', true);
  rules = upsertSenderRule(rules, 'x@example.org', 'grants', true);
  rules[0] = { ...(rules[0] as SenderRule), promoted: true };
  rules = upsertSenderRule(rules, 'x@example.org', 'donors', true);
  assert(rules[0]?.categoryId === 'donors', 'ignored the new category');
  assert(rules[0]?.confirmations === 1, 'averaged two contradictory histories');
  assert(rules[0]?.promoted === false, 'left a stale native Outlook rule in place');
});

check('a correction learns the address, never the domain', () => {
  // Inferring "@gmail.com means Grants" from one grantseeker would be a disaster.
  const rules = learnFromCorrection([], {
    sender: 'someone@gmail.com',
    subject: 's',
    fromCategoryId: null,
    toCategoryId: 'grants',
    at: '2026-01-01T00:00:00Z',
  });
  assert(rules[0]?.pattern === 'someone@gmail.com', `learned ${rules[0]?.pattern}`);
});

// ---------------------------------------------------------------------------
// First run: reading how she already files mail
// ---------------------------------------------------------------------------

check('her shorter category names map onto ours', () => {
  assert(matchExistingCategory('Donors', SEED_TAXONOMY)?.categoryId === 'donors', 'missed "Donors"');
  assert(matchExistingCategory('Board', SEED_TAXONOMY)?.categoryId === 'board', 'missed "Board"');
});

check('categories that are hers alone are left unmapped', () => {
  assert(matchExistingCategory('Smith Family Fund', SEED_TAXONOMY) === null, 'claimed a fund name');
  assert(matchExistingCategory('Kathy', SEED_TAXONOMY) === null, 'claimed a person');
});

check('a sender filed under two categories teaches nothing', () => {
  const history = [
    mail({ id: 'a', from: 'mixed@example.org', categories: ['Grants'] }),
    mail({ id: 'b', from: 'mixed@example.org', categories: ['Events'] }),
    mail({ id: 'c', from: 'clear@example.org', categories: ['Grants'] }),
  ];
  const result = bootstrapSenderRules(history, SEED_TAXONOMY, []);
  const patterns = result.rules.map((r) => r.pattern);
  assert(!patterns.includes('mixed@example.org'), 'learned from an ambiguous sender');
  assert(patterns.includes('clear@example.org'), 'missed an unambiguous sender');
});

// ---------------------------------------------------------------------------
// Category names Outlook will not round-trip
// ---------------------------------------------------------------------------

check('a comma is removed from a category name', () => {
  // Outlook treats a category list as comma-delimited, so "Grants, Pending"
  // comes back as two categories and every name-to-id lookup breaks.
  assert(!sanitizeCategoryName('Grants, Pending').includes(','), 'kept a comma');
});

check('duplicate names are made unique', () => {
  const fixed = normalizeTaxonomy([
    { id: 'a', name: 'Grants', color: 'Preset0', description: '' },
    { id: 'b', name: 'Grants', color: 'Preset1', description: '' },
  ]);
  assert(fixed[0]?.name !== fixed[1]?.name, 'left two categories sharing one name');
});

// ---------------------------------------------------------------------------
// The weekly digest
// ---------------------------------------------------------------------------

runDigestChecks(check, assert);
await runDesktopChecks(checkAsync, assert);
await runStateChecks(check, assert);

// ---------------------------------------------------------------------------

console.log(`${passed} check(s) passed.`);
if (failures.length > 0) {
  console.log(`\n${failures.length} FAILED:`);
  for (const f of failures) console.log(`  ✗ ${f}`);
}
process.exit(failures.length > 0 ? 1 : 0);
