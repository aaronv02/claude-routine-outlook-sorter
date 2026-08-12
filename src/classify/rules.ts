import type { Category, Correction, MailSummary, SenderRule, Suggestion } from '../types.js';
import { NEEDS_REVIEW_ID, mapExistingCategories } from '../taxonomy.js';

/**
 * Layer 1: learned sender rules.
 *
 * Free, instant, deterministic, and by far the highest-precision signal in
 * personal email - who sent it predicts where it belongs better than anything in
 * the body. This layer exists to keep the LLM's scarce free-tier quota for mail
 * that actually needs judgement, and it doubles as the pool of candidates for
 * promotion into native Outlook rules.
 */

export function normalizeSender(address: string): string {
  return address.trim().toLowerCase();
}

/** "@example.org" for "person@example.org". Empty string if unparseable. */
export function domainOf(address: string): string {
  const at = address.lastIndexOf('@');
  return at === -1 ? '' : address.slice(at).toLowerCase();
}

/**
 * Best matching rule for a sender.
 *
 * An exact address always beats a domain rule. That ordering matters at a
 * foundation: mail from `@coloradogives.org` is Finance in general, but a named
 * relationship manager there might be filed differently, and the specific rule
 * must win.
 */
export function matchSender(rules: SenderRule[], from: string): SenderRule | undefined {
  const sender = normalizeSender(from);
  if (!sender) return undefined;

  const exact = rules.find((r) => r.pattern === sender);
  if (exact) return exact;

  const domain = domainOf(sender);
  if (!domain) return undefined;
  return rules.find((r) => r.pattern === domain);
}

/** Rule-layer verdict, or null when this sender is unknown. */
export function classifyByRule(rules: SenderRule[], mail: MailSummary): Suggestion | null {
  const rule = matchSender(rules, mail.from);
  if (!rule) return null;

  // Confidence grows with corroboration but is capped below 1: a sender rule is
  // a strong prior, not a certainty, and people do change what they email about.
  const corroboration = rule.confirmations * 2 + Math.min(rule.hits, 20) / 10;
  const confidence = Math.min(0.97, 0.72 + corroboration * 0.05);

  return {
    categoryId: rule.categoryId,
    confidence,
    reason:
      rule.confirmations > 0
        ? `You've filed ${rule.pattern} here ${rule.confirmations} time(s).`
        : `Learned from previous mail from ${rule.pattern}.`,
  };
}

/**
 * Add or reinforce a rule.
 *
 * `confirmed` distinguishes an explicit user correction from our own inference.
 * Only confirmations count toward promotion into a native Outlook rule, because
 * promoting a guess would bake a mistake into the mailbox where it keeps
 * applying itself long after the add-in is closed.
 */
export function upsertSenderRule(
  rules: SenderRule[],
  pattern: string,
  categoryId: string,
  confirmed: boolean,
): SenderRule[] {
  const key = normalizeSender(pattern);
  const next = [...rules];
  const idx = next.findIndex((r) => r.pattern === key);

  if (idx === -1) {
    next.push({
      pattern: key,
      categoryId,
      hits: confirmed ? 0 : 1,
      confirmations: confirmed ? 1 : 0,
      promoted: false,
      createdAt: new Date().toISOString(),
    });
    return next;
  }

  const existing = next[idx] as SenderRule;
  if (existing.categoryId !== categoryId) {
    // The user changed their mind about this sender. Reset the evidence rather
    // than averaging two contradictory histories, and un-promote so the stale
    // native Outlook rule gets rewritten on the next promotion pass.
    next[idx] = {
      ...existing,
      categoryId,
      hits: 0,
      confirmations: confirmed ? 1 : 0,
      promoted: false,
    };
    return next;
  }

  next[idx] = {
    ...existing,
    hits: existing.hits + (confirmed ? 0 : 1),
    confirmations: existing.confirmations + (confirmed ? 1 : 0),
  };
  return next;
}

/**
 * Fold a user correction into the rule set.
 *
 * Deliberately learns the exact address, never the domain. Inferring
 * "@gmail.com means Grants" from one grantseeker would be catastrophic;
 * domain rules are only ever created explicitly by the user in the UI.
 */
export function learnFromCorrection(rules: SenderRule[], correction: Correction): SenderRule[] {
  return upsertSenderRule(rules, correction.sender, correction.toCategoryId, true);
}

/**
 * Domains where the address says nothing about the subject.
 *
 * A domain rule for one of these would be a catastrophe rather than a
 * shortcut - it would file every private individual she corresponds with into
 * whichever category her grantseekers happen to use.
 */
const SHARED_DOMAINS = new Set([
  '@gmail.com',
  '@googlemail.com',
  '@outlook.com',
  '@hotmail.com',
  '@live.com',
  '@yahoo.com',
  '@ymail.com',
  '@aol.com',
  '@icloud.com',
  '@me.com',
  '@mac.com',
  '@msn.com',
  '@comcast.net',
  '@proton.me',
  '@protonmail.com',
]);

/**
 * Distinct senders at a domain that must agree before a domain rule is drawn.
 *
 * Two is a coincidence. Three people at the same organization all filed the
 * same way is a pattern about the organization.
 */
const MIN_SENDERS_PER_DOMAIN = 3;

/**
 * Share of a domain's senders that must agree.
 *
 * Deliberately a majority rather than unanimity. Requiring every sender to
 * agree threw away the whole domain over a single exception - and a single
 * exception is the normal shape of these relationships: everyone at a giving
 * platform is Finance except the one relationship manager. `matchSender`
 * already resolves that correctly, because an exact rule beats a domain rule,
 * so the exception keeps its own rule and the other twenty senders still cost
 * nothing to classify.
 */
const DOMAIN_MAJORITY = 0.75;

export interface BootstrapResult {
  rules: SenderRule[];
  /**
   * Her category names we matched to ours by similarity rather than exactly,
   * so the first run can show what it assumed.
   */
  inferred: { from: string; toCategoryId: string }[];
  /** How many senders were learned from her existing filing. */
  learned: number;
  /** Domains covered wholesale, e.g. "@coloradogives.org". */
  domains: string[];
}

/**
 * Seed rules from mail the user has already categorized.
 *
 * Her existing categories *are* a labelled training set, so first run should
 * mine them silently before asking her anything. Only senders that map
 * unambiguously to a single category are taken - a sender seen under two
 * different categories teaches us nothing reliable and is left to the LLM.
 *
 * Matching her category names to ours is fuzzy, and has to be: our seed names
 * are specific ("Donors & Gifts", "Board & Governance") while hers will be
 * short ("Donors", "Board"). Requiring an exact match meant this harvested
 * almost nothing, so day one began with no learned senders at all and sent the
 * entire inbox to the LLM - straight into the free daily cap.
 *
 * The matcher declines whenever there isn't a clear winner. Categories that are
 * genuinely hers alone - a fund name, a person, a project - stay unmapped, and
 * their mail is classified normally.
 */
export function bootstrapSenderRules(
  history: MailSummary[],
  taxonomy: Category[],
  existing: SenderRule[],
): BootstrapResult {
  const { lookup, inferred } = mapExistingCategories(
    history.flatMap((m) => m.categories),
    taxonomy,
  );

  const observed = new Map<string, Set<string>>();
  for (const mail of history) {
    const sender = normalizeSender(mail.from);
    if (!sender) continue;
    for (const catName of mail.categories) {
      const id = lookup.get(catName.trim().toLowerCase());
      if (!id || id === NEEDS_REVIEW_ID) continue;
      const set = observed.get(sender) ?? new Set<string>();
      set.add(id);
      observed.set(sender, set);
    }
  }

  let rules = existing;
  let learned = 0;
  // Only categories that actually produced a rule. A mapping that resolved but
  // taught us nothing - because every sender under it was ambiguous - is not an
  // assumption worth putting on the one screen she reads.
  const used = new Set<string>();
  // Unambiguous senders, grouped by domain, for the domain pass below.
  const byDomain = new Map<string, Map<string, string>>();

  for (const [sender, categories] of observed) {
    if (categories.size !== 1) continue;
    const [only] = [...categories];
    if (!only) continue;

    const domain = domainOf(sender);
    if (domain && !SHARED_DOMAINS.has(domain)) {
      const group = byDomain.get(domain) ?? new Map<string, string>();
      group.set(sender, only);
      byDomain.set(domain, group);
    }

    if (rules.some((r) => r.pattern === sender)) continue;
    rules = upsertSenderRule(rules, sender, only, true);
    used.add(only);
    learned++;
  }

  // Domain pass. At a foundation most correspondents are institutions, so a
  // domain that files unanimously is the single highest-leverage thing layer 1
  // can learn: it covers colleagues she has never heard from yet, which is
  // exactly the mail that would otherwise cost an LLM call.
  const domains: string[] = [];
  for (const [domain, senders] of byDomain) {
    if (senders.size < MIN_SENDERS_PER_DOMAIN) continue;

    const tally = new Map<string, number>();
    for (const categoryId of senders.values()) {
      tally.set(categoryId, (tally.get(categoryId) ?? 0) + 1);
    }

    let only: string | undefined;
    let best = 0;
    for (const [categoryId, count] of tally) {
      if (count > best) {
        best = count;
        only = categoryId;
      }
    }

    if (!only) continue;
    // Both bars, not either: enough senders to be a pattern, and a clear
    // enough majority that the domain means one thing with exceptions rather
    // than genuinely mixed traffic.
    if (best < MIN_SENDERS_PER_DOMAIN) continue;
    if (best / senders.size < DOMAIN_MAJORITY) continue;
    if (rules.some((r) => r.pattern === domain)) continue;

    // Inferred, not confirmed. It is strong enough to classify on - which is
    // where the quota saving comes from - but promoting a domain into a
    // permanent native Outlook rule off our own inference would apply itself
    // to strangers indefinitely. That still requires her say-so.
    rules = upsertSenderRule(rules, domain, only, false);
    used.add(only);
    domains.push(domain);
  }
  return {
    rules,
    inferred: inferred.filter((i) => used.has(i.toCategoryId)),
    learned,
    domains,
  };
}

/** Senders that have earned a native Outlook rule but don't have one yet. */
export function promotionCandidates(rules: SenderRule[], threshold: number): SenderRule[] {
  return rules.filter(
    (r) => !r.promoted && r.categoryId !== NEEDS_REVIEW_ID && r.confirmations >= threshold,
  );
}
