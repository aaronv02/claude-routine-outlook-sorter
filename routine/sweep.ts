/**
 * Headless driver for Inbox Steward, for running as a scheduled Claude routine.
 *
 * The add-in and this share one classifier. Everything that decides anything -
 * the taxonomy and its descriptions, sender rules, the confidence gate,
 * correction detection, promotion into native Outlook rules - is imported from
 * `src/`, not reimplemented. The only substitution is layer 2: instead of
 * calling Gemini from the browser, the routine hands the unresolved messages to
 * Claude and reads its answers back.
 *
 * That split is why this is two commands rather than one:
 *
 *   plan   - authenticate, learn from any corrections she made in Outlook, run
 *            layer 1, and write out the messages that still need judgement.
 *   apply  - take Claude's rankings, run them through the same confidence gate
 *            the add-in uses, write the categories, and promote what's earned it.
 *
 * Deterministic mailbox I/O stays in code where it is cheap and repeatable.
 * Claude is asked to do the one part that actually needs a model.
 */

import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';

import type { Correction, Decision, MailSummary, Suggestion } from '../src/types.js';
import {
  applyCategories,
  ensureMasterCategories,
  listCategorizedHistory,
  listRecentInbox,
  type CategoryUpdate,
} from '../src/graph.js';
import { detectCorrections, needsClassification } from '../src/engine.js';
import {
  bootstrapSenderRules,
  classifyByRule,
  learnFromCorrection,
  normalizeSender,
  upsertSenderRule,
} from '../src/classify/rules.js';
import { categoryToApply, gate } from '../src/classify/confidence.js';
import { markPromoted, promoteRules } from '../src/promote.js';
import { NEEDS_REVIEW_ID, categoryById, selectableCategories } from '../src/taxonomy.js';
import { loadEnv } from './env.js';
import { mergeCategories } from './merge.js';
import { loadState, saveState, type RoutineState } from './store.js';
import { assertConfigured, pollDeviceCode, redeemRefreshToken, startDeviceCode } from './auth.js';

const HERE = dirname(new URL(import.meta.url).pathname);
const PLAN_PATH = resolve(HERE, '.local/plan.json');
const VERDICTS_PATH = resolve(HERE, '.local/verdicts.json');

/** How much of the inbox one run looks at. */
const INBOX_WINDOW = Number(process.env.STEWARD_INBOX_WINDOW ?? '120');

/** Depth of the one-off first-run scan for already-filed mail. */
const HISTORY_DEPTH = 600;

/** Few-shot examples carried into the classification prompt. */
const EXAMPLE_LIMIT = 15;

/** Corrections retained in state. The add-in budgets the same way. */
const CORRECTION_MEMORY = 30;

// ---------------------------------------------------------------------------
// Token handling
// ---------------------------------------------------------------------------

/**
 * A refresh token has to come from somewhere on a cold start, and after that the
 * rotated one is kept in mailbox state so the 90-day window keeps rolling
 * without anyone touching the environment again.
 *
 * Reading the env var first is deliberate: it is the only way to recover when
 * the stored token has gone stale.
 */
async function authenticate(): Promise<{ token: string; state: RoutineState }> {
  // Configuration comes from .env unless the environment already carries it,
  // so a scheduled run can use secrets while a local run uses the file.
  await loadEnv();
  assertConfigured();
  const fromEnv = process.env.STEWARD_REFRESH_TOKEN;

  if (fromEnv) {
    const set = await redeemRefreshToken(fromEnv);
    const state = await loadState(set.accessToken);
    state.routine = { ...state.routine, refreshToken: set.refreshToken };
    return { token: set.accessToken, state };
  }

  // No env token: the only remaining copy is in the mailbox, and reading the
  // mailbox needs a token. Break the circle with a bootstrap file written by
  // `login`.
  const bootstrap = await readBootstrapToken();
  if (!bootstrap) {
    throw new Error(
      'No refresh token available. Run `npm run login` once, or set STEWARD_REFRESH_TOKEN.',
    );
  }
  const set = await redeemRefreshToken(bootstrap);
  const state = await loadState(set.accessToken);

  // Prefer whatever the mailbox has if it is newer than the bootstrap file.
  const stored = state.routine?.refreshToken;
  if (stored && stored !== bootstrap) {
    const rolled = await redeemRefreshToken(stored).catch(() => null);
    if (rolled) {
      state.routine = { ...state.routine, refreshToken: rolled.refreshToken };
      return { token: rolled.accessToken, state };
    }
  }

  state.routine = { ...state.routine, refreshToken: set.refreshToken };
  return { token: set.accessToken, state };
}

const BOOTSTRAP_PATH = resolve(HERE, '.local/refresh-token');

async function readBootstrapToken(): Promise<string | null> {
  try {
    return (await readFile(BOOTSTRAP_PATH, 'utf8')).trim() || null;
  } catch {
    return null;
  }
}

async function writeJson(path: string, value: unknown): Promise<void> {
  await mkdir(dirname(path), { recursive: true });
  await writeFile(path, JSON.stringify(value, null, 2), 'utf8');
}

// ---------------------------------------------------------------------------
// login
// ---------------------------------------------------------------------------

async function login(): Promise<void> {
  assertConfigured();
  const code = await startDeviceCode();
  console.log(`\n${code.message}\n`);
  console.log(`  URL:  ${code.verificationUri}`);
  console.log(`  Code: ${code.userCode}\n`);
  console.log('Waiting for sign-in...');

  const set = await pollDeviceCode(code);
  await mkdir(dirname(BOOTSTRAP_PATH), { recursive: true });
  await writeFile(BOOTSTRAP_PATH, set.refreshToken, { encoding: 'utf8', mode: 0o600 });

  const state = await loadState(set.accessToken);
  state.routine = { ...state.routine, refreshToken: set.refreshToken };
  await saveState(set.accessToken, state);

  console.log(`\nSigned in. Refresh token written to routine/.local/refresh-token (mode 600).`);
  console.log('For a cloud routine, put that value in STEWARD_REFRESH_TOKEN as a secret instead.\n');
}

// ---------------------------------------------------------------------------
// plan
// ---------------------------------------------------------------------------

interface PendingMessage {
  id: string;
  from: string;
  fromName: string;
  subject: string;
  received: string;
  hasAttachments: boolean;
  mailingList: boolean;
  preview?: string;
  /** Categories already on the message, so `apply` can preserve the ones that aren't ours. */
  existing: string[];
}

interface Plan {
  generatedAt: string;
  generation: number;
  confidenceThreshold: number;
  dataSharing: string;
  /** Layer 1 answers, already decided. Carried through so `apply` writes them in the same pass. */
  ruleDecisions: { id: string; categoryId: string; confidence: number; existing: string[] }[];
  categories: { id: string; name: string; description: string }[];
  examples: { sender: string; subject: string; category: string }[];
  pending: PendingMessage[];
  notes: string[];
}

async function plan(): Promise<void> {
  const { token, state } = await authenticate();
  const notes: string[] = [];

  // A category must exist in the mailbox master list before it can be applied,
  // so this is not housekeeping - skipping it makes every write silently useless.
  const ensured = await ensureMasterCategories(
    token,
    state.taxonomy.map((c) => ({ name: c.name, color: c.color })),
  );
  if (ensured.created.length > 0) notes.push(`Created categories: ${ensured.created.join(', ')}.`);

  // --- First run: mine her existing filing ---------------------------------
  if (!state.settings.bootstrapped) {
    const history = await listCategorizedHistory(token, HISTORY_DEPTH);
    const result = bootstrapSenderRules(history, state.taxonomy, state.senderRules);
    state.senderRules = result.rules;
    state.settings.bootstrapped = true;
    state.settings.generation++;
    notes.push(
      `First run: learned ${result.learned} sender(s) from ${history.length} already-categorized message(s).`,
    );
    if (result.inferred.length > 0) {
      const pairs = result.inferred
        .map((i) => `"${i.from}" -> ${categoryById(state.taxonomy, i.toCategoryId)?.name ?? i.toCategoryId}`)
        .join('; ');
      notes.push(`Matched her existing categories to ours: ${pairs}.`);
    }
  }

  const includePreview = state.settings.dataSharing === 'full';
  const mail = await listRecentInbox(token, INBOX_WINDOW, includePreview);

  // --- Corrections she made in Outlook since last run ----------------------
  const corrections = detectCorrections(mail, state.taxonomy);
  if (corrections.length > 0) {
    for (const correction of corrections) {
      state.senderRules = learnFromCorrection(state.senderRules, correction);
    }
    state.recentCorrections = [...corrections, ...state.recentCorrections].slice(0, CORRECTION_MEMORY);
    // Anything parked in Needs Review is worth reconsidering now that the
    // inputs have changed - and only now.
    state.settings.generation++;
    notes.push(
      `Learned ${corrections.length} correction(s) she made in Outlook: ${corrections
        .map((c) => `${c.sender} -> ${categoryById(state.taxonomy, c.toCategoryId)?.name ?? c.toCategoryId}`)
        .join('; ')}.`,
    );
  }

  const candidates = needsClassification(mail, state.taxonomy, state.settings.generation);

  // --- Layer 1 -------------------------------------------------------------
  const ruleDecisions: Plan['ruleDecisions'] = [];
  const pending: PendingMessage[] = [];

  for (const item of candidates) {
    const hit = classifyByRule(state.senderRules, item);
    if (hit) {
      ruleDecisions.push({
        id: item.id,
        categoryId: hit.categoryId,
        confidence: hit.confidence,
        existing: item.categories,
      });
    } else {
      pending.push(toPending(item, includePreview));
    }
  }

  if (state.settings.dataSharing === 'rules' && pending.length > 0) {
    notes.push(
      `Sharing is set to rules-only, so ${pending.length} message(s) from unknown senders were left for review and no content leaves the mailbox.`,
    );
  }

  const output: Plan = {
    generatedAt: new Date().toISOString(),
    generation: state.settings.generation,
    confidenceThreshold: state.settings.confidenceThreshold,
    dataSharing: state.settings.dataSharing,
    ruleDecisions,
    categories: selectableCategories(state.taxonomy).map((c) => ({
      id: c.id,
      name: c.name,
      description: c.description,
    })),
    examples: renderExamples(state.recentCorrections, state.taxonomy),
    // Honour the sharing setting even here: with 'rules' nothing is offered for
    // classification at all, so nothing can leave by accident.
    pending: state.settings.dataSharing === 'rules' ? [] : pending,
    notes,
  };

  await writeJson(PLAN_PATH, output);

  // State is saved now rather than after `apply`. Corrections and the bootstrap
  // are things she did, and they must not be lost because a later step failed.
  await saveState(token, state);

  console.log(`plan: ${mail.length} message(s) in window, ${candidates.length} unresolved.`);
  console.log(`plan: ${ruleDecisions.length} decided by sender rules (free).`);
  console.log(`plan: ${output.pending.length} need classification.`);
  for (const note of notes) console.log(`note: ${note}`);
  console.log(`plan written to ${PLAN_PATH}`);
}

function toPending(item: MailSummary, includePreview: boolean): PendingMessage {
  const out: PendingMessage = {
    id: item.id,
    from: item.from,
    fromName: item.fromName,
    subject: item.subject,
    received: item.received,
    hasAttachments: item.hasAttachments,
    mailingList: Boolean(item.listId),
    existing: item.categories,
  };
  if (includePreview && item.preview) out.preview = item.preview.replace(/\s+/g, ' ').slice(0, 600);
  return out;
}

/** Her own past judgements, most recent first. This is the entire learning mechanism. */
function renderExamples(
  corrections: Correction[],
  taxonomy: RoutineState['taxonomy'],
): Plan['examples'] {
  return corrections
    .filter((c) => c.toCategoryId !== NEEDS_REVIEW_ID)
    .slice(0, EXAMPLE_LIMIT)
    .map((c) => ({
      sender: c.sender,
      subject: c.subject.slice(0, 80),
      category: categoryById(taxonomy, c.toCategoryId)?.id ?? c.toCategoryId,
    }));
}

// ---------------------------------------------------------------------------
// apply
// ---------------------------------------------------------------------------

interface Verdicts {
  verdicts: {
    id: string;
    ranked: { category: string; confidence: number; reason: string }[];
  }[];
}

async function apply(): Promise<void> {
  const plan = JSON.parse(await readFile(PLAN_PATH, 'utf8')) as Plan;

  let verdicts: Verdicts = { verdicts: [] };
  try {
    verdicts = JSON.parse(await readFile(VERDICTS_PATH, 'utf8')) as Verdicts;
  } catch {
    // No verdicts file is a legitimate state: rules-only sharing, or a run where
    // every message was resolved by layer 1.
    if (plan.pending.length > 0) {
      console.warn(
        `apply: no verdicts file at ${VERDICTS_PATH}; ${plan.pending.length} message(s) will be marked Needs Review.`,
      );
    }
  }

  const { token, state } = await authenticate();
  const validIds = new Set(selectableCategories(state.taxonomy).map((c) => c.id));
  const verdictById = new Map(verdicts.verdicts.map((v) => [v.id, v]));

  const decisions: { decision: Decision; existing: string[]; sender: string }[] = [];

  // Layer 1 answers pass through the same gate, so a rule and a model verdict
  // are treated identically from here on.
  for (const rd of plan.ruleDecisions) {
    decisions.push({
      decision: {
        ...gate(rd.id, [{ categoryId: rd.categoryId, confidence: rd.confidence, reason: 'Known sender.' }], plan.confidenceThreshold),
        source: 'rule',
      },
      existing: rd.existing,
      sender: '',
    });
  }

  for (const pendingItem of plan.pending) {
    const verdict = verdictById.get(pendingItem.id);
    const ranked: Suggestion[] = (verdict?.ranked ?? [])
      // Structured output constrains shape, not vocabulary: a category id that
      // isn't in the taxonomy is dropped rather than trusted.
      .filter((r) => validIds.has(r.category))
      .map((r) => ({
        categoryId: r.category,
        confidence: Math.max(0, Math.min(1, Number(r.confidence) || 0)),
        reason: (r.reason ?? '').trim() || 'Classified by content.',
      }));

    decisions.push({
      decision:
        ranked.length > 0
          ? gate(pendingItem.id, ranked, plan.confidenceThreshold)
          : {
              messageId: pendingItem.id,
              ranked: [{ categoryId: NEEDS_REVIEW_ID, confidence: 0, reason: 'No usable category returned.' }],
              source: 'unresolved',
              gated: true,
            },
      existing: pendingItem.existing,
      sender: pendingItem.from,
    });
  }

  if (decisions.length === 0) {
    console.log('apply: nothing to write.');
    await saveState(token, state);
    return;
  }

  // --- Write ---------------------------------------------------------------
  const updates: CategoryUpdate[] = [];
  const applied = new Map<string, { categoryId: string; sender: string; gated: boolean }>();

  for (const entry of decisions) {
    const categoryId = categoryToApply(entry.decision);
    const category = categoryById(state.taxonomy, categoryId);
    if (!category) continue;

    updates.push({
      messageId: entry.decision.messageId,
      categories: mergeCategories(entry.existing, state.taxonomy, category.name),
      provenance: {
        categoryId,
        confidence: entry.decision.ranked[0]?.confidence ?? 0,
        generation: plan.generation,
      },
    });
    applied.set(entry.decision.messageId, {
      categoryId,
      sender: entry.sender,
      gated: entry.decision.gated,
    });
  }

  const result = await applyCategories(token, updates);

  // --- Learn ---------------------------------------------------------------
  // Only messages Graph confirmed, and only confident ones. An unconfirmed
  // write teaches nothing, and a gated guess is not evidence about a sender.
  let learned = 0;
  for (const id of result.succeeded) {
    const record = applied.get(id);
    if (!record || record.gated || !record.sender) continue;
    if (record.categoryId === NEEDS_REVIEW_ID) continue;
    if (!normalizeSender(record.sender)) continue;
    state.senderRules = upsertSenderRule(state.senderRules, record.sender, record.categoryId, false);
    learned++;
  }

  // --- Promote -------------------------------------------------------------
  const promotion = await promoteRules(
    token,
    state.senderRules,
    state.taxonomy,
    state.settings.promoteThreshold,
  );
  if (promotion.promotedPatterns.length > 0) {
    state.senderRules = markPromoted(state.senderRules, promotion.promotedPatterns);
  }

  state.routine = {
    ...state.routine,
    lastRunAt: new Date().toISOString(),
    runs: (state.routine?.runs ?? 0) + 1,
  };
  await saveState(token, state);

  // --- Digest --------------------------------------------------------------
  const counts = new Map<string, number>();
  for (const id of result.succeeded) {
    const record = applied.get(id);
    if (!record) continue;
    const name = categoryById(state.taxonomy, record.categoryId)?.name ?? record.categoryId;
    counts.set(name, (counts.get(name) ?? 0) + 1);
  }

  const digest = {
    labelled: result.succeeded.length,
    failed: result.failed.length,
    byCategory: Object.fromEntries([...counts].sort((a, b) => b[1] - a[1])),
    needsReview: [...applied.values()].filter((a) => a.categoryId === NEEDS_REVIEW_ID).length,
    sendersLearned: learned,
    rulesCreated: promotion.created,
    rulesUpdated: promotion.updated,
    promoted: promotion.promotedPatterns,
    quotaLimited: promotion.quotaLimited,
    notes: [...plan.notes, ...(promotion.note ? [promotion.note] : [])],
  };

  console.log(JSON.stringify(digest, null, 2));
  for (const failure of result.failed.slice(0, 5)) {
    console.warn(`apply: failed ${failure.id}: ${failure.reason}`);
  }
}

// ---------------------------------------------------------------------------

const command = process.argv[2];

const commands: Record<string, () => Promise<void>> = { login, plan, apply };
const run = commands[command ?? ''];

if (!run) {
  console.error('usage: npm run <login|plan|apply>');
  process.exit(1);
}

run().catch((err) => {
  console.error(err instanceof Error ? err.message : err);
  process.exit(1);
});
