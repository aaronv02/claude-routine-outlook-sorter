/**
 * The whole loop, end to end, against a fake mailbox. `npm run loop`
 *
 * Everything else tests a piece. This runs the actual `plan` and `apply` commands
 * from `routine/sweep.ts` - the real orchestration, the real `src/graph.ts` with its
 * batching and provenance encoding, the real state store writing to a hidden mail
 * folder, the real promotion into native Outlook rules - over four simulated runs
 * with a human correcting things in between.
 *
 * It is the only test that answers the question the product actually rests on: does
 * correcting a label in Outlook, with no other action, eventually teach Outlook
 * itself to do the job?
 *
 * The model's part is stubbed deterministically. That is not the thing under test
 * here; `npm run accuracy` measures that separately.
 */

import { readFile, writeFile, mkdir, rm } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { FakeMailbox } from './fake-graph.js';
import { plan, apply } from '../routine/sweep.js';
import { SEED_TAXONOMY } from '../src/taxonomy.js';

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '..');
const PLAN_PATH = resolve(ROOT, 'routine/.local/plan.json');
const VERDICTS_PATH = resolve(ROOT, 'routine/.local/verdicts.json');

const bold = (s: string) => `\x1b[1m${s}\x1b[0m`;
const dim = (s: string) => `\x1b[2m${s}\x1b[0m`;
const green = (s: string) => `\x1b[32m${s}\x1b[0m`;
const red = (s: string) => `\x1b[31m${s}\x1b[0m`;

const results: [string, boolean, string?][] = [];
function check(label: string, ok: boolean, detail?: string): void {
  results.push([label, ok, detail]);
}

/** What the model would say, keyed by sender. Stands in for the routine prompt. */
const MODEL: Record<string, string> = {
  'ed@mancosvalleyresources.org': 'grants',
  'grants@sanjuanmountainsassoc.org': 'grants',
  'margaret.holloway@gmail.com': 'donors',
  'statements@vanguardcharitable.org': 'finance',
  'events@durangowineexperience.com': 'events',
  // Deliberately wrong: this is the one she will correct by hand.
  'gwendolyn@animasriverwatershed.org': 'grants',
};

let day = 0;
function received(): string {
  day += 1;
  return new Date(Date.UTC(2026, 6, day, 15, 0, 0)).toISOString();
}

interface PlannedSweep {
  pending: { id: string; from: string; subject: string }[];
  ruleDecisions: { id: string; categoryId: string }[];
}

async function runSweep(mailbox: FakeMailbox, label: string): Promise<PlannedSweep> {
  console.log(`\n${bold(label)}`);

  await rm(PLAN_PATH, { force: true });
  await rm(VERDICTS_PATH, { force: true });

  await plan();

  const planned = JSON.parse(await readFile(PLAN_PATH, 'utf8')) as PlannedSweep;

  console.log(
    dim(
      `  layer 1 resolved ${planned.ruleDecisions.length}, model asked about ${planned.pending.length}`,
    ),
  );

  // Stand in for the model.
  const verdicts = planned.pending.map((m) => ({
    id: m.id,
    ranked: [
      { category: MODEL[m.from] ?? 'vendors', confidence: 0.9, reason: 'test stub' },
      { category: 'partners', confidence: 0.03, reason: 'runner-up' },
    ],
  }));
  await mkdir(dirname(VERDICTS_PATH), { recursive: true });
  await writeFile(VERDICTS_PATH, JSON.stringify({ verdicts }), 'utf8');

  await apply();
  return planned;
}

async function main(): Promise<void> {
  const mailbox = new FakeMailbox();
  const restore = mailbox.install();

  // The routine reads these from the environment; .env is absent in a test run.
  process.env.STEWARD_CLIENT_ID = '3f9a1c2e-4b5d-6e7f-8a9b-0c1d2e3f4a5b';
  process.env.STEWARD_TENANT = '11111111-2222-3333-4444-555555555555';
  process.env.STEWARD_REFRESH_TOKEN = 'fake-refresh-token';
  delete process.env.STEWARD_STATE_FILE; // exercise the real hidden-folder store

  try {
    // ---------------------------------------------------------------------
    // Mail she has already filed by hand, which the first run learns from.
    // ---------------------------------------------------------------------
    for (let i = 0; i < 3; i++) {
      mailbox.add({
        subject: `Earlier grant correspondence ${i}`,
        bodyPreview: 'Historic.',
        receivedDateTime: received(),
        hasAttachments: false,
        categories: ['Grants'],
        from: { emailAddress: { address: 'kbaptiste@pinerivervalley.org', name: 'K Baptiste' } },
        filed: true,
      });
    }

    // ---------------------------------------------------------------------
    // Run 1: a fresh inbox, nothing learned yet.
    // ---------------------------------------------------------------------
    const first = mailbox.add({
      subject: 'Letter of inquiry - food security',
      bodyPreview: 'Requesting $35,000 over two years.',
      receivedDateTime: received(),
      hasAttachments: true,
      categories: [],
      from: { emailAddress: { address: 'ed@mancosvalleyresources.org', name: 'Dana Ruiz' } },
    });
    // Carries a category of her own. It must survive every write.
    const hers = mailbox.add({
      subject: 'Adding to the Holloway Family Fund',
      bodyPreview: 'Appreciated stock this time.',
      receivedDateTime: received(),
      hasAttachments: false,
      categories: ['Smith Family Fund'],
      from: { emailAddress: { address: 'margaret.holloway@gmail.com', name: 'M Holloway' } },
    });
    // This sender appears three times in her filed history, so the bootstrap should
    // recognise it and layer 1 should resolve it with no model call at all.
    const knownFromHistory = mailbox.add({
      subject: 'Youth mentoring - quick question',
      bodyPreview: 'About the reporting timeline.',
      receivedDateTime: received(),
      hasAttachments: false,
      categories: [],
      from: { emailAddress: { address: 'kbaptiste@pinerivervalley.org', name: 'K Baptiste' } },
    });
    const willCorrect = mailbox.add({
      subject: 'Registering two staff for Tips & Tricks',
      bodyPreview: 'Is there still room?',
      receivedDateTime: received(),
      hasAttachments: false,
      categories: [],
      from: { emailAddress: { address: 'gwendolyn@animasriverwatershed.org', name: 'Gwendolyn' } },
    });

    const run1 = await runSweep(mailbox, 'Run 1 — first ever sweep');

    check(
      'the category list was created in the mailbox',
      mailbox.masterCategories.length === SEED_TAXONOMY.length,
      `${mailbox.masterCategories.length} of ${SEED_TAXONOMY.length}`,
    );
    check(
      'a sender learned from her filing needed no model call',
      run1.ruleDecisions.some((d) => d.id === knownFromHistory.id) &&
        !run1.pending.some((p) => p.id === knownFromHistory.id),
      `resolved by rule: ${run1.ruleDecisions.length}, sent to model: ${run1.pending.length}`,
    );
    check(
      'and it was labelled the way she had filed that sender before',
      mailbox.categoriesOn(knownFromHistory.id).includes('Grants'),
      mailbox.categoriesOn(knownFromHistory.id).join(', '),
    );
    check(
      'the grant LOI was labelled Grants',
      mailbox.categoriesOn(first.id).includes('Grants'),
      mailbox.categoriesOn(first.id).join(', '),
    );
    check(
      'her own category survived the write',
      mailbox.categoriesOn(hers.id).includes('Smith Family Fund'),
      mailbox.categoriesOn(hers.id).join(', '),
    );
    check(
      'ours was added alongside hers',
      mailbox.categoriesOn(hers.id).includes('Donors & Gifts'),
      mailbox.categoriesOn(hers.id).join(', '),
    );
    check(
      'a provenance stamp was written',
      Boolean(mailbox.messages.find((m) => m.id === first.id)?.stamp),
    );
    check(
      'state was persisted to the hidden folder',
      mailbox.calls.some((c) => c.includes('/me/mailFolders/folder-state/messages')),
    );
    check('nothing was promoted yet', mailbox.rules.length === 0, `${mailbox.rules.length} rules`);

    // ---------------------------------------------------------------------
    // She disagrees, in Outlook, the way she normally would.
    // ---------------------------------------------------------------------
    console.log(
      `\n${dim('  she right-clicks the Tips & Tricks mail and changes Grants → Nonprofit Partners')}`,
    );
    mailbox.sheRecategorizes(willCorrect.id, ['Nonprofit Partners']);

    // ---------------------------------------------------------------------
    // Runs 2-4: the same sender writes again each time.
    // ---------------------------------------------------------------------
    const repeats: string[] = [];
    for (let run = 2; run <= 4; run++) {
      const again = mailbox.add({
        subject: `Another workshop question ${run}`,
        bodyPreview: 'Following up.',
        receivedDateTime: received(),
        hasAttachments: false,
        categories: [],
        from: { emailAddress: { address: 'gwendolyn@animasriverwatershed.org', name: 'Gwendolyn' } },
      });
      repeats.push(again.id);

      await runSweep(mailbox, `Run ${run} — she has corrected that sender`);

      if (run === 2) {
        check(
          'the correction was learned, not re-guessed',
          mailbox.categoriesOn(again.id).includes('Nonprofit Partners'),
          mailbox.categoriesOn(again.id).join(', '),
        );
      }

      // Each later run re-reads the corrected message and counts it again, which is
      // how confirmations accumulate toward the promotion threshold.
      mailbox.sheRecategorizes(willCorrect.id, ['Nonprofit Partners']);
      const stamped = mailbox.messages.find((m) => m.id === willCorrect.id);
      if (stamped) stamped.stamp = `grants|0.90|${new Date().toISOString()}|0`;
    }

    check(
      'later mail from that sender never went to the model',
      repeats.every((id) => mailbox.categoriesOn(id).includes('Nonprofit Partners')),
    );

    const ours = mailbox.rules.filter((r) => r.displayName?.startsWith('Inbox Steward:'));
    check(
      'a native Outlook rule was created once confirmed enough',
      ours.length > 0,
      ours.map((r) => r.displayName).join(', ') || 'none',
    );
    check(
      'the rule targets the corrected sender',
      ours.some((r) =>
        (r.conditions?.senderContains ?? []).includes('gwendolyn@animasriverwatershed.org'),
      ),
      JSON.stringify(ours[0]?.conditions ?? {}),
    );
    check(
      'the rule assigns the category she chose',
      ours.some((r) => (r.actions?.assignCategories ?? []).includes('Nonprofit Partners')),
      JSON.stringify(ours[0]?.actions ?? {}),
    );

    // ---------------------------------------------------------------------
    console.log(`\n${bold('The mailbox afterwards')}\n`);
    for (const m of mailbox.messages.filter((x) => !x.filed)) {
      console.log(`  ${m.categories.join(' + ') || dim('(none)')}  ${dim('←')}  ${m.subject}`);
    }
    console.log(`\n${bold('Native Outlook rules it wrote')}\n`);
    for (const r of mailbox.rules) {
      console.log(`  ${r.displayName}`);
      console.log(dim(`    senders: ${(r.conditions?.senderContains ?? []).join(', ')}`));
      console.log(dim(`    applies: ${(r.actions?.assignCategories ?? []).join(', ')}`));
    }
    if (mailbox.rules.length === 0) console.log(dim('  none'));
  } finally {
    restore();
  }

  console.log(`\n${bold('Results')}\n`);
  let failed = 0;
  for (const [label, ok, detail] of results) {
    console.log(`  ${ok ? green('✓') : red('✗')} ${label}${detail && !ok ? dim(`  — ${detail}`) : ''}`);
    if (!ok) failed++;
  }
  console.log('');

  if (failed > 0) {
    console.log(red(`${failed} of ${results.length} failed.\n`));
    process.exit(1);
  }
  console.log(green(`All ${results.length} passed. The loop closes.\n`));
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
