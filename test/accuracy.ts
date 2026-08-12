/**
 * How well does it actually sort? `npm run accuracy <verdicts.json>`
 *
 * Scores a set of model verdicts against the labelled fixtures in
 * `test/fixtures/emails.json`, through the real confidence gate rather than a
 * simplified copy of it. So this measures the thing that ships: what would end up
 * on the messages, including what would correctly land in ⚠ Needs Review.
 *
 * Two numbers matter and they measure different things:
 *
 *   top-1  - the label that would actually be written was right.
 *   top-3  - the right label was among the three offered.
 *
 * top-3 is not padding. The taskpane shows three ranked chips and the MCP server
 * proposes before applying, so in both of those a top-3 hit costs one click. Only
 * the unattended schedule is top-1-or-nothing.
 *
 * Expect roughly 80% top-1 and 90% top-3. Personal-email foldering is genuinely
 * harder than it sounds: published benchmarks land around 60-80% for a single
 * forced guess. The 95%+ figures quoted elsewhere come from spam/ham and
 * news-topic corpora, not personal folder taxonomies.
 *
 * A third number is arguably the most important: how often it is confidently
 * WRONG. A wrong label teaches her the tool cannot be trusted; an honest
 * Needs Review costs her a moment. The gate exists to trade the first for the
 * second, and this reports whether it does.
 */

import { readFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { gate, categoryToApply } from '../src/classify/confidence.js';
import { NEEDS_REVIEW_ID, SEED_TAXONOMY, categoryById } from '../src/taxonomy.js';
import { DEFAULT_SETTINGS } from '../src/taxonomy.js';

const HERE = dirname(fileURLToPath(import.meta.url));

interface Fixture {
  from: string;
  fromName: string;
  subject: string;
  preview: string;
  expected: string;
}

interface Verdicts {
  verdicts: { ref: number; ranked: { category: string; confidence: number; reason?: string }[] }[];
}

const bold = (s: string) => `\x1b[1m${s}\x1b[0m`;
const dim = (s: string) => `\x1b[2m${s}\x1b[0m`;
const green = (s: string) => `\x1b[32m${s}\x1b[0m`;
const red = (s: string) => `\x1b[31m${s}\x1b[0m`;
const yellow = (s: string) => `\x1b[33m${s}\x1b[0m`;

async function main(): Promise<void> {
  const verdictPath = process.argv[2];
  if (!verdictPath) {
    console.error('usage: tsx test/accuracy.ts <verdicts.json>');
    process.exit(1);
  }

  const fixtures = JSON.parse(
    await readFile(resolve(HERE, 'fixtures/emails.json'), 'utf8'),
  ) as Fixture[];
  const { verdicts } = JSON.parse(await readFile(verdictPath, 'utf8')) as Verdicts;

  const byRef = new Map(verdicts.map((v) => [v.ref, v]));
  const threshold = DEFAULT_SETTINGS.confidenceThreshold;
  const validIds = new Set(SEED_TAXONOMY.map((c) => c.id));

  let top1 = 0;
  let top3 = 0;
  let confidentlyWrong = 0;
  let gatedButRight = 0;
  let gatedAndWrong = 0;
  let missing = 0;

  const rows: string[] = [];
  const confusion = new Map<string, number>();

  fixtures.forEach((fixture, ref) => {
    const verdict = byRef.get(ref);
    if (!verdict) {
      missing++;
      return;
    }

    // Same validation the real pipeline applies: an id outside the taxonomy is
    // dropped rather than trusted.
    const ranked = verdict.ranked
      .filter((r) => validIds.has(r.category) && r.category !== NEEDS_REVIEW_ID)
      .map((r) => ({
        categoryId: r.category,
        confidence: Math.max(0, Math.min(1, Number(r.confidence) || 0)),
        reason: r.reason ?? '',
      }));

    const decision = gate(String(ref), ranked, threshold);
    const written = categoryToApply(decision);

    const first = decision.ranked[0]?.categoryId;
    const inTop3 = decision.ranked.some((r) => r.categoryId === fixture.expected);
    const rightFirst = first === fixture.expected;

    if (rightFirst) top1++;
    if (inTop3) top3++;

    if (written === NEEDS_REVIEW_ID) {
      if (inTop3) gatedButRight++;
      else gatedAndWrong++;
    } else if (written !== fixture.expected) {
      confidentlyWrong++;
      confusion.set(
        `${fixture.expected} → ${written}`,
        (confusion.get(`${fixture.expected} → ${written}`) ?? 0) + 1,
      );
    }

    const mark =
      written === fixture.expected
        ? green('✓')
        : written === NEEDS_REVIEW_ID
          ? yellow('?')
          : red('✗');

    if (written !== fixture.expected) {
      const name = (id: string) => categoryById(SEED_TAXONOMY, id)?.name ?? id;
      rows.push(
        `  ${mark} ${dim(`[${ref}]`)} ${fixture.subject.slice(0, 58)}\n` +
          `      wanted ${bold(name(fixture.expected))}, wrote ${bold(name(written))}` +
          (inTop3 ? dim('  (right answer was offered)') : ''),
      );
    }
  });

  const scored = fixtures.length - missing;
  const pct = (n: number) => `${((n / scored) * 100).toFixed(0)}%`;

  console.log(`\n${bold('Sorting accuracy')}  ${dim(`${scored} messages, threshold ${threshold}`)}\n`);
  console.log(`  top-1  ${bold(pct(top1))}  ${dim(`${top1}/${scored} — the label it would write`)}`);
  console.log(`  top-3  ${bold(pct(top3))}  ${dim(`${top3}/${scored} — right answer offered`)}`);
  console.log('');
  console.log(`  ${green('written and correct')}      ${scored - confidentlyWrong - gatedButRight - gatedAndWrong}`);
  console.log(`  ${yellow('sent to Needs Review')}     ${gatedButRight + gatedAndWrong}  ${dim(`(${gatedButRight} of which it actually knew)`)}`);
  console.log(`  ${red('confidently wrong')}       ${confidentlyWrong}  ${dim(pct(confidentlyWrong))}`);

  if (missing > 0) console.log(`\n  ${yellow(`${missing} message(s) had no verdict`)}`);

  if (rows.length > 0) {
    console.log(`\n${bold('Everything it did not write correctly')}\n`);
    console.log(rows.join('\n'));
  }

  if (confusion.size > 0) {
    console.log(`\n${bold('Confident mistakes, by pair')}  ${dim('sharpen these descriptions first')}\n`);
    for (const [pair, count] of [...confusion].sort((a, b) => b[1] - a[1])) {
      console.log(`  ${count}x  ${pair}`);
    }
  }

  console.log('');
  if (top1 / scored < 0.8) {
    console.log(
      yellow('top-1 is below the 80% target. Sharpen the category descriptions in src/taxonomy.ts.\n'),
    );
  }
  if (confidentlyWrong / scored > 0.1) {
    console.log(
      yellow('More than 10% confidently wrong. Consider raising confidenceThreshold.\n'),
    );
  }
}

main().catch((err) => {
  console.error(err instanceof Error ? err.message : String(err));
  process.exit(1);
});
