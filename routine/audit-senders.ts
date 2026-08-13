/**
 * Audits the automated-sender filter against the real mailbox. `npm run audit`
 *
 * The filter decides who counts as a person waiting on a reply, and it is the
 * single most consequential guess this tool makes: every bulk sender it misses
 * becomes a line in the weekly "waiting on you" list, and a list padded with
 * newsletters stops being read within a fortnight. Meanwhile every real person it
 * wrongly filters disappears from that list silently, which is worse.
 *
 * The defaults were tuned against a nonprofit-sector inbox, so they anchor on the
 * old shape - newsletter@, updates@, news@. Auditing a different real inbox showed
 * them catching only a third of the bulk mail actually present, because modern
 * marketing puts the brand in the local part and the sending platform in a
 * subdomain: news@em.oakley.com, ebay@reply.ebay.com, alerts@notify.wellsfargo.com.
 *
 * So this exists to be run against HER mailbox, on her machine, by whoever set it
 * up. It prints addresses and counts - no subjects, no message content - and sends
 * nothing anywhere. Ten seconds, and it turns a guess about her mail into an
 * observation.
 */

import { loadEnv } from './env.js';
import { assertConfigured, redeemRefreshToken } from './auth.js';
import { fetchInboxSince, fetchMailbox } from '../src/digest/fetch.js';
import { DEFAULT_IGNORED_SENDERS, isAutomatedSender } from '../src/digest/senders.js';
import { addressedToAny, fromAddress, fromName } from '../src/digest/types.js';

const bold = (s: string) => `\x1b[1m${s}\x1b[0m`;
const dim = (s: string) => `\x1b[2m${s}\x1b[0m`;
const green = (s: string) => `\x1b[32m${s}\x1b[0m`;
const yellow = (s: string) => `\x1b[33m${s}\x1b[0m`;

const DEFAULT_DAYS = 60;

interface SenderStat {
  address: string;
  name: string;
  count: number;
  /** How many were addressed to her directly, i.e. could reach the waiting list. */
  toHer: number;
  automated: boolean;
}

async function main(): Promise<void> {
  await loadEnv();
  assertConfigured();

  const days = Number(process.argv[2] ?? '') || DEFAULT_DAYS;

  const refreshToken = process.env.STEWARD_REFRESH_TOKEN;
  if (!refreshToken) {
    throw new Error('No refresh token available. Run `npm run setup` first.');
  }

  const tokens = await redeemRefreshToken(refreshToken);
  const mailbox = await fetchMailbox(tokens.accessToken);

  const addresses = [
    mailbox.address,
    ...(process.env.STEWARD_ALSO_ADDRESSED_AS ?? '')
      .split(',')
      .map((a) => a.trim().toLowerCase())
      .filter(Boolean),
  ].filter(Boolean);

  const since = new Date(Date.now() - days * 86_400_000);
  const inbox = await fetchInboxSince(tokens.accessToken, since);

  const patterns = [
    ...DEFAULT_IGNORED_SENDERS,
    ...(process.env.STEWARD_IGNORED_SENDERS ?? '')
      .split(',')
      .map((p) => p.trim())
      .filter(Boolean),
  ];

  const stats = new Map<string, SenderStat>();
  for (const message of inbox) {
    const address = fromAddress(message);
    if (address === '') continue;

    let stat = stats.get(address);
    if (!stat) {
      stat = {
        address,
        name: fromName(message),
        count: 0,
        toHer: 0,
        automated: isAutomatedSender(address, patterns),
      };
      stats.set(address, stat);
    }
    stat.count++;
    if (addressedToAny(message, addresses)) stat.toHer++;
  }

  const all = [...stats.values()];
  const automated = all.filter((s) => s.automated).sort((a, b) => b.count - a.count);
  const people = all.filter((s) => !s.automated).sort((a, b) => b.count - a.count);

  console.log(`\n${bold('Sender filter audit')}`);
  console.log(
    dim(
      `${mailbox.address} — ${inbox.length} message(s) over ${days} days, ${all.length} distinct senders\n`,
    ),
  );

  console.log(`${green('Treated as bulk')} — never shown as waiting on a reply  ${dim(`(${automated.length})`)}`);
  for (const s of automated.slice(0, 40)) {
    console.log(`  ${String(s.count).padStart(3)}  ${s.address}`);
  }
  if (automated.length > 40) console.log(dim(`  ... and ${automated.length - 40} more`));

  console.log(
    `\n${yellow('Treated as a person')} — can appear as waiting on a reply  ${dim(`(${people.length})`)}`,
  );
  for (const s of people.slice(0, 60)) {
    // The ones addressed to her are the ones that actually reach the waiting list;
    // the rest are already excluded by the To-line rule.
    const reach = s.toHer > 0 ? '' : dim('  (never on her To line, so already excluded)');
    console.log(`  ${String(s.count).padStart(3)}  ${s.address}${reach}`);
  }
  if (people.length > 60) console.log(dim(`  ... and ${people.length - 60} more`));

  // The actionable subset: high volume, addressed to her, and looks like a machine
  // by its own naming. Anything a human would obviously recognise as bulk.
  const suspicious = people
    .filter((s) => s.toHer > 0 && s.count >= 3)
    .filter((s) => /^(no|do)|reply|mail|info|news|alert|notif|offer|promo|deal|market|team|support|hello|contact|billing|account|service|member|update/i.test(s.address.split('@')[0] as string) || (s.address.split('@')[1] ?? '').split('.').length >= 3);

  console.log(`\n${bold('Worth reviewing')}`);
  if (suspicious.length === 0) {
    console.log(dim('  Nothing obvious. The defaults look right for this mailbox.\n'));
  } else {
    console.log(
      dim(
        '  High-volume senders reaching her To line whose names look machine-generated.\n  Anything here that is really bulk should be added to STEWARD_IGNORED_SENDERS.\n',
      ),
    );
    for (const s of suspicious) console.log(`  ${String(s.count).padStart(3)}  ${s.address}`);

    // Full addresses, not domain patterns. A domain pattern is tempting because it
    // covers every address a sender rotates through, but it also silences everyone
    // else at that domain - and this list will contain false positives by design.
    // "accounting@herCPA.com" looks machine-generated and is a person; excluding
    // the whole domain would hide her accountant from the waiting list without
    // saying so.
    console.log(`\n${bold('To exclude the ones that really are bulk, add to .env:')}\n`);
    console.log(`  STEWARD_IGNORED_SENDERS=${suspicious.map((s) => s.address).join(',')}\n`);
    console.log(
      dim(
        '  Prune that line first - it lists candidates, not conclusions. Some of these\n  will be people whose address merely looks automated. Removing a person from\n  the waiting list happens silently, so it is the error worth avoiding.\n\n  To cover a sender that rotates addresses, replace it with @their-domain.com\n  yourself, once you are sure nobody there writes to her personally.\n',
      ),
    );
  }
}

main().catch((err) => {
  console.error(err instanceof Error ? err.message : String(err));
  process.exit(1);
});
