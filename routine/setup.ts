/**
 * Guided setup. `npm run setup`.
 *
 * Written to be run once, with the mailbox owner sitting next to whoever is
 * doing it. Everything that can be automated is; the one step that cannot - a
 * human clicking through an app registration in the Entra portal - is reduced to
 * four numbered clicks and then verified rather than trusted.
 *
 * The design rule here is that no step is left to be discovered later. A missing
 * permission, a registration created as the wrong platform, a tenant that blocks
 * app creation: each of those is caught in this script, at the moment it can
 * still be fixed, and named in words that say what to do about it. The
 * alternative is an opaque Graph 403 three days into a schedule nobody is
 * watching.
 */

import { createInterface } from 'node:readline/promises';
import { spawn } from 'node:child_process';
import { stdin, stdout, platform } from 'node:process';

import { loadEnv, saveEnv, ENV_PATH } from './env.js';
import {
  REQUIRED_SCOPES,
  missingScopes,
  pollDeviceCode,
  startDeviceCode,
  type TokenSet,
} from './auth.js';
import { loadState, saveState } from './store.js';
import { bootstrapSenderRules } from '../src/classify/rules.js';
import { ensureMasterCategories, listCategorizedHistory, listRecentInbox } from '../src/graph.js';
import { categoryById } from '../src/taxonomy.js';

const bold = (s: string) => `\x1b[1m${s}\x1b[0m`;
const dim = (s: string) => `\x1b[2m${s}\x1b[0m`;
const green = (s: string) => `\x1b[32m${s}\x1b[0m`;
const red = (s: string) => `\x1b[31m${s}\x1b[0m`;
const yellow = (s: string) => `\x1b[33m${s}\x1b[0m`;

const rl = createInterface({ input: stdin, output: stdout });

/**
 * Set once setup has printed its closing instructions.
 *
 * Without this, closing stdin midway - Ctrl-D, or a non-interactive shell -
 * leaves `rl.question()` pending forever, the event loop empties, and Node exits
 * 0 having said nothing. Someone half-configured would be told they were fine.
 */
let finished = false;

rl.on('close', () => {
  if (finished) return;
  console.log(dim('\n\nSetup stopped before it finished. Nothing is scheduled yet.'));
  console.log(dim('Run `npm run setup` again to pick it up from the start.\n'));
  process.exit(130);
});

function heading(n: number, of: number, title: string): void {
  console.log(`\n${bold(`Step ${n} of ${of} — ${title}`)}`);
  console.log(dim('─'.repeat(60)));
}

/** Best-effort conveniences. A failure here is never worth mentioning. */
async function copyToClipboard(text: string): Promise<boolean> {
  const cmd =
    platform === 'darwin' ? 'pbcopy' : platform === 'win32' ? 'clip' : 'xclip -selection clipboard';
  return new Promise((done) => {
    try {
      const child = spawn(cmd, { shell: true, stdio: ['pipe', 'ignore', 'ignore'] });
      child.on('error', () => done(false));
      child.on('close', (code) => done(code === 0));
      child.stdin.end(text);
    } catch {
      done(false);
    }
  });
}

function openInBrowser(url: string): void {
  const cmd = platform === 'darwin' ? 'open' : platform === 'win32' ? 'start ""' : 'xdg-open';
  try {
    spawn(`${cmd} "${url}"`, { shell: true, stdio: 'ignore', detached: true }).unref();
  } catch {
    // The URL is printed regardless.
  }
}

const GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000';

/**
 * Delegated Microsoft Graph permission ids, for the manifest paste.
 *
 * These are the well-known, tenant-independent ids for the four scopes this tool
 * needs. They are verified after sign-in against what Entra actually granted, so
 * a wrong id here fails loudly in step 3 rather than silently much later.
 */
const MANIFEST_PERMISSIONS = [
  { name: 'Mail.ReadWrite', id: '024d486e-b451-40bb-833d-3e66d98c5c73' },
  { name: 'MailboxSettings.ReadWrite', id: '818c620a-27a9-40bd-a6a5-d96f7d610b4b' },
  { name: 'User.Read', id: 'e1fe6dd8-ba31-4d61-89e7-88639da4683d' },
  { name: 'offline_access', id: '7427e0e9-2fba-42fe-b0c0-848c9e6a8182' },
];

/**
 * The same configuration as the click path, as manifest JSON.
 *
 * Offered as a shortcut rather than the main route: pasting it means editing JSON
 * in the portal's manifest editor, and a botched merge produces an app that is
 * misconfigured in ways the error messages don't obviously point at. The buttons
 * are slower and much harder to get wrong.
 *
 * Compact on purpose - it is something to paste, not to read.
 */
function manifestSnippet(): string {
  const scopes = MANIFEST_PERMISSIONS.map((p) => `{ "id": "${p.id}", "type": "Scope" }`).join(
    ',\n      ',
  );
  return `"isFallbackPublicClient": true,
"publicClient": { "redirectUris": ["https://login.microsoftonline.com/common/oauth2/nativeclient"] },
"requiredResourceAccess": [
  {
    "resourceAppId": "${GRAPH_APP_ID}",
    "resourceAccess": [
      ${scopes}
    ]
  }
]`;
}

const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

async function main(): Promise<void> {
  await loadEnv();

  console.log(`\n${bold('Outlook Sorter — setup')}`);
  console.log('Takes about ten minutes. Have the mailbox owner nearby: they sign in at step 3.\n');

  const major = Number(process.versions.node.split('.')[0]);
  if (major < 20) {
    console.log(red(`Node ${process.versions.node} is too old. Install Node 20 or newer first.`));
    process.exit(1);
  }

  // -------------------------------------------------------------------------
  heading(1, 5, 'Register an app so this can reach the mailbox');

  console.log(`This is the only step that isn't automated. Six clicks, about three minutes.
Everything after it is checked for you, so mistakes here get caught, not inherited.

  ${bold('1.')} Open ${bold('https://aka.ms/appregistrations')} → ${bold('New registration')}.
     Name it "Outlook Sorter", leave every other option alone, press ${bold('Register')}.

  ${bold('2.')} Left menu → ${bold('Authentication')} → ${bold('Add a platform')} →
     ${bold('Mobile and desktop applications')} → tick the ${bold('nativeclient')} box → ${bold('Configure')}.

  ${bold('3.')} Still on that page, scroll to ${bold('Allow public client flows')} and set it to
     ${bold('Yes')}. Press ${bold('Save')}. ${dim('(Sign-in fails without this. It is the one people miss.)')}

  ${bold('4.')} Left menu → ${bold('API permissions')} → ${bold('Add a permission')} →
     ${bold('Microsoft Graph')} → ${bold('Delegated permissions')}. Search for and tick each of:

       ${bold(REQUIRED_SCOPES.join('   '))}   ${bold('offline_access')}

     Then ${bold('Add permissions')}. None of them need an administrator's approval.
`);

  const shouldOpen = await rl.question('Open the portal in your browser now? [Y/n] ');
  if (!shouldOpen.trim().toLowerCase().startsWith('n')) {
    openInBrowser('https://aka.ms/appregistrations');
  }

  console.log(`
${dim('Comfortable with JSON and want to skip clicks 2-4? Open the app\'s Manifest')}
${dim('page instead and merge these keys into it. Say y to copy it to your clipboard.')}`);
  const wantManifest = await rl.question('Copy the manifest shortcut? [y/N] ');
  if (wantManifest.trim().toLowerCase().startsWith('y')) {
    const snippet = manifestSnippet();
    const copied = await copyToClipboard(snippet);
    console.log(`\n${dim(snippet.split('\n').map((l) => `  ${l}`).join('\n'))}`);
    console.log(copied ? `\n  ${green('✓ copied')}` : `\n  ${dim('(copy it from above)')}`);
  }

  console.log(`
${yellow('If you cannot create an app registration at all')}, this tenant blocks it for
non-administrators. Whoever administers the foundation's Microsoft 365 account has
to do this one step. Nothing else here works until the registration exists.
`);

  // -------------------------------------------------------------------------
  heading(2, 5, 'Tell this tool about that app');

  console.log(`On the app's ${bold('Overview')} page, copy ${bold('Application (client) ID')}.\n`);

  let clientId = '';
  for (;;) {
    clientId = (await rl.question('Application (client) ID: ')).trim();
    if (UUID.test(clientId)) break;
    console.log(red("That doesn't look like an ID. It should look like 3f9a1c2e-4b5d-... (36 characters)."));
  }

  // Asked for rather than defaulted. Signing in against "organizations" fails
  // with AADSTS50059 on some tenants, and that error names nothing useful - one
  // extra copy-paste from the same screen removes the whole failure mode.
  console.log(`\n${dim('Directory (tenant) ID is on that same Overview page, just below it.')}`);

  let tenant = '';
  for (;;) {
    tenant = (await rl.question('Directory (tenant) ID: ')).trim();
    if (UUID.test(tenant)) break;
    console.log(red('That should also be a 36-character ID, from the same page.'));
  }

  process.env.STEWARD_CLIENT_ID = clientId;
  process.env.STEWARD_TENANT = tenant;

  // -------------------------------------------------------------------------
  heading(3, 5, 'The mailbox owner signs in');

  console.log('This is the part they do. It creates the key this tool uses from then on.\n');

  let code;
  try {
    code = await startDeviceCode();
  } catch (err) {
    console.log(red('\nCould not start sign-in.'));
    // The message from auth.ts already names the likely cause and the page to
    // fix it on, so print it rather than guessing over the top of it.
    console.log(`\n${err instanceof Error ? err.message.split('\n')[0] : String(err)}\n`);
    console.log(`Fix that, then run ${bold('npm run setup')} again.\n`);
    console.log(dim(err instanceof Error ? err.message.split('\n').slice(2).join('\n') : ''));
    process.exit(1);
  }

  const codeCopied = await copyToClipboard(code.userCode);
  console.log(`  Go to:  ${bold(code.verificationUri)}`);
  console.log(`  Enter:  ${bold(code.userCode)}${codeCopied ? green('   ✓ copied') : ''}\n`);
  console.log(dim('  Sign in as the person whose mail is being sorted - not as yourself,'));
  console.log(dim('  unless it is your own mailbox. Whoever signs in is whose mail this reads.\n'));

  openInBrowser(code.verificationUri);
  console.log('  Waiting...');

  let tokens: TokenSet;
  try {
    tokens = await pollDeviceCode(code);
  } catch (err) {
    console.log(red(`\nSign-in did not complete: ${err instanceof Error ? err.message : err}`));
    console.log('\nRun `npm run setup` again to retry.\n');
    process.exit(1);
  }

  // Sign-in succeeding does not mean the permissions were added. Catch it here,
  // where the fix is one portal page away, rather than as a 403 mid-schedule.
  const missing = missingScopes(tokens.grantedScopes);
  if (missing.length > 0) {
    console.log(red(`\nSigned in, but these permissions were not granted: ${missing.join(', ')}`));
    console.log(
      `\nGo back to the app registration → ${bold('API permissions')} → ${bold('Add a permission')} →\n${bold('Microsoft Graph')} → ${bold('Delegated permissions')}, add the ones above, then run\n${bold('npm run setup')} again.\n`,
    );
    process.exit(1);
  }

  console.log(green('\n  ✓ Signed in, and all four permissions are present.'));

  await saveEnv({
    STEWARD_CLIENT_ID: clientId,
    STEWARD_TENANT: tenant,
    STEWARD_REFRESH_TOKEN: tokens.refreshToken,
  });
  console.log(`  ${green('✓')} Saved to ${ENV_PATH} ${dim('(mode 600, gitignored)')}`);

  // -------------------------------------------------------------------------
  heading(4, 5, 'Prepare the mailbox');

  const state = await loadState(tokens.accessToken);

  const whose = await whoami(tokens.accessToken);
  if (whose) console.log(`Signed in as ${bold(whose)}.\n`);

  console.log(
    `The ${state.taxonomy.length} labels need to exist in the mailbox before they can be applied.\nThey are ordinary Outlook categories - coloured tags. Nothing is moved or deleted.\n`,
  );
  const makeCats = await rl.question('Create them now? [Y/n] ');
  if (makeCats.trim().toLowerCase().startsWith('n')) {
    console.log(yellow('\nSkipped. Nothing will get labelled until they exist.\n'));
  } else {
    const ensured = await ensureMasterCategories(
      tokens.accessToken,
      state.taxonomy.map((c) => ({ name: c.name, color: c.color })),
    );
    console.log(
      `  ${green('✓')} ${ensured.created.length} created, ${ensured.existing.length} already there.`,
    );
  }

  // Reading how she already files mail is the single most reassuring number in
  // this whole process, so it happens here rather than silently on the first
  // scheduled run.
  if (!state.settings.bootstrapped) {
    console.log('\nReading how this mailbox is already filed, to learn its regular senders...');
    const history = await listCategorizedHistory(tokens.accessToken, 600);
    const result = bootstrapSenderRules(history, state.taxonomy, state.senderRules);
    state.senderRules = result.rules;
    state.settings.bootstrapped = true;
    state.settings.generation++;

    console.log(
      `  ${green('✓')} Learned ${bold(String(result.learned))} sender(s) from ${history.length} already-categorized message(s).`,
    );
    if (result.inferred.length > 0) {
      console.log(dim('\n  Existing categories matched to ours:'));
      for (const i of result.inferred) {
        console.log(
          dim(`    "${i.from}" → ${categoryById(state.taxonomy, i.toCategoryId)?.name ?? i.toCategoryId}`),
        );
      }
      console.log(dim('  Anything not listed was left alone.'));
    }
    if (result.learned === 0) {
      console.log(
        yellow(
          '\n  Nothing was learned - this mailbox has little or no existing categorization.\n  That is fine; it just means the first few days lean on the model more.',
        ),
      );
    }
    await saveState(tokens.accessToken, state);
  }

  const inbox = await listRecentInbox(tokens.accessToken, 25, false);
  console.log(`\n  ${green('✓')} Can read the inbox (${inbox.length} recent message(s) visible).`);

  // -------------------------------------------------------------------------
  heading(5, 5, 'One real sweep, so you can see it work');

  console.log(
    `This runs ${bold('npm run plan')}: it reads the inbox and works out what it would\nlabel. It writes nothing to the mailbox.\n`,
  );
  const doPlan = await rl.question('Run it? [Y/n] ');
  rl.close();

  if (!doPlan.trim().toLowerCase().startsWith('n')) {
    console.log('');
    await runPlan();
  }

  finished = true;
  console.log(`
${bold('Done.')} What's left:

  ${bold('1.')} Look at ${bold('routine/.local/plan.json')}. The "pending" list is the mail it wants
     help with. If that looks like real mail, everything works.

  ${bold('2.')} Schedule it. Create a scheduled Claude task with:

       working directory   this folder
       prompt              the text in routine/PROMPT.md
       environment         STEWARD_CLIENT_ID and STEWARD_REFRESH_TOKEN
                           ${dim('(both are in .env - copy them into the runner as secrets)')}

     Hourly during the working day is the right cadence.

  ${bold('3.')} Tell her the only thing she needs to know: ${bold('if a label is wrong, change it in')}
     ${bold('Outlook the way she normally would.')} It learns from that. There is nothing
     to open and no button to press.

${dim('The sign-in expires after 90 days of no runs, or immediately if her password')}
${dim('changes. When that happens the routine reports "invalid_grant" - rerun npm run setup.')}
`);
}

async function whoami(token: string): Promise<string | null> {
  try {
    const res = await fetch('https://graph.microsoft.com/v1.0/me?$select=displayName,mail', {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok) return null;
    const body = (await res.json()) as { displayName?: string; mail?: string };
    return [body.displayName, body.mail].filter(Boolean).join(' — ') || null;
  } catch {
    return null;
  }
}

/**
 * Run the real command rather than reimplementing it, so this verifies the thing
 * the schedule will actually invoke.
 */
function runPlan(): Promise<void> {
  return new Promise((done) => {
    const child = spawn('npm', ['run', '--silent', 'plan'], { stdio: 'inherit', shell: false });
    child.on('close', () => done());
    child.on('error', () => done());
  });
}

main().catch((err) => {
  rl.close();
  const message = err instanceof Error ? err.message : String(err);
  // Ctrl-C or a piped stdin running out. Not an error worth a stack trace.
  if (/readline was closed/i.test(message)) {
    console.log(dim('\nSetup cancelled. Run `npm run setup` again when ready.\n'));
    process.exit(130);
  }
  console.error(red(`\n${message}`));
  process.exit(1);
});
