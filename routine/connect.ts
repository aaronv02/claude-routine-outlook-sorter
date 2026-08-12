/**
 * `npm run connect` / `npm run disconnect`.
 *
 * Separate from `setup` so the MCP side can be added, repaired, or removed later
 * without redoing the sign-in.
 */

import { connectToDesktop, disconnectFromDesktop, configPath } from './desktop.js';

const bold = (s: string) => `\x1b[1m${s}\x1b[0m`;
const dim = (s: string) => `\x1b[2m${s}\x1b[0m`;
const green = (s: string) => `\x1b[32m${s}\x1b[0m`;

async function main(): Promise<void> {
  if (process.argv[2] === 'disconnect') {
    const { path, removed } = await disconnectFromDesktop();
    console.log(
      removed
        ? `${green('✓')} Removed from ${path}`
        : `Nothing to remove — no entry found in ${path}`,
    );
    console.log(`\n${bold('Fully quit Claude Desktop and reopen it')} for this to take effect.\n`);
    return;
  }

  const { path, replaced, preserved } = await connectToDesktop();

  console.log(`${green('✓')} ${replaced ? 'Updated' : 'Added'} the outlook-sorter entry in:`);
  console.log(`  ${path}`);
  if (preserved.length > 0) {
    console.log(dim(`  (left your other servers alone: ${preserved.join(', ')})`));
  }

  console.log(`
${bold('Now fully quit Claude Desktop and reopen it.')}

Closing the window is not enough. On Windows, right-click the Claude icon in the
system tray near the clock, choose ${bold('Quit')}, then open Claude again.

To check it worked, ask Claude: ${bold('"who is waiting on me?"')}
It should answer with real messages rather than saying it cannot see your email.
`);
}

main().catch((err) => {
  console.error(`\n${err instanceof Error ? err.message : String(err)}`);
  console.error(`\nConfig path was: ${configPath()}\n`);
  process.exit(1);
});
