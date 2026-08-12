import { readFile, writeFile, chmod } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';

/**
 * Configuration from a `.env` file, so nobody has to remember to `export`
 * anything before running a command.
 *
 * Written by `npm run setup` and read by every other command. Deliberately not a
 * dependency: this is fifteen lines of parsing, and the whole project otherwise
 * installs nothing at runtime.
 *
 * A real environment variable always wins over the file. That matters for the
 * scheduled run, where the refresh token should come from the runner's secret
 * store rather than a file on disk.
 */

const HERE = dirname(new URL(import.meta.url).pathname);
export const ENV_PATH = resolve(HERE, '../.env');

export async function loadEnv(): Promise<void> {
  let text: string;
  try {
    text = await readFile(ENV_PATH, 'utf8');
  } catch {
    return;
  }

  for (const line of text.split('\n')) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;

    const eq = trimmed.indexOf('=');
    if (eq < 1) continue;

    const key = trimmed.slice(0, eq).trim();
    // Quotes are stripped because a refresh token pasted from a terminal often
    // arrives wearing them.
    const value = trimmed.slice(eq + 1).trim().replace(/^["']|["']$/g, '');

    if (process.env[key] === undefined) process.env[key] = value;
  }
}

/**
 * Rewrite `.env`, preserving anything already in it that we don't manage.
 *
 * Mode 600 because this file holds a 90-day credential to someone's mailbox.
 */
export async function saveEnv(values: Record<string, string>): Promise<void> {
  const existing = new Map<string, string>();
  try {
    const text = await readFile(ENV_PATH, 'utf8');
    for (const line of text.split('\n')) {
      const eq = line.indexOf('=');
      if (eq > 0 && !line.trim().startsWith('#')) {
        existing.set(line.slice(0, eq).trim(), line.slice(eq + 1).trim());
      }
    }
  } catch {
    // No file yet. Nothing to preserve.
  }

  for (const [key, value] of Object.entries(values)) existing.set(key, value);

  const body = [
    '# Written by `npm run setup`. Holds a credential to the mailbox - never commit it.',
    ...[...existing].map(([k, v]) => `${k}=${v}`),
    '',
  ].join('\n');

  await writeFile(ENV_PATH, body, 'utf8');
  await chmod(ENV_PATH, 0o600);
}
