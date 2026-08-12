import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { homedir, platform } from 'node:os';

/**
 * Registering this project as an MCP server with Claude Desktop.
 *
 * Claude Desktop launches MCP servers itself, as child processes, reading them
 * from a JSON config file. So "connecting" means adding one entry to that file.
 */

const HERE = dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = resolve(HERE, '..');

export const SERVER_KEY = 'outlook-sorter';

/**
 * Where Claude Desktop keeps its config.
 *
 * Linux is included for development only - Claude Desktop does not ship there,
 * so the path is a guess and nothing will read it.
 */
export function configPath(): string {
  if (platform() === 'win32') {
    const appData = process.env.APPDATA ?? join(homedir(), 'AppData', 'Roaming');
    return join(appData, 'Claude', 'claude_desktop_config.json');
  }
  if (platform() === 'darwin') {
    return join(homedir(), 'Library', 'Application Support', 'Claude', 'claude_desktop_config.json');
  }
  return join(homedir(), '.config', 'Claude', 'claude_desktop_config.json');
}

/**
 * The command Claude Desktop should run.
 *
 * Absolute paths to the current Node binary and to tsx's entry point, rather than
 * `npm run mcp` or `npx tsx`. Claude Desktop spawns servers without a shell and
 * without the user's interactive PATH, so a bare `npm` or `npx` frequently isn't
 * found - and on Windows it would need the `.cmd` shim, which spawn won't resolve
 * on its own. Absolute paths sidestep the whole category of problem.
 */
export function serverEntry(): { command: string; args: string[]; cwd: string } {
  return {
    command: process.execPath,
    args: [
      resolve(PROJECT_ROOT, 'node_modules', 'tsx', 'dist', 'cli.mjs'),
      resolve(PROJECT_ROOT, 'routine', 'mcp.ts'),
    ],
    cwd: PROJECT_ROOT,
  };
}

export interface ConnectResult {
  path: string;
  /** True when an entry under our key was already there and got replaced. */
  replaced: boolean;
  /** Other servers found in the file, left untouched. */
  preserved: string[];
}

/**
 * Add or update our entry, preserving everything else in the file.
 *
 * Read-modify-write rather than write: she may well have other MCP servers
 * configured, and clobbering them to install ours would be an unforced error that
 * is invisible until something else stops working.
 */
export async function connectToDesktop(atPath?: string): Promise<ConnectResult> {
  // The path is injectable so the merge can be tested without touching the real
  // config - a test that clobbers someone's actual MCP servers is worse than no
  // test at all.
  const path = atPath ?? configPath();

  let config: Record<string, unknown> = {};
  try {
    const parsed = JSON.parse(await readFile(path, 'utf8')) as Record<string, unknown>;
    if (parsed && typeof parsed === 'object') config = parsed;
  } catch {
    // No file, or unreadable. Either way we start from an empty object rather
    // than refusing - a missing config is the normal first-run case.
  }

  const servers =
    config.mcpServers && typeof config.mcpServers === 'object'
      ? (config.mcpServers as Record<string, unknown>)
      : {};

  const replaced = Object.prototype.hasOwnProperty.call(servers, SERVER_KEY);
  const preserved = Object.keys(servers).filter((key) => key !== SERVER_KEY);

  servers[SERVER_KEY] = serverEntry();
  config.mcpServers = servers;

  await mkdir(dirname(path), { recursive: true });
  await writeFile(path, `${JSON.stringify(config, null, 2)}\n`, 'utf8');

  return { path, replaced, preserved };
}

export async function disconnectFromDesktop(atPath?: string): Promise<{ path: string; removed: boolean }> {
  const path = atPath ?? configPath();

  let config: Record<string, unknown>;
  try {
    config = JSON.parse(await readFile(path, 'utf8')) as Record<string, unknown>;
  } catch {
    return { path, removed: false };
  }

  const servers = config.mcpServers as Record<string, unknown> | undefined;
  if (!servers || !Object.prototype.hasOwnProperty.call(servers, SERVER_KEY)) {
    return { path, removed: false };
  }

  delete servers[SERVER_KEY];
  await writeFile(path, `${JSON.stringify(config, null, 2)}\n`, 'utf8');
  return { path, removed: true };
}
