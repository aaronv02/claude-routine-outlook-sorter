/**
 * Runs one of the routine prompts through the Claude Code CLI, headless.
 *
 *   npm run sort       - the hourly sorting sweep
 *   npm run summary    - the Friday end-of-week summary
 *
 * This is what a scheduled task on her own machine invokes. The two prompts are
 * otherwise meant to be pasted into a hosted scheduled task; this exists so the
 * whole thing can live on one computer with nothing in the cloud.
 *
 * Every run appends to a log file. When something breaks weeks later, on a machine
 * nobody is sitting at, the log is the only evidence there will be - and a
 * scheduled task that fails silently is worse than no scheduled task at all.
 */

import { spawn } from 'node:child_process';
import { readFile, writeFile, mkdir, appendFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { platform } from 'node:process';

import { extractPrompt } from './prompt-file.js';
import { loadEnv } from './env.js';

const HERE = dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = resolve(HERE, '..');

const JOBS = {
  sort: { file: 'PROMPT.md', log: 'sort.log', label: 'sorting sweep' },
  summary: { file: 'PROMPT-WEEKLY.md', log: 'weekly.log', label: 'weekly summary' },
} as const;

type JobName = keyof typeof JOBS;

/**
 * Generous, because a first run bootstraps from hundreds of already-filed messages
 * and a weekly summary reads six weeks of mail. Still bounded, so a wedged run
 * cannot overlap the next scheduled one indefinitely.
 */
const TIMEOUT_MS = 20 * 60_000;

async function logTo(logPath: string, text: string): Promise<void> {
  await mkdir(dirname(logPath), { recursive: true });
  await appendFile(logPath, text, 'utf8');
}

async function main(): Promise<void> {
  const name = process.argv[2] as JobName | undefined;
  const job = name ? JOBS[name] : undefined;

  if (!job) {
    console.error(`usage: tsx routine/run-prompt.ts <${Object.keys(JOBS).join('|')}>`);
    process.exit(1);
  }

  await loadEnv();

  const logPath = resolve(PROJECT_ROOT, 'routine', '.local', 'logs', job.log);
  const started = new Date();
  const stamp = started.toISOString();

  await logTo(logPath, `\n===== ${stamp}  ${job.label} starting\n`);

  let prompt: string;
  try {
    const document = await readFile(resolve(PROJECT_ROOT, 'routine', job.file), 'utf8');
    prompt = extractPrompt(document, job.file);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    await logTo(logPath, `FAILED to read the prompt: ${message}\n`);
    console.error(message);
    process.exit(1);
  }

  // Also written to disk, purely so a failed unattended run can be diagnosed:
  // this is the exact text that was handed to the model.
  const promptCopy = resolve(PROJECT_ROOT, 'routine', '.local', `prompt-${name}.txt`);
  await mkdir(dirname(promptCopy), { recursive: true });
  await writeFile(promptCopy, prompt, 'utf8');

  const output = await runClaude(prompt, logPath);

  const seconds = Math.round((Date.now() - started.getTime()) / 1000);
  await logTo(logPath, `----- finished in ${seconds}s, exit ${output.code}\n`);

  // The model's own summary goes to stdout as well as the log, so running this by
  // hand shows what a scheduled run would have produced.
  if (output.text.trim() !== '') console.log(output.text.trim());

  if (output.code !== 0) {
    console.error(`\nThe ${job.label} failed. Full log: ${logPath}`);
    process.exit(output.code);
  }

  // Exit zero with nothing to show is the failure mode worth shouting about: the
  // scheduled task reports success, week after week, having done nothing. Treated
  // as an error so it surfaces in the log and in the task's exit code.
  if (output.text.trim() === '') {
    const message =
      'The CLI exited successfully but produced no output, so nothing was done. ' +
      'Check that `claude` is signed in: run `claude` once in this folder.';
    await logTo(logPath, `EMPTY RESULT: ${message}\n`);
    console.error(message);
    process.exit(3);
  }
}

/**
 * The fixed instruction passed as the CLI's query argument.
 *
 * `claude -p` takes its query from argv; stdin is piped *content*, per the
 * documented `cat file | claude -p "query"` form. So the long prompt travels as
 * content and this one short sentence is the query that points at it.
 *
 * It is a constant, authored here, containing no shell metacharacters - which is
 * what makes passing it as an argument safe where passing the prompt itself was
 * not.
 */
const QUERY = 'The text piped to you is your task. Follow every instruction in it exactly.';

/**
 * Run one prompt through the CLI, headless.
 *
 * The prompt goes in as piped content on STDIN, never as an argument, and the
 * argument is the fixed QUERY above. Three reasons, each sufficient on its own:
 *
 *  - It is the documented shape. `claude -p` reads its query from argv and treats
 *    stdin as content to work on. An earlier version of this file piped the prompt
 *    with no query at all, which would have run the schedule and produced nothing -
 *    the worst possible failure, because it looks like success.
 *  - Safety. Passing the prompt in argv alongside `shell: true` hands the whole
 *    thing to a shell for interpretation: its code fences, backticks and `$` become
 *    commands. An even earlier version did exactly that, and running it produced a
 *    stream of `/bin/sh: Do: command not found` while the shell executed fragments
 *    of the instructions. Node warns about this (DEP0190) because arguments are
 *    concatenated unescaped.
 *  - Size and shape. These prompts are multi-line and a few kilobytes. Command
 *    lines have length limits, and newlines inside a Windows `cmd.exe` argument are
 *    their own category of problem.
 *
 * Windows still needs `cmd.exe` to resolve the CLI's `.cmd` shim, but it is invoked
 * with an explicit argument array rather than a shell string, so nothing is
 * re-parsed.
 */
function runClaude(
  prompt: string,
  logPath: string,
): Promise<{ code: number; text: string }> {
  return new Promise((done) => {
    const isWindows = platform === 'win32';
    const command = isWindows ? 'cmd.exe' : 'claude';
    const args = isWindows ? ['/c', 'claude', '-p', QUERY] : ['-p', QUERY];

    const child = spawn(command, args, {
      cwd: PROJECT_ROOT,
      shell: false,
      stdio: ['pipe', 'pipe', 'pipe'],
    });

    child.stdin.on('error', () => {
      // The CLI exiting before the prompt is fully written shows up here as EPIPE.
      // The exit code below is the real diagnosis; this would only add noise.
    });
    child.stdin.end(prompt, 'utf8');

    let text = '';
    let failed = false;

    const timer = setTimeout(() => {
      failed = true;
      void logTo(logPath, `TIMED OUT after ${TIMEOUT_MS / 60_000} minutes; killed.\n`);
      child.kill();
    }, TIMEOUT_MS);

    child.stdout.on('data', (chunk: Buffer) => {
      const piece = chunk.toString();
      text += piece;
      void logTo(logPath, piece);
    });

    child.stderr.on('data', (chunk: Buffer) => {
      void logTo(logPath, chunk.toString());
    });

    child.on('error', (err) => {
      clearTimeout(timer);
      const hint =
        (err as NodeJS.ErrnoException).code === 'ENOENT'
          ? 'The Claude Code CLI was not found. Install it with: npm install -g @anthropic-ai/claude-code'
          : err.message;
      void logTo(logPath, `FAILED to start claude: ${hint}\n`);
      console.error(hint);
      done({ code: 1, text });
    });

    child.on('close', (code) => {
      clearTimeout(timer);
      done({ code: failed ? 124 : (code ?? 1), text });
    });
  });
}

main().catch(async (err) => {
  console.error(err instanceof Error ? err.message : String(err));
  process.exit(1);
});
