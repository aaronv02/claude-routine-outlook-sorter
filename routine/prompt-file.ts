/**
 * Pulling the runnable prompt out of a prompt document.
 *
 * `PROMPT.md` is written for a person: it opens with an explanation of what the
 * routine is and where it fits, then a marker, then the prompt itself. Feeding the
 * whole file to a model would hand it the commentary as if it were instructions -
 * including the sentence telling the reader to paste the text below, which is
 * exactly the kind of thing a model will earnestly try to act on.
 *
 * So the marker is load-bearing rather than decorative, and this is the one place
 * that knows it.
 */

export const PROMPT_MARKER = '## Prompt (copy from here down)';

export function extractPrompt(document: string, sourceName = 'the prompt file'): string {
  const at = document.indexOf(PROMPT_MARKER);
  if (at === -1) {
    throw new Error(
      `${sourceName} has no "${PROMPT_MARKER}" heading, so there is no way to tell the explanation from the prompt.`,
    );
  }

  const body = document.slice(at + PROMPT_MARKER.length);

  // A leading "---" is the horizontal rule under the heading in some of these
  // documents; it carries no instruction.
  const cleaned = body.replace(/^\s*-{3,}\s*/, '').trim();

  if (cleaned === '') throw new Error(`${sourceName} has nothing after the marker.`);
  return cleaned;
}
