import type { Category } from '../src/types.js';

/**
 * The full category set to write to a message.
 *
 * Graph's PATCH on `categories` replaces the array wholesale rather than merging
 * into it, so every category that should survive the write has to be named in
 * it. That makes this the single most destructive line in the routine: a
 * category of hers that we fail to carry forward is silently deleted from her
 * mail, and unlike a wrong label there is nothing left to notice.
 *
 * So: keep everything that isn't ours, drop the ones that are, add the new one.
 * Dropping ours is what makes a re-label a replacement instead of an
 * accumulation - without it a message corrected twice ends up wearing three of
 * our labels at once.
 *
 * It lives in its own module rather than inline in `sweep.ts` so the test
 * harness can import it without executing the CLI.
 */
export function mergeCategories(
  existing: string[],
  ours: Category[],
  applying: string,
): string[] {
  const ourNames = new Set(ours.map((c) => c.name.trim().toLowerCase()));
  const foreign = existing.filter((name) => !ourNames.has(name.trim().toLowerCase()));

  // Case-insensitive, because Outlook compares category names that way and
  // would otherwise end up holding two spellings of the same label.
  const already = foreign.some((name) => name.trim().toLowerCase() === applying.trim().toLowerCase());

  return already ? foreign : [...foreign, applying];
}
