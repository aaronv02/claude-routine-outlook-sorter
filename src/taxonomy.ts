import type { Category, Settings, PersistedState } from './types.js';

/**
 * The bucket for anything the classifier isn't sure about.
 *
 * This category is load-bearing. A confident wrong label costs more trust than
 * an admitted unknown, so everything below the confidence threshold lands here
 * rather than in a plausible-looking guess.
 */
export const NEEDS_REVIEW_ID = 'needs-review';

/**
 * Cold-start taxonomy for the Executive Director of a community foundation,
 * derived from the foundation's actual programs and funds.
 *
 * There is no filing history to learn from on day one, so these descriptions are
 * the entire model at first. They are written for an LLM reader: concrete
 * nouns, real program names, and explicit disambiguation against neighbouring
 * categories. All of it is editable in the taskpane.
 */
export const SEED_TAXONOMY: Category[] = [
  {
    id: 'donors',
    name: 'Donors & Gifts',
    color: 'Preset0',
    description:
      'Correspondence with individual donors, families, and businesses who give. Gift notifications and acknowledgements, questions about opening or adding to a donor advised fund, field of interest fund, or designated fund, planned giving and bequest conversations, stock and QCD transfers, and thank-you correspondence. Distinguish from Finance: a gift arriving from a named person is Donors & Gifts, whereas a custodian statement or platform payout report is Finance.',
  },
  {
    id: 'grants',
    name: 'Grants',
    color: 'Preset1',
    description:
      'Anything from or about organizations seeking money. Letters of inquiry, grant applications and attachments, questions about eligibility or deadlines, grant agreements, interim and final grant reports, declines and appeals, and the LAUNCH Fund and Community Emergency Relief Fund pipelines. Distinguish from Nonprofit Partners: a nonprofit asking for or reporting on funding is Grants, whereas the same nonprofit asking about a workshop or fiscal sponsorship is Nonprofit Partners.',
  },
  {
    id: 'scholarships',
    name: 'Scholarships',
    color: 'Preset2',
    description:
      'Student scholarship applicants and their families, transcripts and recommendation letters, scholarship review committee scheduling and deliberations, award notifications, disbursement and enrollment verification with schools, and scholarship fund donors asking about their named award. Student-facing mail belongs here rather than in Grants.',
  },
  {
    id: 'board',
    name: 'Board & Governance',
    color: 'Preset3',
    description:
      'Board of directors and committee business. Meeting agendas, packets, minutes, board member recruitment and onboarding, conflict of interest and policy documents, bylaws, audit and finance committee mail, executive committee threads, and strategic planning. Mail from a board member wearing a donor hat is Donors & Gifts; mail about governance is here.',
  },
  {
    id: 'finance',
    name: 'Finance',
    color: 'Preset4',
    description:
      'Money mechanics and record keeping. Investment and custodian statements, bank notices, ColoradoGives and other giving-platform payout and disbursement reports, invoices and bills, bookkeeping and QuickBooks correspondence, payroll, audit fieldwork, 990 preparation, and insurance. Automated statements and reports belong here even when they concern donor gifts.',
  },
  {
    id: 'events',
    name: 'Events',
    color: 'Preset5',
    description:
      'Logistics for the foundation\'s fundraising and community events: Durango Wine Experience, Hoedown at the Mancos Opera House, 19th Hole Concerts, Payroll Department\'s Pitch Palooza, Making a Difference Speaker Series, and Tips & Tricks or Year-End Ask workshop sessions. Includes sponsors and sponsorship packets, venues, caterers, ticketing, auction items, volunteers, and run-of-show. Event sponsorship solicitation is Events, not Donors & Gifts.',
  },
  {
    id: 'partners',
    name: 'Nonprofit Partners',
    color: 'Preset6',
    description:
      'The regional nonprofit sector the foundation serves, in a non-funding capacity. Fiscal sponsorship arrangements, professional development workshop registration and questions, agency fund holders, capacity-building requests, collaboration and referral, and the CAUSE Youth Internship and DWE Nonprofit Partners programs. If they are asking for grant money it is Grants; if they are asking for help, training, or partnership it is here.',
  },
  {
    id: 'press',
    name: 'Press & Community',
    color: 'Preset7',
    description:
      'Outward-facing community communication. Reporters and media requests, press releases, interview and podcast invitations, award nominations, letters to the editor, speaking invitations, community announcements from Durango and the five-county region (Archuleta, La Plata, Dolores, Montezuma, San Juan), and chamber or civic group mail.',
  },
  {
    id: 'sector',
    name: 'Sector News',
    color: 'Preset8',
    description:
      'Bulk philanthropy-sector reading with no action required. Newsletters and bulletins from the Council on Foundations, Colorado Nonprofit Association, Philanthropy Colorado, Chronicle of Philanthropy, Candid, webinar and conference promotions, and sector research digests. Characteristically a mailing list with an unsubscribe link and no personal addressing. Low urgency by definition.',
  },
  {
    id: 'vendors',
    name: 'Vendors & Admin',
    color: 'Preset9',
    description:
      'Running the office. Software and SaaS notifications and renewals, IT and Microsoft 365 service mail, password resets and security alerts, office supplies, phone and internet, landlord and facilities, professional memberships and dues, and general administrative housekeeping. Automated transactional mail from tools belongs here unless it concerns money movement, which is Finance.',
  },
  {
    id: NEEDS_REVIEW_ID,
    name: '⚠ Needs Review',
    color: 'Preset10',
    description:
      'Reserved. Never choose this category directly - it is applied automatically when no other category clears the confidence threshold.',
  },
];

/**
 * A scheduled sweep has no screen to ask on, so labels are always written
 * directly. The confidence gate is the approval step: anything that doesn't
 * clear `confidenceThreshold` becomes ⚠ Needs Review rather than a guess.
 */
export const DEFAULT_SETTINGS: Settings = {
  dataSharing: 'full',
  confidenceThreshold: 0.65,
  promoteThreshold: 3,
  bootstrapped: false,
  generation: 0,
};

export const STATE_VERSION = 1;

export function defaultState(): PersistedState {
  return {
    version: STATE_VERSION,
    settings: { ...DEFAULT_SETTINGS },
    taxonomy: SEED_TAXONOMY.map((c) => ({ ...c })),
    senderRules: [],
    recentCorrections: [],
  };
}

/**
 * Outlook's own ceiling on a category display name.
 */
const CATEGORY_NAME_MAX = 255;

/**
 * Make a user-typed category name safe to round-trip through Outlook.
 *
 * The comma is the one that matters. Outlook treats a category list as
 * comma-delimited in several of its representations, so a category named
 * "Grants, Pending" comes back from Graph as two categories - "Grants" and
 * "Pending". That breaks this add-in in a way that is worse than cosmetic:
 * `detectCorrections` maps the names on a message back to category ids, so a
 * split name either matches nothing (the correction is silently never learned)
 * or, if the fragment happens to collide with a real category, is read as a
 * correction TO that category. The second case is the dangerous one - a wrong
 * signal that promotion can then bake into a permanent native Outlook rule.
 *
 * Control characters get the same treatment for the same reason: they survive
 * a POST and then compare unequal on the way back.
 */
export function sanitizeCategoryName(raw: string): string {
  return Array.from(raw)
    .map((char) => {
      if (char === ',') return ' ';
      const code = char.codePointAt(0) ?? 0;
      return code <= 31 || code === 127 ? ' ' : char;
    })
    .join('')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, CATEGORY_NAME_MAX)
    .trim();
}

/**
 * Resolve a name that no other category is already using.
 *
 * Two categories sharing a name collapses them everywhere it counts: the
 * name-to-id map in `detectCorrections` keeps only one, and
 * `ensureMasterCategories` dedupes by name so Outlook only ever gets one label
 * for the pair. Clicking "Add category" twice is enough to cause it, so this is
 * a likely accident rather than a hypothetical one.
 */
export function uniqueCategoryName(
  desired: string,
  taxonomy: Category[],
  selfId?: string,
): string {
  const taken = new Set(
    taxonomy.filter((c) => c.id !== selfId).map((c) => c.name.trim().toLowerCase()),
  );
  if (!taken.has(desired.trim().toLowerCase())) return desired;

  for (let n = 2; ; n++) {
    const candidate = `${desired} ${n}`;
    if (!taken.has(candidate.trim().toLowerCase())) return candidate;
  }
}

/**
 * Repair a taxonomy read back from storage.
 *
 * State persisted before names were validated can still hold a comma or a
 * duplicate, and every read path trusts these names, so the fix belongs at the
 * boundary rather than at each use site.
 */
export function normalizeTaxonomy(taxonomy: Category[]): Category[] {
  const seen: Category[] = [];
  for (const category of taxonomy) {
    const sanitized = sanitizeCategoryName(category.name);
    // An empty result means the name was punctuation and nothing else. Falling
    // back to the id keeps the category addressable instead of nameless.
    const base = sanitized || category.id;
    const name = uniqueCategoryName(base, seen, category.id);
    seen.push(name === category.name ? category : { ...category, name });
  }
  return seen;
}

/** Categories the user may actually pick, i.e. everything except the gate bucket. */
export function selectableCategories(taxonomy: Category[]): Category[] {
  return taxonomy.filter((c) => c.id !== NEEDS_REVIEW_ID);
}

export function categoryById(taxonomy: Category[], id: string): Category | undefined {
  return taxonomy.find((c) => c.id === id);
}

export function categoryByName(taxonomy: Category[], name: string): Category | undefined {
  const needle = name.trim().toLowerCase();
  return taxonomy.find((c) => c.name.trim().toLowerCase() === needle);
}

// ---------------------------------------------------------------------------
// Matching her existing categories to ours
// ---------------------------------------------------------------------------

/**
 * Words carrying no signal about what a category is for.
 */
const STOPWORDS = new Set(['and', 'the', 'of', 'for', 'to', 'a', 'my', 'our']);

/**
 * A token has to be this long before matching it alone is worth anything.
 * Without this, "IT" and "Ops" collide with unrelated categories.
 */
const DISTINCTIVE_TOKEN_LENGTH = 4;

/** Best score we'll accept as a match at all. */
const MATCH_FLOOR = 0.6;

/**
 * How far ahead of the runner-up the winner must be.
 *
 * Same principle as the confidence gate: a narrow margin means something fits
 * but we can't tell what, and guessing is worse than declining. It matters more
 * here than at the gate, because a bootstrap match creates a CONFIRMED sender
 * rule, and confirmed rules are what get promoted into permanent native Outlook
 * rules. A wrong guess here outlives the add-in.
 */
const MATCH_MARGIN = 0.15;

/**
 * Crude plural folding, applied to both sides so it only ever aligns them.
 *
 * Without it "Donor" declines while "Donors" matches, which is an arbitrary
 * distinction to draw over how someone happened to name a folder. Guarded by
 * length so short words aren't mangled into collisions - "news" stays "news".
 */
function stem(token: string): string {
  return token.length > 4 && token.endsWith('s') ? token.slice(0, -1) : token;
}

function tokenize(name: string): string[] {
  return name
    .toLowerCase()
    .split(/[^a-z0-9]+/)
    .filter((t) => t.length > 0 && !STOPWORDS.has(t))
    .map(stem);
}

/** Character bigrams, for similarity that survives plurals and word order. */
function bigrams(s: string): string[] {
  const flat = s.toLowerCase().replace(/[^a-z0-9]/g, '');
  const out: string[] = [];
  for (let i = 0; i < flat.length - 1; i++) out.push(flat.slice(i, i + 2));
  return out;
}

function diceCoefficient(a: string, b: string): number {
  const left = bigrams(a);
  const right = bigrams(b);
  if (left.length === 0 || right.length === 0) return 0;

  const pool = new Map<string, number>();
  for (const g of left) pool.set(g, (pool.get(g) ?? 0) + 1);

  let shared = 0;
  for (const g of right) {
    const n = pool.get(g) ?? 0;
    if (n > 0) {
      shared++;
      pool.set(g, n - 1);
    }
  }
  return (2 * shared) / (left.length + right.length);
}

/**
 * How well one of her category names corresponds to one of ours, 0 to 1.
 *
 * Two views, best one wins. Containment catches the common real case, where her
 * name is a shorter version of ours - "Donors" against "Donors & Gifts", "Admin"
 * against "Vendors & Admin". Bigram similarity catches the rest, where the words
 * differ in form rather than meaning.
 */
export function categoryNameSimilarity(theirs: string, ours: string): number {
  const mine = tokenize(ours);
  const hers = tokenize(theirs);
  if (mine.length === 0 || hers.length === 0) return 0;

  const mineSet = new Set(mine);
  const shared = hers.filter((t) => mineSet.has(t));

  // Containment only counts on a token substantial enough to mean something.
  const distinctive = shared.some((t) => t.length >= DISTINCTIVE_TOKEN_LENGTH);
  const containment = distinctive ? shared.length / Math.min(hers.length, mine.length) : 0;

  return Math.max(containment, diceCoefficient(theirs, ours));
}

export interface CategoryMatch {
  categoryId: string;
  score: number;
  /** True when her name is ours exactly, so no judgement was involved. */
  exact: boolean;
}

/**
 * Map one of her existing Outlook categories onto ours, or decline.
 *
 * Declining is the common and correct outcome for categories that are genuinely
 * hers alone - a fund name, a person, a project. Those stay unmapped and their
 * mail goes to the classifier like anything else.
 */
export function matchExistingCategory(
  theirs: string,
  taxonomy: Category[],
): CategoryMatch | null {
  const candidates = selectableCategories(taxonomy);

  const exact = candidates.find(
    (c) => c.name.trim().toLowerCase() === theirs.trim().toLowerCase(),
  );
  if (exact) return { categoryId: exact.id, score: 1, exact: true };

  const scored = candidates
    .map((c) => ({ categoryId: c.id, score: categoryNameSimilarity(theirs, c.name) }))
    .sort((a, b) => b.score - a.score);

  const best = scored[0];
  if (!best || best.score < MATCH_FLOOR) return null;

  const runnerUp = scored[1];
  if (runnerUp && best.score - runnerUp.score < MATCH_MARGIN) return null;

  return { categoryId: best.categoryId, score: best.score, exact: false };
}

/**
 * Resolve every category name seen in her mailbox against our taxonomy.
 *
 * Returned as a lookup keyed by her lowercased name, plus the list of inexact
 * matches so the first run can show its work rather than silently reinterpreting
 * how she files things.
 */
export function mapExistingCategories(
  names: Iterable<string>,
  taxonomy: Category[],
): { lookup: Map<string, string>; inferred: { from: string; toCategoryId: string }[] } {
  const lookup = new Map<string, string>();
  const inferred: { from: string; toCategoryId: string }[] = [];
  const seen = new Set<string>();

  for (const raw of names) {
    const key = raw.trim().toLowerCase();
    if (!key || seen.has(key)) continue;
    seen.add(key);

    const match = matchExistingCategory(raw, taxonomy);
    if (!match) continue;

    lookup.set(key, match.categoryId);
    if (!match.exact) inferred.push({ from: raw.trim(), toCategoryId: match.categoryId });
  }

  return { lookup, inferred };
}
