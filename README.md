# Claude Routine: Outlook Sorter

Two scheduled Claude routines over one Outlook mailbox:

- **The sorter** — labels mail with Outlook **categories**, learns from
  corrections made in Outlook itself, and then teaches Outlook's own rules to keep
  doing it without the routine running. Runs hourly.
- **The end-of-week summary** — what is still waiting on a reply, what she flagged,
  what she never opened, unanswered invitations, and the week ahead. Runs Friday.

They share one sign-in and one state file, and neither needs the other.

Nothing moves and nothing is deleted. Categories are coloured tags, so every
action is visible in Outlook and reversible by hand.

Built for the Executive Director of a community foundation, so the default
categories are grants, donors, scholarships, board, events, and so on. They're all
editable in [src/taxonomy.ts](src/taxonomy.ts).

> **Why "routine"?** Most email sorters are a web app you have to visit. This one
> is a prompt on a schedule. The labels land in Outlook, where she already works —
> she never opens anything new, and there is no dashboard, no second inbox, and no
> database holding her mail.

---

## The sorter

**Three layers, cheapest first.**

1. **Sender rules** — free, instant, deterministic. `sender@example.org → Grants`.
   Grown from every correction. Handles the repetitive majority with no model call
   at all.
2. **Claude** — only for senders no rule covers. The routine prompt carries the
   category descriptions plus recent corrections as examples, which is the entire
   learning mechanism: no retraining, no embeddings, no vector store.
3. **A confidence gate** — anything that doesn't clear the bar, or that is nearly
   tied between two categories, gets `⚠ Needs Review` instead of a confident wrong
   guess. There is no human in the loop on a schedule, so this gate is the only
   thing standing between a guess and a label.

**Then promotion, which is the important part.** Once a sender has been confirmed
a few times, the routine writes it into a **native Outlook rule**. Native rules
run server-side on delivery, sync to every client, and cost nothing. From then on
that sender is categorized the moment mail arrives — no sweep required.

Only **user-confirmed** senders are promoted. A native rule keeps applying itself
invisibly, so baking in a guess would compound quietly.

**Corrections need no new habit.** She changes a category in Outlook the normal
way. Each message carries a hidden MAPI stamp recording what we assigned, so on
the next sweep any divergence is read as a correction. That writes a sender rule,
feeds the examples in the prompt, and counts toward promotion.

The smart layer teaches the always-on layer and gradually works itself out of a
job.

---

## The end-of-week summary

A Friday report on what the week left behind. Read-only — it reports on the
mailbox and calendar and changes nothing in either.

**Waiting on you** is the section that earns the report: mail addressed to her that
was never answered, oldest first. It is also the one with the most ways to be
wrong, so every rule protects precision over recall — better to under-report than
to cry wolf. A message counts as waiting only when she was on the **To** line
rather than CC'd, the sender looks like a person rather than a mailing list,
nothing went out from her mailbox in that thread afterwards, and it is older than a
grace period so this morning's mail isn't counted as a failure. A thread collapses
to one entry dated from the *first* unanswered ask, because that's the honest
answer to "how long have they been waiting?"

**The rest**, in the order the report puts them:

| | |
|---|---|
| Flagged | Messages she flagged and never cleared. Her own explicit "come back to this", so reported regardless of age or sender. |
| Needs an RSVP | Next week's unanswered invitations — the ones she can still act on. |
| Next week | Hours already committed, and what's on the calendar. |
| Unread | Grouped by sender and split people-vs-bulk. "37 unread" produces guilt and no action; "Tessa sent 3, plus 22 newsletters" produces a decision. |
| Your week | Meetings, hours, mail sent, busiest day. |

**On that last section:** Outlook records that a meeting existed and that mail was
sent — not what was decided or achieved, and not whether she actually attended,
only her RSVP. Phone calls, site visits, and anything done in another system are
invisible. It is framed as *activity*, and the prompt is explicitly told not to
dress it up as accomplishment.

**A late run reports the week it missed.** If the Friday slot is missed — a
throttled runner, an expired token — a Monday-through-Thursday run reports the
*previous* week rather than a week that has barely started, unless that week was
already reported. Without this, the week that mattered is never reported at all.

---

## What you need

- Node 20+
- A **work or school Microsoft 365 account** — the mailbox to be sorted.
- Permission to register an app in that account's directory. Most tenants let any
  user do this; some restrict it to administrators, and if yours does, that one
  step has to be done by whoever administers the Microsoft 365 account.
- A way to run a scheduled Claude task with this repository as its working
  directory.

No Gemini or OpenAI key. No database. No server. No hosting.

---

## Setup

```bash
npm install
npm run setup
```

That's it. `npm run setup` is a wizard: it walks through the one manual step,
takes the sign-in, writes its own configuration, prepares the mailbox, and
finishes by running a real sweep so you can see it working before you schedule
anything.

Budget about ten minutes, and **have the mailbox owner with you** — they sign in
at step 3, and whoever signs in is whose mail gets sorted.

What it does, in order:

1. **Register an app** — the only part that isn't automated. Four numbered clicks
   in the Entra portal, printed with the exact button names. A JSON shortcut is
   offered if you'd rather paste than click.
2. **Take the client and tenant IDs**, both from the same portal page.
3. **Sign in**, by device code — a URL and a short code, copied to your clipboard.
   Then it checks that every permission was actually granted, and names any that
   weren't.
4. **Prepare the mailbox** — creates the category labels, then reads how the
   mailbox is already filed to learn its regular senders, and tells you how many
   it found.
5. **Run one sweep** — reads the inbox and writes `routine/.local/plan.json`
   showing what it would label. Nothing is written to the mailbox in this step.

Every step verifies itself. A missing permission, a registration made as the wrong
platform, a tenant that blocks app creation — each is caught while you're still
sitting there, and named in words that say what to do about it.

### Then schedule them

The wizard prints this at the end. Create **two** scheduled Claude tasks, both with
this folder as their working directory and both carrying `STEWARD_CLIENT_ID`,
`STEWARD_TENANT`, and `STEWARD_REFRESH_TOKEN` in their environment — all three are
written into `.env` by setup. Copy them into the runner as secrets.

| | Prompt | When |
|---|---|---|
| Sorter | [routine/PROMPT.md](routine/PROMPT.md) | Hourly during the working day |
| Weekly summary | [routine/PROMPT-WEEKLY.md](routine/PROMPT-WEEKLY.md) | Friday mid-afternoon |

Sorting more often than hourly buys little — promoted senders are already being
labelled on arrival by Outlook itself.

Either can be scheduled without the other.

### If something goes wrong

Run `npm run setup` again. It's safe to repeat: it re-signs in and rewrites `.env`,
and the mailbox work it does is idempotent.

<details>
<summary>What the app registration needs, if you'd rather do it by hand</summary>

1. [Entra portal → App registrations](https://aka.ms/appregistrations) → **New
   registration**. Name it anything.
2. **Authentication** → **Add a platform** → **Mobile and desktop applications**
   → check `https://login.microsoftonline.com/common/oauth2/nativeclient`.
3. Same page: **Allow public client flows** → **Yes**.
4. **API permissions** → **Microsoft Graph** → **Delegated permissions**:
   `Mail.ReadWrite`, `MailboxSettings.ReadWrite`, `Calendars.Read`, `User.Read`,
   `offline_access`. None require admin consent.
5. Copy the **Application (client) ID** and **Directory (tenant) ID** from
   **Overview**.

It must be a **public client**, not a single-page application. Refresh tokens
issued to an SPA are capped at 24 hours by Entra, which is useless for a scheduled
job; public-client tokens last 90 days and the window rolls forward on every run.

These are *delegated* permissions, so the credential can only ever reach the one
mailbox that signs in. It grants nothing over anyone else's mail.

Then `npm run login` instead of `npm run setup`.
</details>

---

## For the person whose mailbox this is

**The labels appear on your mail by themselves.** Nothing to open, nothing to
press.

**To correct one, just change the category in Outlook the way you normally
would.** Right-click the message, pick the category you actually wanted. Next
sweep notices you disagreed and learns from it: that sender becomes permanently
correct, and similar mail from *new* senders gets better too. The first week or
two you'll be correcting it a fair bit, then it quiets down. That's the design
working, not a fault.

**Anything it isn't sure about gets `⚠ Needs Review`** rather than guessed at. A
wrong label is worse than an honest "I don't know."

**To undo everything:** delete the categories from Outlook's category list, which
removes every label it applied; and delete the `Inbox Steward:` rules under
Outlook's own Rules settings, which stops the automatic sorting.

---

## Where state lives

In a **hidden mail folder** in the mailbox, called `Inbox Steward`, holding one
message whose body is JSON. Hidden folders don't appear in any Outlook client, so
it doesn't clutter anything.

This is deliberate: nothing about this mailbox is stored on anyone else's server,
and deleting that folder deletes every trace of the tool. Set
`STEWARD_STATE_FILE=/path/to/state.json` to use a local file instead — that's for
development.

One honest note: the rotated refresh token is kept in that state, which means a
key to the mailbox is stored inside the mailbox it opens. That's circular, and
it's a deliberate trade for not requiring a secret store. If you'd rather it
weren't, set `STEWARD_REFRESH_TOKEN` in the environment on every run and delete the
`routine.refreshToken` field — the environment variable always wins.

---

## What leaves the mailbox

For senders no rule covers yet, the sender, subject, and roughly the first 600
characters go to Claude. `dataSharing` in [src/taxonomy.ts](src/taxonomy.ts) has
three settings:

| Setting | What leaves |
|---|---|
| `full` | Sender, subject, and a short preview. Best accuracy. *(default)* |
| `metadata` | Sender and subject only. No message text, ever. |
| `rules` | Nothing. Only senders it has already learned get labelled. |

Everything else — the sender rules, the corrections, the state — stays in the
mailbox.

---

## Tuning it

Almost all the accuracy lives in the **category descriptions** in
[src/taxonomy.ts](src/taxonomy.ts). They're written for a model to read: concrete
nouns, real program names, and explicit disambiguation against neighbouring
categories — spelling out, for instance, that a nonprofit asking for money is
*Grants* while the same nonprofit asking about a workshop is *Nonprofit Partners*.

When something is consistently misfiled, sharpen the description rather than
touching the code.

Other knobs, in `DEFAULT_SETTINGS`:

- `confidenceThreshold` — lower labels more mail, higher sends more to Needs
  Review.
- `promoteThreshold` — how many confirmations a sender needs before it earns a
  native Outlook rule. Higher is more conservative.
- `STEWARD_INBOX_WINDOW` (env, default 120) — how many recent inbox messages one
  sorting sweep looks at.

For the weekly summary:

- `STEWARD_TIMEZONE` (default `America/Denver`) — sets the week's boundaries. Worth
  getting right: it shifts the whole reporting window, and a cloud runner is
  almost certainly on UTC.
- `STEWARD_ALSO_ADDRESSED_AS` — comma-separated role aliases and shared mailboxes
  she is also reached at, e.g. `grants@example.org,info@example.org`. **Set this if
  it applies.** "Was she actually asked?" is decided by looking for her address on
  the To line, so mail to an alias the digest doesn't know about never appears as
  waiting — and the failure is silent, showing up as a suspiciously short list
  rather than an error.
- `STEWARD_WAITING_GRACE_HOURS` (default 18) — how long mail may sit before it
  counts as waiting.
- `STEWARD_IGNORED_SENDERS` — extra sender patterns never counted as waiting, added
  to the built-in list. `@example.org` matches the domain, `news@` matches the
  start of the local part, anything else is a substring.

---

## Commands

| | |
|---|---|
| `npm run setup` | Guided first-time setup. Run this first. |
| `npm run login` | Re-sign-in only, without the rest of setup. |
| `npm run plan` | Learn corrections, run sender rules, write `plan.json`. |
| `npm run apply` | Gate the verdicts, write categories, promote senders. |
| `npm run weekly` | Read the week and write `digest.json`. Touches nothing. |
| `npm test` | 64 offline checks. No network, no mailbox, no key. |
| `npm run typecheck` | |

The split is the point: deterministic mailbox I/O stays in code, where it's cheap
and repeatable. Claude is asked to do only the part that actually needs a model.

---

## Known limits

- **The weekly summary can only see Outlook.** Phone calls, site visits, hallway
  conversations, and work done in any other system are invisible to it. It reports
  activity, not accomplishment, and it should never be read as a performance
  measure.
- **No approval step on the sorter.** A schedule has no screen to ask on, so labels
  are always written directly. The confidence gate is the only safeguard, which is why the
  honesty-about-confidence instruction in the prompt is load-bearing rather than
  decorative.
- **No UI for editing categories.** Descriptions are where nearly all the accuracy
  lives, and editing them means editing `src/taxonomy.ts`. If she needs to do that
  herself, use the add-in version instead (below).
- **90-day credential ceiling.** Entra expires the refresh token if it goes unused
  for 90 days, and revokes it on a password change. The routine then fails with
  `invalid_grant`, and someone has to run `npm run setup` again. Whoever installs
  this is the person who will hear about that; she can't fix it herself.
- **Native rule quota.** Exchange allows roughly 256 KB of rule data per mailbox.
  Senders are consolidated at up to 60 per rule and promotion stops at 160 KB to
  leave room.
- **Delegates can't manage the master category list**, so the first run needs the
  mailbox owner rather than someone with delegated access.
- **December 31, 2026:** a new admin-consent `Mail-Advanced.ReadWrite` permission
  will gate modifying *sensitive* message properties. The evidence says it covers
  subject/body/recipients and **not** `categories`, but that came from a secondary
  source — worth re-checking before that date.

---

## The other version

This is the scheduled counterpart to
[southwest-community-foundation-email-sorter-and-end-of-week-summary](https://github.com/aaronv02/southwest-community-foundation-email-sorter-and-end-of-week-summary),
an **Outlook add-in** running the same classifier from a taskpane. They share the
taxonomy, the gate, the sender rules, and the promotion logic.

| | Add-in | This routine |
|---|---|---|
| Sorts new mail unattended | Only once a sender is promoted | Yes, every sweep |
| End-of-week summary | No | Yes |
| Stored credential | None at all | A 90-day refresh token |
| UI in Outlook | Full taskpane | Labels only |
| Edit categories in-app | Yes | No — edit the source |
| Classifier | Gemini, from the browser | Claude, on a schedule |

**Run one or the other, not both.** They share the provenance stamp on each
message, so corrections stay detectable either way, but they keep separate sender
rules and would both try to write `Inbox Steward:` rules.

## License

MIT.
