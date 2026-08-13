# Claude Routine: Outlook Sorter

Three ways into one Outlook mailbox, sharing one sign-in and one state file:

- **The sorter** — labels mail with Outlook **categories**, learns from corrections
  made in Outlook itself, and then teaches Outlook's own rules to keep doing it
  without the routine running. Scheduled hourly.
- **The end-of-week summary** — what is still waiting on a reply, what she flagged,
  what she never opened, unanswered invitations, and the week ahead. Scheduled
  Friday.
- **The MCP server** — the same analysis, on demand, in Claude Desktop. "Who is
  waiting on me?", "catch me up, I was out Tuesday", "what did I tell the board
  about the audit?"

The first two are push: something fires and a report appears. The third is pull,
for the questions a schedule can't anticipate. Any of them works without the
others.

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

## Asking it things (the MCP server)

The schedules can only answer questions someone thought of in advance. This is the
other direction: eight tools that let Claude read the live mailbox while she talks
to it.

| Tool | For |
|---|---|
| `whats_waiting` | "who is waiting on me?", "anything I've left hanging more than a week?" |
| `weekly_summary` | "how was my week?", "catch me up, I was out Tuesday and Wednesday" |
| `whats_next_week` | "what's coming up?", "any meetings I haven't responded to?" |
| `find_mail` | "anything from Tessa this month?", "what did I tell the board about the audit?" |
| `read_message` | The full text of one message, by reference |
| `list_categories` | The labels and what belongs in each |
| `suggest_labels` | Proposes labels for unlabelled mail. **Reads only.** |
| `apply_labels` | Writes labels — only after she has seen and agreed to them |

Results carry short references (`m1`, `m2`) rather than Graph's 150-character ids,
so she can say "read m3" and it resolves.

**The two write tools are deliberately split.** `suggest_labels` changes nothing,
and its description instructs Claude to show her the proposals and never call
`apply_labels` in the same turn. So the conversational path has an approval step
the schedule cannot have — and because she approved each one explicitly, those
labels count as confirmations toward a sender earning a native Outlook rule.
Promotion itself is left to the scheduled sweep: writing a permanent mailbox rule
as a side effect of a chat message is more than she agreed to.

It runs on her machine, launched by Claude Desktop over stdio, and only while
Claude is open. Nothing happens while she's away — that's what the schedules are
for.

---

## Setup

**On the machine that owns the mailbox**, open PowerShell (Windows) and paste:

```
irm https://raw.githubusercontent.com/aaronv02/claude-routine-outlook-sorter/main/install.ps1 | iex
```

macOS or Linux:

```bash
curl -fsSL https://raw.githubusercontent.com/aaronv02/claude-routine-outlook-sorter/main/install.sh | bash
```

That checks Node, downloads the project to `C:\outlook-sorter` (or `~/outlook-sorter`),
installs its dependencies, and starts the wizard. Nothing else to download and no
GitHub account needed — the repo is public.

Already have the folder? Just:

```bash
npm install
npm run setup
```

Either way, that's it. `npm run setup` is a wizard: it walks through the one manual step,
takes the sign-in, writes its own configuration, prepares the mailbox, and
finishes by running a real sweep so you can see it working before you schedule
anything.

Budget about fifteen minutes, and **have the mailbox owner with you** — they sign in
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
6. **Connect to Claude Desktop** — optional, and only useful on the machine where
   Claude Desktop is installed. Adds one entry to its config and leaves any other
   MCP servers alone.
7. **Install the schedules** — two Windows scheduled tasks, so the sorting and the
   Friday summary run on their own. Offers to install the Claude Code CLI if it is
   missing.

Every step verifies itself. A missing permission, a registration made as the wrong
platform, a tenant that blocks app creation — each is caught while you're still
sitting there, and named in words that say what to do about it.

### Then connect and schedule

Step 6 of the wizard offers to register the MCP server with Claude Desktop. After
it does, **fully quit Claude Desktop and reopen it** — closing the window is not
enough; quit it from the system tray. Then ask Claude *"who is waiting on me?"* to
confirm it can see the mailbox.

You can do this later instead with `npm run connect`, and undo it with
`npm run disconnect`.

### Then schedule them

Step 7 of the wizard offers to do this. It installs two Windows scheduled tasks:

| Task | When |
|---|---|
| `Outlook Sorter - Sort` | Hourly, 8:00–18:00 |
| `Outlook Sorter - Weekly` | Fridays at 15:00 |

Or run it yourself, with whatever hours suit:

```
.\scripts\install-tasks.ps1 -StartHour 7 -EndHour 19 -SummaryTime 16:30
```

Everything then lives on one computer. No cloud runner, no secret anywhere but that
machine, nothing depending on anyone else's account.

**This needs the Claude Code CLI**, which is what actually reads the mail and
decides — the installer offers to install it, and it must be signed in to Claude
once. Do that while you're still sitting there.

Both tasks run only while that user is logged on. That's deliberate rather than a
limitation papered over: running as a service would mean storing the account
password, and the weekly summary already reports the week it *missed* if a run
doesn't happen. A laptop closed on Friday afternoon is a case the tool handles, not
one it has to prevent.

Sorting more often than hourly buys little — promoted senders are already being
labelled on arrival by Outlook itself.

Remove them again with `.\scripts\uninstall-tasks.ps1`. That stops the schedule and
changes nothing else.

<details>
<summary>Running it in the cloud instead</summary>

Create two scheduled Claude tasks, both with this folder as their working directory
and both carrying `STEWARD_CLIENT_ID`, `STEWARD_TENANT`, and
`STEWARD_REFRESH_TOKEN` in their environment — all three are written into `.env` by
setup. Copy them in as secrets.

| | Prompt | When |
|---|---|---|
| Sorter | [routine/PROMPT.md](routine/PROMPT.md) | Hourly during the working day |
| Weekly summary | [routine/PROMPT-WEEKLY.md](routine/PROMPT-WEEKLY.md) | Friday mid-afternoon |

Runs whether the laptop is on or not, at the cost of a 90-day credential living in
someone else's runner. State lives in the mailbox either way, so an ephemeral
sandbox is fine.
</details>

Either routine can be scheduled without the other.

### Three things the installer will not do

Not oversights — each one needs a person:

1. **Install Node.** On a managed laptop that wants an administrator password, and
   silently installing a runtime on someone's work machine isn't an installer's
   business. It tells you where to get it and stops.
2. **Sign in.** The mailbox owner does that herself, in the wizard. That consent is
   the entire security model.
3. **Restart Claude Desktop.** The wizard says when; a human quits it from the
   system tray.

You also can't ask Claude Desktop to do the install for you — it has no terminal,
and the MCP server that would give it one is what the install creates.

### Where to put it

Somewhere **permanent and not synced**. Claude Desktop stores absolute paths into
the folder, so moving or deleting it later silently breaks the connection. And on
Microsoft 365 machines, Documents and Desktop are usually redirected into OneDrive,
where `node_modules` means thousands of files churning through sync and file-lock
failures during install. `C:\outlook-sorter` is the default for both reasons; the
installer warns if you point it somewhere synced.

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

## Does it work?

Two harnesses, both offline, neither needing a mailbox.

**`npm run accuracy`** scores model verdicts against 50 labelled fixture emails
written for this foundation — real program names, real fund names, the kinds of
message that actually arrive. It runs them through the *real* confidence gate, so it
measures what would land on the messages, and reports three numbers: top-1 (the
label it would write), top-3 (the right answer was offered), and **how often it is
confidently wrong**, which is the one that matters most. A wrong label teaches her
the tool can't be trusted; an honest ⚠ Needs Review costs her a moment.

`test/fixtures/verdicts-reference.json` is one recorded run: 50/50 top-1, zero
confidently wrong. **Treat that as a floor, not a forecast.** The fixtures were
written as clear exemplars of these categories, and real mail is messier —
forwarded threads, three-word subjects, one message covering two topics, people who
are a donor and a board member at once. The honest expectation on live mail is the
80% top-1 / 90% top-3 the published benchmarks suggest. If you want a real number,
run a week of her actual mail through `npm run plan` and check the labels by hand.

**`npm run scenario`** builds a fake week whose facts are known — a thread she never
answered, one she did, mail she was only CC'd on, a person at `goodnews@` who must
not be mistaken for bulk, a declined meeting, an all-day out-of-office marker — and
asserts the digest reaches the right conclusion about each. Run it after touching
anything in `src/digest/`. Numbers that look plausible are the failure mode, so it
states what the week actually contains rather than trusting the output to look
sensible.

**`npm run loop`** is the one that matters most. It runs the actual `plan` and
`apply` commands against a fake mailbox served over a stubbed `fetch` — so the real
orchestration, the real `src/graph.ts` with its JSON batching and provenance
encoding, the real state store writing to a hidden mail folder, and the real
promotion into native Outlook rules all execute. Four sweeps, with a human
correcting a label in between, answering the question the whole product rests on:

> Does correcting a label in Outlook, with no other action, eventually teach Outlook
> itself to do the job?

It asserts that it does — the correction is learned, later mail from that sender
never reaches the model again, and once confirmed enough times a native
`Inbox Steward:` rule appears with the right sender and the category *she* chose.
It intercepts at `fetch` rather than mocking the Graph module on purpose: mocking
`graph.ts` would skip exactly the parts most likely to be wrong.

All three found real bugs the first time they ran, and a fourth came out of auditing
a real inbox. See the notes at the end of this file.

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
| `npm run sort` | One sorting sweep, start to finish, via the Claude Code CLI. |
| `npm run summary` | One end-of-week summary, same way. |
| `npm run audit` | Check the sender filter against this mailbox. Read-only. |
| `npm run connect` | Register the MCP server with Claude Desktop. |
| `npm run disconnect` | Remove it again. Leaves other MCP servers alone. |
| `npm run mcp` | Run the MCP server by hand. Claude Desktop normally does this. |
| `npm test` | 84 offline checks. No network, no mailbox, no key. |
| `npm run accuracy <verdicts.json>` | Score sorting against the 50 labelled fixtures. |
| `npm run scenario` | Run a known fake week through the real digest. |
| `npm run loop` | The whole sort→correct→learn→promote loop, against a fake mailbox. |
| `npm run typecheck` | |

The split is the point: deterministic mailbox I/O stays in code, where it's cheap
and repeatable. Claude is asked to do only the part that actually needs a model.

---

## Known limits

- **Scheduled tasks need that user logged on.** They don't run on the lock screen as
  a service, because that would mean storing the account password. A missed run is
  picked up when the machine next wakes.
- **The MCP server needs Claude Desktop open**, on the machine it was installed on,
  and a full quit-and-reopen after connecting — closing the window is not enough.
  It cannot be reached from a phone or from claude.ai.
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
| Ask questions in Claude | No | Yes, via the MCP server |
| Stored credential | None at all | A 90-day refresh token |
| UI in Outlook | Full taskpane | Labels only |
| Edit categories in-app | Yes | No — edit the source |
| Classifier | Gemini, from the browser | Claude, on a schedule |

**Run one or the other, not both.** They share the provenance stamp on each
message, so corrections stay detectable either way, but they keep separate sender
rules and would both try to write `Inbox Steward:` rules.

## Not yet verified against a live mailbox

Everything buildable without her mailbox is built and tested. These are the things
only the real thing can settle, in rough order of how likely they are to bite:

1. **The PowerShell scripts have never executed.** `install.ps1` and
   `install-tasks.ps1` were written on a Mac. They handle the two known Windows
   traps (PowerShell 5.1 negotiating TLS 1.0, and `Invoke-WebRequest`'s progress bar)
   but have not been run. If `install.ps1` fails, fall back to the manual route:
   download the zip, `npm install`, `npm run setup` — that path is tested.
2. **The Claude Code CLI invocation.** `npm run sort` pipes the prompt as content and
   passes a fixed query, matching the documented `cat file | claude -p "query"` form.
   Run `npm run sort` by hand once before trusting the schedule, and confirm it
   prints a summary. It exits non-zero if the CLI produces nothing, so a silent
   no-op cannot masquerade as a successful week.
3. **Graph behaviours taken from documentation**, not observed: creating a folder
   with `isHidden`, listing it back with `includeHiddenFolders=true`, filtering the
   state message by `subject eq`, and `Prefer: outlook.timezone` on `calendarView`.
   If any is wrong, `npm run plan` says so on the first run.
4. **Whether the tenant lets a non-admin register an app.** Some don't. The wizard
   says so at that step.

`npm run plan` is the five-minute version of all of it, and it writes nothing.

## What a real inbox taught it

The synthetic tests above are written by the same person who wrote the code, which
limits what they can find. So the sender filter was also audited against a real
60-day inbox — 45 distinct senders, metadata only.

It caught **13 of 45**. Precision was perfect: all three actual humans in that inbox
were correctly left alone. But recall was 31%, and every missed bulk sender is a line
in the "waiting on a reply" list, which is exactly how that list stops being read.

The cause was a pattern no invented fixture had: **the local part is now the brand,
and the sending platform is in a subdomain.** `venmo@email.venmo.com`,
`news@em.oakley.com`, `ebay@reply.ebay.com`, `offers@promos.discounttire.com`,
`alerts@notify.wellsfargo.com`, `Microsoft365@infomails.microsoft.com`. Anchoring on
the local part cannot see any of it.

Adding a bulk-subdomain rule took it to **25 of 45**, with all three humans still
untouched. It stops there deliberately: `mail`, `email`, `e`, `m`, `t`, `my`, `go`,
`service` and `communication` were all present in that inbox as marketing
subdomains, and are all excluded from the rule anyway — `mail.some-university.edu`
is a real shape for a real person, and dropping one person from the waiting list is
worse than leaving several newsletters in it.

The remaining misses are consumer retail mail, which barely features in a foundation
mailbox. If some sender does keep appearing, add it to `STEWARD_IGNORED_SENDERS`
rather than waiting for the defaults to improve.

### Auditing the mailbox this is actually for

That audit was of a *different* inbox, so it proves the method rather than the
settings. Run the same check against the real mailbox, on the machine where it's
installed:

```bash
npm run audit          # last 60 days
npm run audit 120      # or a wider window
```

It prints every distinct sender split into "treated as bulk" and "treated as a
person", with counts, and marks which ones actually reach her To line — because a
sender who never does is already excluded by the To-line rule and can be ignored.
Then it proposes a `STEWARD_IGNORED_SENDERS` line for the high-volume senders whose
addresses look machine-generated.

**Addresses and counts only — no subjects, no message content, and nothing leaves the
machine.** It's read-only.

Two things about that proposed line. It lists **candidates, not conclusions**: some
will be people whose address merely looks automated, and `accounting@her-cpa.com` is
the obvious shape. And it emits full addresses rather than domain patterns on
purpose — a domain pattern covers every address a sender rotates through, but it also
silences everyone else at that domain. Broaden one to `@their-domain.com` yourself,
once you're sure nobody there writes to her personally.

Prune it before pasting. Excluding a real person removes them from the waiting list
silently, which is the one failure mode here with no symptom.

## Two bugs these tests found

Worth recording, because both were invisible and both would have quietly degraded
the thing she relies on.

**Two categories claimed the same mail.** The Events description listed "Tips &
Tricks or Year-End Ask workshop sessions" while Nonprofit Partners claimed
"professional development workshop registration and questions". A nonprofit
registering for Tips & Tricks matched both, the confidence split near-evenly, and
the gate correctly refused to guess — so it landed in ⚠ Needs Review every time,
forever, with no way to tell from the outside that the *taxonomy* was at fault
rather than the model. Events now owns arranging the room; Partners owns registering
for it, and each says so explicitly.

**A webinar invitation appeared as "waiting on a reply".** `webinars@candid.org`
wasn't in the ignored-sender list, so a Candid webinar promotion was reported as
something a human was waiting on her to answer. Exactly the kind of false positive
that gets a nag list ignored inside two weeks. Added, along with `webinar@`,
`invitations@`, and `events-noreply@` — while deliberately still leaving `info@` and
`events@` alone, since at a small nonprofit those are usually a person.

## License

MIT.
