# Day one

The install, in order, with what to do when a step misbehaves. Everything else is
in [README.md](README.md); this is the page to have open at her desk.

Budget 30 minutes. She needs to be there for step 3 only, but it's worth her seeing
step 5.

---

## Before you go

Three things, because each one can end the visit early:

- **Does her machine have Node 20+?** `node --version`. If not, install LTS from
  nodejs.org — **this may need an admin password you don't have.** Find out now, not
  at her desk.
- **Is Claude Desktop installed on that machine?** Only matters for the ask-anything
  half, but that's the half she'll use most.
- **Can a non-admin register an app in her Microsoft tenant?** Some tenants forbid
  it. If hers does, whoever administers the foundation's Microsoft 365 has to do one
  step, and no amount of preparation gets around it.

Optional but recommended: do a full dry run against your own mailbox first, or a
free [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program)
tenant. Then nothing at her desk is happening for the first time.

---

## At her desk

### 1. Install

PowerShell:

```
irm https://raw.githubusercontent.com/aaronv02/claude-routine-outlook-sorter/main/install.ps1 | iex
```

Downloads to `C:\outlook-sorter`, installs dependencies, starts the wizard.

> **If it fails:** download
> [the zip](https://github.com/aaronv02/claude-routine-outlook-sorter/archive/refs/heads/main.zip),
> extract to `C:\outlook-sorter`, then `npm install` and `npm run setup`. That route
> is fully tested; the PowerShell one has never run on a real Windows machine.
>
> Don't extract to Documents or Desktop — those are usually OneDrive-synced, and
> `node_modules` there causes sync churn and file-lock failures.

### 2. Register the app — the only manual part

The wizard prints four numbered clicks. The one people miss is
**Allow public client flows → Yes**, on the Authentication page.

> **If you can't create an app registration at all:** the tenant blocks it. Stop
> here and get the administrator.

### 3. She signs in

The wizard shows a URL and a code. **She** enters it — whoever signs in is whose
mail gets sorted.

Afterwards the wizard checks that all five permissions were actually granted, and
names any that weren't. Sign-in succeeding does not mean the permissions are there.

> **If sign-in won't start:** the error names the likely cause. `AADSTS50059` means
> the tenant ID is wrong; anything else at that stage is almost always the
> public-client setting from step 2.

### 4. Let it prepare the mailbox

It creates the category labels, then reads how she's already been filing mail.
**Watch the "learned N senders" number** — that's how much it starts out knowing. If
it's 0, her mailbox has little existing categorization, which is fine but means the
first week leans on the model more.

### 5. Look at one real sweep

The wizard runs `npm run plan`, which reads her inbox and writes what it *would*
label. It writes nothing to the mailbox.

**Open `routine\.local\plan.json` and read the `pending` list.** This is the moment
to catch a wrong assumption — if those look like real messages with sensible senders,
the plumbing works. Then:

```bash
npm run apply
```

Check Outlook. Labels should be on those messages, and any categories of her own
should still be there alongside them.

### 6. Check the sender filter against her mail

```bash
npm run audit
```

Read-only, addresses and counts only. It lists who it treats as bulk versus a
person, and proposes a `STEWARD_IGNORED_SENDERS` line.

**Prune that line before pasting it.** It lists candidates, not conclusions —
`accounting@her-cpa.com` looks automated and is a person. Excluding a real person
removes them from the weekly waiting list with no symptom.

### 7. Prove the scheduled path once

```bash
npm run sort
```

This is the whole hourly routine, run by hand. It should print a short summary. It
exits non-zero if the Claude Code CLI produced nothing, so a silent no-op can't
masquerade as success.

> **If it can't find `claude`:** `npm install -g @anthropic-ai/claude-code`, then run
> `claude` once in that folder to sign in. The scheduled tasks need that sign-in.

### 8. Turn it on

```
.\scripts\install-tasks.ps1
```

Hourly sorting 8:00–18:00, summary Fridays at 15:00. Add `-StartHour 7 -EndHour 19
-SummaryTime 16:30` to change that.

### 9. Connect Claude Desktop, if it's on that machine

The wizard offers this at step 7, or `npm run connect` any time. Then **fully quit
Claude Desktop from the system tray and reopen it** — closing the window is not
enough, and this is the step that makes people think the install failed.

Confirm by asking Claude: *"who is waiting on me?"*

---

## Before you leave

Give her [HANDOFF.md](HANDOFF.md). It's one page and the only thing she needs.

Say the one sentence that matters: **if a label is wrong, change it in Outlook the
way you normally would.** That's the whole interface.

---

## What will break, and when

- **In 90 days**, or immediately if her password changes, the sign-in expires. The
  routine fails with `invalid_grant` and someone runs `npm run setup` again. That
  someone is you — she can't fix it.
- **If the folder moves or is deleted**, Claude Desktop loses the MCP server, because
  the config stores absolute paths.
- **Before 31 December 2026**, check whether the new `Mail-Advanced.ReadWrite`
  permission covers `categories`. The evidence says it doesn't, but that came from a
  secondary source.

## If something is wrong later

Logs are in `routine\.local\logs\`. `npm run sort` and `npm run summary` reproduce
what a scheduled run does, on demand, and print the same thing.

`npm run plan` is always safe — it writes nothing.
