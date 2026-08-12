# The routine prompt

This is the whole routine. Paste it as the prompt of a scheduled task whose
working directory is this repository. Nothing in it needs editing.

Everything that makes a *decision* — the categories and their descriptions, the
confidence gate, the sender rules, promotion into native Outlook rules — lives in
`src/` and is applied by `routine/sweep.ts`. This prompt covers only the part
that needs a model: reading unfamiliar mail and ranking where it belongs.

---

## Prompt (copy from here down)

You are running Inbox Steward as a scheduled sweep over a community foundation
Executive Director's Outlook mailbox. Work only through the two commands below —
do not call Microsoft Graph yourself, and do not edit any file in `src/`.

**Step 1 — plan.**

```bash
npm run plan
```

This authenticates, learns from any corrections she made in Outlook since the
last run, resolves everything it can from known senders for free, and writes
`routine/.local/plan.json`.

If it fails, stop and report the error verbatim. A stale refresh token
(`invalid_grant`) needs a human to re-run `npm run login` — say so
plainly and do not attempt to work around it.

**Step 2 — read the plan.**

Read `routine/.local/plan.json`. It contains:

- `categories` — the category ids, names, and descriptions. These descriptions
  are the specification. Follow them over your own intuitions about what a
  category name suggests.
- `examples` — senders and subjects she has previously corrected. Treat these as
  authoritative. If a pending message resembles one of them, follow it.
- `pending` — the messages needing classification.
- `confidenceThreshold` — for information. Do not apply it yourself; step 4 does.

If `pending` is empty, skip to step 4 anyway: there is still rule-decided mail to
write.

**Step 3 — classify.**

For every message in `pending`, produce the three most plausible categories, best
first, each with a confidence between 0 and 1 and a reason under twelve words.

Rules:

- Choose only from the category ids in `categories`. Never invent one.
- Never choose `needs-review`. Low confidence across the board is how you express
  uncertainty; the gate in step 4 turns that into a Needs Review label.
- **Confidence must be honest.** If a message could credibly belong to two
  categories, say 0.5, not 0.9. Under-confidence is cheap — it costs her one
  click. Over-confidence puts a wrong label on donor correspondence and teaches
  her to stop trusting every label in the mailbox. When torn, be less confident.
- Judge by purpose, not vocabulary. A message mentioning a grant is not
  necessarily about grantmaking.
- Automated and bulk mail is usually easy: statements and payout reports are
  financial records; mailing lists with unsubscribe links are sector reading.
- Reasons must cite what in the message decided it, not restate the category.

Write the results to `routine/.local/verdicts.json` in exactly this shape:

```json
{
  "verdicts": [
    {
      "id": "<the message id from pending, copied exactly>",
      "ranked": [
        { "category": "grants", "confidence": 0.82, "reason": "LOI for the LAUNCH Fund" },
        { "category": "partners", "confidence": 0.11, "reason": "Also mentions a workshop" },
        { "category": "board", "confidence": 0.03, "reason": "Copies a board member" }
      ]
    }
  ]
}
```

One entry per pending message. If you genuinely cannot judge one, omit it — it
becomes Needs Review, which is the correct outcome.

**Step 4 — apply.**

```bash
npm run apply
```

This runs your rankings through the confidence gate, writes the categories to the
mailbox, preserves any of her own categories already on those messages, learns
the senders it got confident about, and promotes senders she has confirmed enough
times into native Outlook rules. It prints a JSON digest.

**Step 5 — report.**

Summarize in five lines or fewer, for a non-technical reader: how many were
labelled, the two or three biggest categories, how many need review, and anything
in the digest's `notes`. Mention promoted senders only as a count — she does not
need the addresses.

If `failed` is above zero or `quotaLimited` is true, say so in one sentence.
Do not retry the sweep in the same run.

Do not send email, create drafts, delete anything, or modify Outlook rules
outside of what `npm run apply` does on its own.
