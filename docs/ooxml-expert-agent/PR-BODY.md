# An OOXML analysis engine, and an AI panel that uses it

Turns Explorer's AI from a persona prompt over a reference corpus into a **deterministic
analysis engine** that reads the open document and computes its answers. The model
narrates a result it did not decide.

**109 commits.** `main` has not moved, so this is a clean fast-forward.

> **The "cannot be rebased" banner is a false alarm.** There is no conflict. `main` has not
> moved — merge-base *is* `origin/main`'s tip — and a local rebase replays all 105 non-merge
> commits cleanly. The branch carries **4 merge commits** from parallel work, and GitHub's
> *Rebase and merge* replays commits one at a time and refuses when merges are present.
> **Squash and merge** and **Merge commit** are both clean and enabled.

---

## The one idea to take away: this app now has two engines

They answer different questions, and neither can do the other's job.

| | Answers | Reads | Same for every file? |
|---|---|---|---|
| **The dictionary** (`rag-data.json`) | *"What is this thing?"* — `w:b` is Bold, it carries `w:val`, it may only sit inside `w:rPr` | Microsoft's published schema | Yes |
| **The compiler** (this PR) | *"Why does my file look wrong?"* — this heading is not bold because Heading1 borrows from Normal, which turns bold off | The open document | No — it is about *your* file |

Like a car manual and a mechanic. The manual cannot tell you your brake pad is worn.
Looking at your car does not teach you what a brake pad is.

**They do not talk to each other.** No analyzer imports the dictionary. Both hand their
results to the AI panel separately, and the two meet in exactly one line
(`selectEvidenceTier` in `services/aiService.ts`). A bigger dictionary would not make the
compiler find one extra fault.

That is the whole reason this PR exists. Growing the dictionary 65× — 29 records to
1,899 — did not answer a single *"why is my file wrong"* question.

---

## Why not just improve retrieval?

Each alternative was considered and rejected with a reason, recorded in
`RESEARCH-STATE.md` §5. Re-propose only with a new argument.

- **Fine-tuning** — a fine-tuned fact has **no source span**, so it cannot be cited or
  checked. That structurally destroys the honesty property the badge depends on.
- **Vector database** — a few thousand structured records whose primary access is an
  exact key lookup. A vector index solves a problem this data does not have.
- **Re-ranking** — presupposes multiple candidates; the exact lookup returns 0 or 1.
- **Embeddings on the main path** — same reason, and they make a retrieved span harder to
  cite, not easier.

The question was never *"what does `w:b` mean"*. It was *"why isn't **this** paragraph
bold"*, and no amount of reference data answers a question about a specific file.

---

## What this adds

**21 analyzers** across Word, Excel and PowerPoint, each detecting a fault that **renders
correctly and is broken anyway** — a dropped OLE embedding behind an intact preview, a
cross-reference to a deleted bookmark still displaying its cached text, a formula whose
cache no longer matches it.

**One `Finding` type** (`services/findings.ts`) that every analyzer emits — code,
severity, part, message, remediation, and `silent`. Prose is *rendered* from findings at
the boundary, never authored as findings, which is what makes the output consumable by
CI or another agent.

**An analyzer registry** (`services/analyzers.ts`) that routes questions and computes a
**capability ledger** — what ran, what was skipped, and what the checks that ran
explicitly *cannot* establish.

**An evidence tier** computed from provenance in code the model never touches. A clean
validation run reports *"no problems found by the checks that ran"* — never *"this file
is correct."*

---

## Three defects found while preparing this PR

Each was caught by a different mechanism, which is worth noting because it says something
about which checks actually earn their place.

**1. A live crash in natural-language search** — caught by CI.
`storageService.ts` called `doc.definition.toLowerCase()` unguarded. `definition` is
optional, and deliberately absent on schema-derived records. That was safe while the
corpus was 29 curated entries, all with prose. This branch takes it to 1,899 records of
which **1,870 have no definition**, so the first one the cursor reached threw — uncaught,
killing the whole search path. Any user typing a question would have hit it. It survived
because the only coverage was **mocks that reimplemented the predicate** rather than
calling it, so the tests agreed with themselves and never touched the real code.

**2. `issueReport.ts` had zero importers** — caught by a reachability sweep.
163 lines plus a 132-line test, built for in-app bug reporting, superseded by the GitHub
issue templates, never removed. The same trap as `ooxmlDiff.ts`, which sat unwired for
489 lines earlier in this branch.

**3. Finding a citation made the badge worse** — caught by reading the tier rule.
`selectEvidenceTier` took the minimum across sources, so a dictionary **miss** plus a
complete computation scored `verified`, while a **hit** plus the same computation scored
`grounded`. Knowing more reduced confidence. The minimum is the right rule for sources
that can be *wrong*; these two cannot be, and they are about different subjects. The
computation now governs. ⚠️ Safe **only** because both inputs are mechanically derived —
a source that can be wrong must go back to a minimum.

---

## Suggested review order

This is large. It is not meaningfully splittable *now* — the analyzers all depend on
`findings.ts` and the registry — but it can be read in layers, and the first three files
carry the design:

| Read | Why |
|---|---|
| `services/findings.ts` | The type everything else emits. ~150 lines. |
| `services/analyzers.ts` | The registry, routing, and capability ledger. |
| `services/wordNotes.ts` | A whole analyzer in 256 lines, including the trap that would fire on every real document. |
| `docs/ooxml-expert-agent/RESEARCH-STATE.md` §8u | Why these 21 and not others. |

Nothing in `services/` is unreachable from the app entry point, and no export is left
without a consumer. Where a helper had no production caller but a test needed it to
observe live behaviour, it moved into the test file rather than staying exported for the
suite's benefit.

The remaining analyzers are the same shape repeated. **Skimming two and trusting the
tests is a reasonable review strategy** — each has a module doc-comment explaining what
silently breaks, and a rules table making severity explicit.

### Diff composition

| Area | Lines | Note |
|---|---:|---|
| `public/rag-data.json` | 29,181 | **generated** by `scripts/ingestSchema.ts` — skim, don't read |
| `services/` | 17,279 | the engine |
| `tests/` | 13,849 | ~0.8 test lines per source line |
| `docs/` | 1,333 | design record and licensing research |

---

## Verification

```
npx tsc --noEmit    ✓
npx eslint .        ✓
npx vitest run      ✓  1,386 passing, 1 skipped
npm run build       ✓
npx playwright test ✓  7 passed
```

Every analyzer was **mutation-tested** — the implementation deliberately broken several
ways to check the tests actually catch it. That found a real gap in nearly every one, and
the recurring cause was **a test passing for the wrong reason** rather than missing
coverage. `.agents/skills/add-analyzer/mutate.py` is the harness.

The practice was adopted partway through, so the earliest modules — the Word cascade,
Excel cell formats, PowerPoint inheritance, package integrity and the semantic diff —
have tests but were never mutated. They are the least-proven code in the diff.

The one skipped test is the real-file suite, which skips when `tests/fixtures/` is empty
and **says so loudly** rather than passing silently.

---

## What is not proven

**The engine has barely met real Office output.** Every test uses hand-written XML,
written by the same people who wrote the code that reads it — so a false positive on a
genuine document is invisible to all of them.

`tests/fixtures/` plus `npm run test:real` exists for exactly this. The fixture binaries
are gitignored, so the directory is safe to point at confidential documents. **This is the
main thing to be sceptical of, and the reason nothing here claims to be validated against
real files.**

Two smaller ones, both recorded rather than hidden:

- The Word and Excel resolvers match exact Transitional namespace constants, so they read
  nothing from ISO Strict packages. `services/conformance.ts` normalises at the pipeline
  level, and the `conformance` analyzer reports the limit — but it is a mapping, not a
  conversion.
- `services/pptAnimation.ts` and `services/excelExternalLinks.ts` were written by agents
  that were interrupted, then finished and tested afterwards. Both had real defects on
  recovery (`Number(null) === 0` in both cases) which are fixed and pinned by tests.

---

## A known limitation in the dictionary, for later

The dictionary is **shallower than its own source**, and this looks like an oversight
rather than a decision. `scripts/ingestSchema.ts` declares each attribute as:

```ts
interface SdkAttribute {
  QName?: string;
  Validators?: {...}[];   // declared, never read
  Version?: string;       // declared, never read
}
```

Only `QName` is extracted. So for `w:jc` — paragraph alignment — the dictionary records
that it has an attribute `w:val`, but **not that `w:val` is one of
`left | center | right | both`**. That is the most useful fact about the element.
`Children` is read, but only to invert into `parents`, so we know what an element sits
inside and never what may sit in it.

The script's own header names those constraints as the reason for choosing this source.
Extracting them is a change to one script plus a regenerate, keeps everything
mechanically derived, and roughly doubles what a dictionary hit is worth. **Not in this
PR.**

---

## Where the effort goes next

1. **Extract `Validators`, `Version` and children** in the ingest script — highest value
   per unit of work, and it preserves the honesty property.
2. **More analyzers**, driven by the gap log rather than guesswork. This is the actual
   differentiator: no other tool tells a person *why their file is wrong*. DrawingML is
   the largest remaining blind spot (§10).
3. **Retrieval over specification prose** — a **third engine**, not a bigger dictionary.
   Only if the gap log shows *"how do I…"* questions are a real share of traffic, and
   with the licensing question answered first (`docs/ooxml-expert-agent/LICENSING.md`).
   Note that the hard part is not retrieval: standards prose is written for implementers,
   and dropping clauses into the panel would make answers longer and worse. Turning a
   clause into an actionable sentence is the work, and the `message` + `remediation`
   split the analyzers already use is the shape to reuse.

Deliberately **not** doing: importing the full specification. It is large enough to ship
to every user, mostly restates structure already held in a more reliable derived form,
raises a licensing question your own research does not settle, and — because prose is not
addressable by key — would drag back the embeddings and re-ranking ruled out above.

---

## Risk

`main` is untouched and this is additive — existing panel behaviour is preserved except
where it was the thing being replaced.

The behavioural change users will notice is the **evidence tier**: answers that
previously appeared with no qualification now carry a badge, and some will read as
`Unverified`. That is the same answer as before, honestly labelled — but it will look
like a regression to anyone who read confidence into the absence of a caveat.

DLP already fails closed on `main` (`decbd66`); this PR does not touch
`services/aiProvider.ts`.
