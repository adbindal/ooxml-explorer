# An OOXML analysis engine, and an AI panel that uses it

Turns Explorer's AI from a persona prompt over a reference corpus into a **deterministic
analysis engine** that reads the open document and computes its answers. The model
narrates a result it did not decide.

**98 commits.** `main` has not moved, so this is a clean fast-forward.

---

## Why

On `main` today, `geminiService.ts` asks the model to *"run a strict validation check for
ECMA-376 compliance… citing exact issues."* We asked a language model to be a validator
and to produce its own citations — the one thing it will confidently invent. Nothing in
the UI distinguished a real citation from a fabricated one.

The corpus behind it has **1,899 records and 29 with human-written meaning** — the rest is
SDK-derived structure. So for ~98% of questions there was nothing to ground an answer in,
and the model filled the gap fluently.

More corpus, re-ranking, embeddings and fine-tuning were each considered and are each
argued in `docs/ooxml-expert-agent/RESEARCH-STATE.md`. The short version: the question was
never *"what does `w:b` mean"*, it was *"why isn't **this** paragraph bold"*, and no
amount of reference data answers a question about a specific file.

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

**An evidence tier** computed from provenance in code the model never touches, taking the
minimum across sources. A clean validation run reports *"no problems found by the checks
that ran"* — never *"this file is correct."*

---

## Suggested review order

This is large. It is not meaningfully splittable *now* — the analyzers all depend on
`findings.ts` and the registry — but it can be read in layers, and the first three files
carry the design:

| Read | Why |
|---|---|
| `services/findings.ts` | The type everything else emits. ~150 lines. |
| `services/analyzers.ts` | The registry, routing, and capability ledger. |
| `services/wordBookmarks.ts` | A representative analyzer, and the smallest. |
| `docs/ooxml-expert-agent/RESEARCH-STATE.md` §8u | Why these 21 and not others. |

The remaining analyzers are the same shape repeated. **Skimming two and trusting the
tests is a reasonable review strategy** — each has a module doc-comment explaining what
silently breaks, and a rules table making severity explicit.

### Diff composition

| Area | Lines | Note |
|---|---:|---|
| `public/rag-data.json` | 29,181 | **generated** by `scripts/ingestSchema.ts` — skim, don't read |
| `services/` | 17,273 | the engine |
| `tests/` | 13,683 | ~0.8 test lines per source line |
| `docs/` | 1,180 | design record and licensing research |

---

## Verification

```
npx tsc --noEmit    ✓
npx eslint .        ✓
npx vitest run      ✓  1,384 passing, 1 skipped
npm run build       ✓
```

Every module was **mutation-tested** — the implementation deliberately broken several ways
to check the tests actually catch it. That found a real gap in nearly every module, and
the recurring cause was a test passing for the wrong reason rather than missing coverage.
`.agents/skills/add-analyzer/mutate.py` is the harness.

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

## Risk

`main` is untouched and this is additive — existing panel behaviour is preserved except
where it was the thing being replaced.

The behavioural change users will notice is the **evidence tier**: answers that
previously appeared with no qualification now carry a badge, and some will read as
`Unverified`. That is the same answer as before, honestly labelled — but it will look
like a regression to anyone who read confidence into the absence of a caveat.

DLP already fails closed on `main` (`decbd66`); this PR does not touch
`services/aiProvider.ts`.
