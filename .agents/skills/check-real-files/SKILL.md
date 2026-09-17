---
name: check-real-files
description: Run the OOXML engine over real Office documents and triage what it reports — separating genuine faults from false positives, from observations that were never faults at all. Use whenever new .docx/.xlsx/.pptx fixtures are available, or after adding an analyzer that has never seen real Office output.
---

# check-real-files

Runs the engine over documents Office actually produced, and decides what each finding
means. **This is the highest-yield check in the repo.** Every other test uses XML written
by the same person who wrote the code reading it, so a false positive on a genuine
document is invisible to all of them.

The first session that did this found **43 findings across three files and not one was a
real fault** — two were serious bugs, one was a report that lied, and the rest were
observations dressed as defects. Every step below exists because of something that
happened in that session. Read `docs/ooxml-expert-agent/RESEARCH-STATE.md` §8ag before
starting; it is the record of that run.

---

## 0. Confidentiality first, before any file is copied

Real fixtures are usually someone's actual work. Getting this wrong is worse than any
bug this skill can find.

```bash
grep -n "fixtures" .gitignore          # must exclude *.docx *.xlsx *.pptx (and *m variants)
git check-ignore -v tests/fixtures/<name>   # confirm per file, after copying
git status --porcelain                 # must show nothing for the binaries, ever
```

Then, throughout:

- **Never** paste document text, cell values, comment bodies or headings into output,
  commits or PR descriptions.
- Part paths (`word/document.xml`), finding codes, element names and counts are engine
  vocabulary and safe. Sheet names and formulas are borderline — use them only when
  diagnosing, never in a commit message.
- Offer to delete the fixtures when finished. They are not yours to keep.

## 1. Run it

```bash
cp <files> tests/fixtures/
npm run test:real
```

The contract: a healthy file must claim **no error and nothing marked `silent`**. Visible
findings (`silent: false`) are printed but do not fail — they point at something already
on the page, and whether it *matters* depends on intent the engine cannot know. Do not
"fix" a visible finding just to get a green run.

Name a deliberately-broken fixture `*.expect-findings.docx` to invert the assertion.

## 2. ⚠️ Suspect your own tooling before you suspect the data

**In the session that produced this skill, the diagnostic tooling was wrong three separate
times, each time producing confident and entirely false results.** Bare-name XSD type
resolution invented 39 divergences. Unexpanded `xsd:group` references made agreement read
70% instead of 90%. A misattributed namespace prefix turned `m:oMath` into `w:oMath`.
Separately, three throwaway probes were wrong — a missing `await`, a wrong argument shape,
a function name that did not exist.

So, before reporting any finding as a false positive:

1. Read the raw markup and confirm the engine's claim is actually false.
2. Check your probe's own assumptions — is the helper async? does it take a map or a
   parsed document? does that export exist?
3. Only then conclude the engine is wrong.

A probe that needs a DOM must run under vitest (jsdom), not plain node.

## 3. Triage every finding into exactly one of four buckets

This is the whole job. Be decisive and write the reason down.

**(a) A real fault.** The document genuinely is broken. Excellent — the engine worked.
Verify it independently, then say so plainly.

**(b) A false positive: the engine's claim is factually untrue.** Fix the engine. Two
found this way, both worth recognising by shape:

- **First-match over a list.** `matching(parts, WORD_BODY)[0]` read whichever body part
  the archive listed first — `word/footnotes.xml` in a real document — so comments
  anchored in the body looked orphaned *and* genuine faults there were missed. Every
  hand-written fixture has exactly one body part, so `[0]` was always right. **Grep for
  `[0]` and `.find(` over part lists whenever a real file disagrees.**
- **An assumption true for one format only.** `cache-is-only-source` assumed chart data
  must live in an embedded workbook — true for Word and PowerPoint, false inside a
  spreadsheet, where formulas resolve against the host workbook. Ask: *is this claim
  format-specific, and is it being applied to all three?*

Where possible, replace the wrong check with the stronger question it was reaching for,
rather than just silencing it. "No embedded workbook" became "do the sheets it names
exist?".

**(c) Not a fault at all.** The engine is stating something true and irrelevant. Remove it
from the findings channel. Tells:

| Tell | Example from the real run |
|---|---|
| Fires on nearly every healthy instance | `chart/translation-risk` — 31 of 43 findings |
| Its own remediation says "No action needed" | `contentControl/unbound-control` |
| It repeats the analyzer's own `cannotDetermine` | `comment/threading-unknown` |
| It is conditional on intent the engine cannot know | placeholder text in a template |

Before removing, check whether the information already reaches the user through the
`explain` path. `translationNotes` were already surfaced under "Decide these before
converting", so deleting the findings lost nothing.

**(d) Correct for most cases, wrong for one type.** Narrow it. `field/no-cached-result`
fired on `FORMCHECKBOX`, which has no result **by design** — its state is in `w:ffData`,
and real output shows `begin=1, separate=0`. Excluding form-field types kept the check and
removed the lie.

## 4. Check `silent` on anything you touch

`silent: true` means *"renders correctly and is broken anyway"* — a claim the reader cannot
check by looking. Asserted about a healthy file it is **the most damaging thing this engine
can say**. A clean deck once reported that 12 things rendered correctly and were broken
anyway, when nothing was broken at all.

If a finding is not describing invisible breakage, it is either `silent: false` or it is
not a finding.

## 5. Pin every fix with a test that runs in CI

The fixtures cannot be committed, so a fix proven only by them is unproven the moment the
files are deleted. Write hand-made XML reproducing the real structure — for the comment
bug, a package whose key order puts `word/footnotes.xml` before `word/document.xml`.

**⚠️ Write the test at the layer the bug actually lived.** The first regression test for
the comment bug **survived reverting the fix**: it called `readComments` directly with
explicit stories, so it proved that function handles multiple stories and said nothing
about the analyzer's part *selection*, which is where the bug was. The test that catches it
goes through `analyzePackage`.

Then mutation-test, per `RESEARCH-STATE.md` and `.agents/skills/add-analyzer/mutate.py`:
revert the fix and confirm the new test fails. **A regression test that survives reverting
the fix is not a regression test.** Also revert in the *other* direction — confirm the
check still fires on a genuinely broken input, so the fix did not simply disable it.

## 6. Report what the files did NOT prove

A green run is not coverage. Tally which analyzers actually ran:

```ts
const run = analyzePackage(parts);
const ledger = capabilityLedger(run);
// ledger.ran / ledger.skipped
```

Note that explain-only analyzers (Word cascade, Excel formats) are reported as *skipped*
on purpose — they apply to the package but contribute no faults, and counting them as
"ran" would overstate what validation checked. That is not a bug.

Then state the gap in named terms. After three files: *"18 of 22 analyzers ran; `equation`
and `conformance` never fired, because no file had an equation and all three were
Transitional."* That sentence is more useful than any pass rate, and it tells the user
which single file to look for next.

## 7. Finish

- Full gate: `npx tsc --noEmit`, `npx eslint .`, `npx vitest run`, `npm run build`.
- Confirm no binary is staged: `git status --porcelain | grep -iE "docx|xlsx|pptx"`.
- Add a `RESEARCH-STATE.md` section: which files (by type, not name, if sensitive), how
  many findings, the bucket each fell into, and what went unexercised.
- Offer to delete the fixtures.

---

## The one-line version

Real files find bugs that hand-written fixtures structurally cannot — but **most of what
they report is not a bug at all**, and the skill is telling those apart without either
trusting the engine or gutting it to get a green run.
