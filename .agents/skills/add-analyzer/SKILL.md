---
name: add-analyzer
description: Add a new OOXML analyzer to the engine — decide whether it is worth building, verify the schema, write it in the house style, mutation-test it, register it, and document it. Use when adding a check for a new OOXML feature, or when the coverage gap log shows something uncovered.
---

# add-analyzer

Adds one analyzer to the OOXML engine. The procedure below is not ceremony — every step
exists because skipping it produced a specific, real defect in this codebase.

**Read `docs/ooxml-expert-agent/RESEARCH-STATE.md` §8u first** (the selection criterion)
and skim §8m–§8ad (what each existing analyzer found). Then follow this in order.

---

## 0. Gather the input

Ask the user for whichever of these they have. **None is required to proceed** — say what
is missing rather than blocking.

- **The feature**: "footnotes", "Excel defined names", "SmartArt".
- **Gap-log output**: the Validator's *Retrieval Stats* button prints coverage gaps —
  normalised part paths and element names, most-requested first.
- **A real file** exhibiting the problem, if they have one.

---

## 1. Decide whether to build it at all

**This is the most valuable step and the easiest to skip. A legitimate outcome of this
skill is "do not build this."**

### The test: does this feature fail INVISIBLY?

Every analyzer that earned its keep detects something a person looking at the document
cannot see. A feature whose faults are *visible* does not need one — a wrong shape looks
wrong and the user reports it without help.

Three signals predict invisible failure:

1. **Indirection** — resolves through `r:id` or an implicit relationship. Break the link
   and a fallback renders. *(OLE, media, images, pivots, charts.)*
2. **Paired markers** — two elements matched by id where losing one is silent.
   *(Bookmarks, comments, revisions, footnotes.)*
3. **Cached or duplicated state** — a stored copy that can drift from its source.
   *(Formula values, field results, pivot caches, chart series, external links.)*

If it hits none of these, it probably wants better *explanation*, not an analyzer.

### Check how much Office and the spec disagree

Deviation counts are a real signal of where "valid per spec" ≠ "renders right in Office":

```bash
curl -sL "https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/toc.json" -o /tmp/toc.json
python3 - <<'PY'
import json, re, collections
d = json.load(open('/tmp/toc.json'))
titles = []
def walk(n):
    if isinstance(n, dict):
        if n.get('toc_title'): titles.append(n['toc_title'])
        for c in (n.get('children') or []): walk(c)
    elif isinstance(n, list):
        for c in n: walk(c)
walk(d['items'])
pat = re.compile(r'^\d+(?:\.\d+)*\s+Part\s+(\d+)\s+Section\s+([\d.]+?),')
c = collections.Counter()
for t in titles:
    m = pat.match(t)
    if m: c[(m.group(1), '.'.join(m.group(2).rstrip('.').split('.')[:2]))] += 1
for k, v in c.most_common(25): print(f"Part {k[0]} §{k[1]:<8} {v}")
PY
```

Only headings and clause references are read here, never Microsoft's prose — see
`docs/ooxml-expert-agent/LICENSING.md` on why that distinction matters.

**Cross the two.** High deviations + invisible failure = build now. High deviations +
visible failure = skip (this is why DrawingML, math and VML were skipped despite topping
the table). **Report the recommendation to the user before writing code.**

---

## 2. Verify the schema BEFORE writing anything

**Never trust a brief, including one written by the user or by you. Four briefs on this
project asserted element paths that do not exist, and every one was caught here.**

```bash
# list every schema file
curl -s "https://api.github.com/repos/dotnet/Open-XML-SDK/contents/data/schemas" \
  | python3 -c "import json,sys; [print(f['name']) for f in json.load(sys.stdin)]"

# fetch one and inspect a type
curl -sL "https://raw.githubusercontent.com/dotnet/Open-XML-SDK/main/data/schemas/schemas_openxmlformats_org_wordprocessingml_2006_main.json" -o /tmp/wml.json
python3 - <<'PY'
import json
d = json.load(open('/tmp/wml.json'))
for t in d['Types']:
    if 'CT_YourType' in t.get('Name', ''):
        print('---', t['Name'], '| base:', t.get('BaseClass'))
        for a in t.get('Attributes', []):
            v = a.get('Validators') or []
            req = 'REQUIRED' if any(x.get('Name') == 'RequiredValidator' for x in v) else ''
            print('   ', a.get('QName'), a.get('Type'), req)
        print('    children:', [c.get('Name','').split('/')[-1] for c in (t.get('Children') or [])])
for e in d.get('Enums') or []:
    if 'YourEnum' in e.get('Name',''): print(e['Name'], [f['Name'] for f in e.get('Facets',[])])
PY
```

`data/namespaces.json` maps prefixes to URIs. **Write down what you verified and what you
did not** — the module header must distinguish the two, and `cannotDetermine` must list
the limits.

---

## 3. Write the module

`services/<name>.ts`. Copy the shape of the closest existing analyzer — `wordBookmarks.ts`
for paired markers, `oleObjects.ts` for relationship chains, `excelFormulas.ts` for cached
state.

Required elements:

- **A module doc-comment that explains what silently breaks.** Not what the elements are —
  what a person would get wrong. Lead with the failure.
- **A rules table** mapping each problem kind to `{ severity, silent }`, with the severity
  reasoning in a comment. `silent: true` means *the document renders correctly and is
  broken anyway* — the most valuable bit in the record.
- **Findings via `finding()`** from `services/findings.ts`, with a namespaced code
  (`yourAnalyzer/kind`). Namespacing is not cosmetic: `duplicate-id` is already used by
  two analyzers.
- **`compute<X>EvidenceForMarkup(parts, rawXml)`** returning `{ lines, unresolved } | null`
  for the panel. Put every limit in `unresolved` — it caps the evidence tier.
- **Tolerate malformed input.** Return nothing; never throw.
- Namespaces arrive normalised to Transitional (`services/conformance.ts`), so compare by
  exact equality against the shared constants.

### 🔴 The false-positive traps

**The highest-value work in these analyzers is knowing what NOT to report.** Each of these
would fire on a large fraction of real documents:

| Trap | Why it fires |
|---|---|
| `w:numId="0"` | Means *remove numbering*, not a lookup. Fires on every doc that ever had a list removed. |
| Footnote separators | `separator`/`continuationSeparator` notes are referenced from `sectPr`, never the body. **100% of real files.** |
| Excel `numFmtId` < 164 | Built into Excel, declared nowhere. Fires on every workbook. |
| `HYPERLINK` without `\l` | The argument is a URL, not a bookmark. |
| `STYLEREF` | The argument is a style name, not a bookmark. |
| Missing `w:fldChar` `separate` | A field that was never calculated. Legal and common. |
| A shared-formula master | Has both `ref` and `si`; testing only `si` flags every master. |
| ISO Strict namespaces | Strict **drops the year** (`/spreadsheetml/main`, not `/2006/main`). |

Two JavaScript traps that produced real defects here:

- **`Number(null)` is `0`** — a missing required attribute silently became index 0 and
  produced confident findings about a range the markup never stated.
- **`undefined !== null` is `true`** — `obj?.field !== null` is true when `obj` is null.
  Coalesce before comparing.

---

## 4. Test, then break it

Write tests in the house style: in-memory XML strings, no binary fixtures, comments saying
*what silently breaks* without each behaviour.

Then **mutation-test — mandatory, not optional.** Every module here passed its tests on the
first run, and mutation testing still found a genuine gap in nearly all of them.

```bash
python3 .agents/skills/add-analyzer/mutate.py \
  --source services/yourAnalyzer.ts \
  --test tests/yourAnalyzer.test.ts \
  --mutations /tmp/mutations.json
```

`mutations.json` is a list of `{"label": "...", "old": "...", "new": "..."}`. See the
script's `--help` for the format.

### When a mutant survives, suspect the TEST first

On this project that has been the answer more often than not:

- **A fixture that would pass either way** — a sort test whose input was already ordered;
  an off-by-one test using a value far from the boundary.
- **An assertion decided by an earlier branch** — an external-link test whose verdict came
  from a "nothing was checked" path, not the rule it targeted.
- **An accidental tie-break** — an ordering test where alphabetical order happened to
  agree with the expected order.
- **A guard that is provably unreachable** — then *delete the guard*, do not write a test
  that pretends to cover it.

If a mutant is genuinely equivalent, **say so in the report instead of inventing a test.**

---

## 5. Check against real files — gated, and honest about it

There is a harness for this: **`npm run test:real`** runs every analyzer over every file
in `tests/fixtures/` and fails if a known-good file produces a finding. The binaries are
gitignored, so that directory is safe to point at confidential work documents.

`npm run fixtures:smoke` writes two generated OPC packages (one valid, one deliberately
broken) if you need a baseline with no access to Office. **They prove the harness, not the
analyzer** — they are written by this repo and share its blind spots.

```bash
npm run test:real
```

- **If files exist**: run the analyzer over each and read every finding. A finding on a
  known-good file is a false positive and must be fixed before shipping — that is the
  whole point of this step.
- **If none exist**: say so explicitly in the report and in the commit message. Record the
  analyzer as **fixture-verified only**. Do not describe it as validated against real
  files; that is the difference between *correct* and *trustworthy*.

**A false positive here outranks a missing check.** An analyzer that fires on a document
Office is happy with makes the whole report untrustworthy, and one bad finding discredits
twenty good ones. If a real file trips your analyzer, fix the analyzer before shipping —
do not add an exception to the fixture.

---

## 6. Register it

In `services/analyzers.ts`, add an entry with:

- `id` matching the code namespace
- `formats`, `appliesTo`
- `analyze` (finds faults) and/or `explain` (describes resolution) — an analyzer may do
  either or both
- `determines` — the questions it answers
- **`cannotDetermine` — what it explicitly cannot establish.** This feeds the capability
  ledger and is the honest half. **An analyzer that declares no limits is almost
  certainly lying**, and a test enforces that this array is non-empty.

---

## 7. Verify, document, commit

```bash
npx tsc --noEmit && npx eslint . && npx vitest run && npm run build
```

All tests must pass, not just the new ones.

Add a section to `docs/ooxml-expert-agent/RESEARCH-STATE.md` covering: what silently
breaks, what was verified versus assumed, the traps encoded, and the mutation results
including any surviving-but-equivalent mutants. Update the header's test count and
analyzer count.

Commit message explains the **user-visible consequence**, not the mechanics — what a
reader of the document would get wrong. End with:

```
Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>
```

---

## What this skill will not do

- **Decide that a feature needs an analyzer just because it was asked for.** Step 1 can and
  should return "do not build this".
- **Claim real-file validation without real files.** Step 5 is gated and says so.
- **Let a model judge its own output.** Findings are computed; the evidence tier is
  computed from provenance. Nothing here asks a model whether it got it right.
