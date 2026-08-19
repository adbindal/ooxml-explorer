# OOXML Expert Agent — Research State & Resume Point

> **READ THIS FIRST. Everything below is written so this work can resume cold, after a
> session ends or a quota runs out. Nothing important lives only in a conversation.**

## Where things stand (2026-08-19)

| | |
|---|---|
| **Branch** | `feat/schema-derived-rag-corpus` — **all work is here; `main` is untouched** |
| **Tests** | 1362 passing; typecheck, lint and production build clean |
| **Architecture** | **Complete.** Analyzer registry, one `Finding` type, question routing, capability ledger, gap log, versioned JSON report |
| **Analyzers** | package integrity, conformance, bookmarks, comments, fields, **tables**, **media**, OLE, pivot tables, formulas, content controls, hyperlinks, **revisions**, charts, style references, **footnotes**, **animations**, **external links**, Word cascade, Excel formats, PowerPoint inheritance — **21** |
| **In flight** | none. Everything is tested and registered. |

## How to resume in five minutes

1. **§9 Next actions** — the ranked to-do list. Start there.
2. **§8u** — *the criterion for what deserves an analyzer*, plus a measured backlog ranked by [MS-OI29500] deviation counts. This is the thinking to reuse, not just the conclusions.
3. **§8r–§8v** — the architecture as built, and why each decision went the way it did.
4. **§8** — the honesty ledger: what is contested, unverified, or was retracted. **Read before citing anything.**

## Standing rules, so they are not re-litigated

- **Deterministic first.** TypeScript resolvers decide what is true; the model only narrates a pre-verified evidence bundle. It never plans, routes, or adjudicates correctness.
- **The badge is computed, never asserted.** Ruled out with reasons recorded: vector DBs, re-ranking, embeddings on the main path, fine-tuning. Re-propose only with a new argument.
- **"Self-improving" means a gap log and a regression ratchet** — not a model writing its own rules or grading its own output. The loop closes through a person. See §8s.
- **Mutation-test every module.** A green suite on first run is not evidence; this has found a real gap in nearly every module here, most often a test passing for the wrong reason.
- **Say what was not verified.** Five citations have already been retracted on this project.

Published write-ups (private artifacts):
- Part 1 — The 29-Tag Problem: https://claude.ai/code/artifact/e1b1738f-2969-42c6-9437-190b0a29465f
- Part 2 — What the Schema Can't Tell You: https://claude.ai/code/artifact/d3fb80ba-ea48-411a-b04a-848c52418147
- Part 3 — Staff Engineer in a Box (the build plan): https://claude.ai/code/artifact/a76cb080-e4fd-4be5-b524-9ca490a94470
- The Analyzer Spine (design review): https://claude.ai/code/artifact/0133a5bf-a174-44dc-bc55-9f742ed48df8

---

## 0. The goal (user's words, restated)

Build an assistant that behaves like a **staff-level domain expert** across **Word, Excel and PowerPoint** — someone who has watched ECMA-376 evolve, knows how the Office apps actually implement it, and can connect markup to rendered output. Named use cases:

1. "Is this file/tag correct if I want it rendered X way in the Office app?"
2. "Compare these two files and tell me the *semantic* differences."
3. Mentor engineers new to OOXML: pitfalls, gotchas, connecting the dots.

Hard constraint carried from shipped code: the **honesty property**. The app shows a Grounded/Unverified badge and must never claim grounding it doesn't have.

---

## 1. Current shipped state (verified by reading code on `main`)

| Fact | Detail |
|---|---|
| Retrieval is **not** semantic | `ragRouter.getRagContext()` does an exact IndexedDB index lookup on tag name, filtered by domain |
| NL fallback | Only when query has a space or >15 chars: Gemini Nano extracts keywords → full cursor scan with `.includes()` → **takes first hit** |
| Corpus size | 29 records: 10 docx, 8 xlsx, 6 pptx, 5 shared |
| `grounded` flag | Computed **in TypeScript** from whether lookup hit. Model never asserts it. **This is correct and must be preserved.** |
| Zip dependency | `jszip ^3.10.1` (only archive dep) |

### Bugs / defects found — ALL FIXED in Stage 0, see §8b

1. **`staticKnowledgeBase.ts` is dead code.** `KNOWLEDGE_BASE` exported, imported by nothing — only the `ReferenceDoc` *type* is used. `services/README.md` claims the router falls back to it. It does not.
2. **All 8 xlsx records carry `"namespace": "r"`** — that's the relationships namespace, not SpreadsheetML. Source: the ingestion branch's generator prompt literally instructs the model to conflate them. User-visible as `<r:worksheet>` under a "Grounded" badge.
3. **Format drift**: `tbl` lists parents prefixed (`w:body`); all 28 others are bare (`body`).
4. **Context-budget bug (`services/geminiService.ts`)** — see §6. High priority.

---

## 2. `origin/feature/rag-ingestion` (stale branch @ 57fb89d) — salvage inventory

Diverged before all the hardening on `main`. **Do not merge or rebase.** Extract assets with `git show` onto a fresh branch.

| Asset | Verdict |
|---|---|
| `apps/agents/schemas/*.xsd` (wml 784 elems, sml 628, pml 315, dml-main 475) | **Take.** ~2,200 element declarations total |
| `mastra/xsdParser.ts` | Right idea, **rewrite** — it's regex over text lines, not a parser |
| `schemas/sdkClassMap.json` | **Superseded** — see §3, use SDK `/data` instead |
| `temp/wordprocessing.g.cs` (3.9 MB) | **Delete.** `generate_sdk_map.ts` fetches live from GitHub; nothing reads this file |
| Mastra/LibSQL LLM-judge workflow | Deliberately cut. Self-correction retry is *contraindicated* (see §5) |

---

## 3. Corpus map (all verified directly unless noted)

| Source | Content | Size | License |
|---|---|---|---|
| **dotnet/Open-XML-SDK `/data`** | `schemas/*.json`: 155 files, 9.2 MB. **1,839 types** across 4 core namespaces. Per-attribute validators (max length, numeric range, required) + **Office-version gating**. `namespaces.json` 168 prefixes. `schematrons.json` 134 KB | 9.2 MB | **MIT** |
| **[MS-OI29500]** | **1,895 clause-keyed entries**, each a literal *spec-says / Office-does* pair. Verified: toc.json = 2,152 nodes. Distribution: **Part 1 §18 (Spreadsheet) 534, §17 (Word) 529, §21 (DrawingML components) 264, §20 (DrawingML framework) 180, §22 (shared) 106, Part 4 §19 104, §19 (Presentation) 100**. ⚠️ **The 534 is misleading**: 203 of them (38%) are per-function calculation notes in §18.17.7 *Function Reference* and say nothing about markup. Strip those and SpreadsheetML has **316 markup-level clauses against Word's 529** — on markup, **Word is the more deviant format by a wide margin**. Also note 534 counts *clauses*, not variations: each page holds lettered items a., b., c., averaging ~3.3 in a 59-page sample | 1.2 MB DOCX | **Needs counsel** |
| **ECMA-376 XSDs** | Strict (21 files) + Transitional (26) + OPC (4) | ~1.8 MB | BSD-ish |
| **ECMA-376 Part 3 (MCE)** | 43 pages. Non-optional — every modern file uses `mc:AlternateContent` | 865 KB | Free |
| **ECMA-376 Part 1** | 5,039 pages reference manual. Chunk on element boundaries | 35 MB | Quote sparingly |

**Key URLs:**
- MS-OI29500 TOC: `https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/toc.json`
- MS-OI29500 DOCX: `https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OI29500/%5BMS-OI29500%5D-260519.docx`
  - **Verified: 1,221,467 bytes, valid OPC package, 107 parts, `document.xml` 13.4 MB uncompressed. `jszip` can read it.**
- SDK data: `https://raw.githubusercontent.com/dotnet/Open-XML-SDK/main/data/schemas/<ns>.json`
- Bulk Open Specs zip: `https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/Zip_Files/MSOFFSTAND.zip` (60 MB)

### ⚠️ Trap: the SDK `Summary` field is NOT a definitions source
100% coverage across all 1,839 types, but median **32 characters** and often wrong:
```
w:p   → "Defines the Paragraph Class."
w:t   → "Text."
w:r   → "Phonetic Guide Text Run."   ← wrong in the general case
```
**Human-readable definitions remain the only field with no free authoritative source.**

---

## 4. Domain knowledge: what the schema can't tell you

### The property resolution cascade (ECMA-376 §17.7.2) — the centrepiece
Six layers, in order:
1. `docDefaults`
2. table style + conditional formatting (gated by `w:tblLook`)
3. numbering properties
4. paragraph style
5. character style
6. **direct formatting**

Within layers 2–5, each style resolves along its `w:basedOn` chain first, using **four different merge semantics**:
- **Merge attributes** — `w:spacing`, `w:ind`, `w:rFonts`
- **Merge children** — `w:pPr`, `w:rPr`
- **Replace wholesale** — border elements (`w:top` etc). Derived element *drops* attributes it omits. Most surprising rule.
- **Conditional** — `w:tblStylePr`

### Toggle properties (12): `b bCs caps emboss i iCs imprint outline shadow smallCaps strike vanish`
- In a style: **toggles** inherited state. As direct formatting: sets absolute value.
- **Does NOT toggle along a `basedOn` chain** (there, later simply overrides).
- **Toggles across style types** (table vs paragraph vs character).
- ⚠️ **UNRESOLVED**: secondary sources say XOR; Microsoft's own OpenXmlPowerTools implements `higher ∧ ¬lower`. They disagree when higher=false, lower=true. **Testable against real Word — do this before shipping.**

### Word-vs-spec deviations with visible effect (sample from the 1,895)
| Spec | Word | Effect |
|---|---|---|
| `tblStyleRowBandSize` defaults to 1 | Defaults to **0**, meaning "no row banding at all" | #1 cause of "my table style doesn't work" |
| Conditional order: banding → rows → cols | Row banding → **col** banding → first/last **col** → first/last **row** | Banding precedence inverted |
| `w:tblLook` named attrs | Reads `@w:val` bitmask **only if no named attr present** | Editing `val` is a silent no-op |
| `w:ilvl` any integer | **Refuses to load** if <0 or >8 | Hard open failure. Exactly 9 list levels |
| Default fonts app-defined | **Times New Roman** all four `rFonts` slots | Bare file renders Times |
| Border `w:sz` unbounded | Clamps to **[2, 96]** eighths of a point | Silent normalization |

### Units registry (partial — needs completion)
twips (1 in = **1440**), half-points (`w:sz`), eighth-points (border `w:sz`, clamped 2–96), 240ths of a line (`w:spacing@line` when `lineRule="auto"`, else twips — **same attribute, two meanings**), fiftieths of a percent (`type="pct"`), EMU (1 in = **914400**, 1 cm = 360000, 1 pt = 12700), 1000ths of a percent (DrawingML colour transforms), hex byte as fraction (`themeTint`/`themeShade`).
❌ **"1/50th mm" does not exist in OOXML** — I invented it in a research brief; agent correctly refused to encode it. Only *fiftieths of a percent* is real.

### SpreadsheetML — the cell format model (researched 2026-08-17, all [RAW] from ECMA-376 PDF + MS-OI29500 clause pages)

🔴 **THE critical finding — Excel does not implement the cascade the spec describes.**
ECMA §18.8.9/18.8.10: *"both the cell style xf records and cell xf records shall be read to understand the full set of formatting applied to a cell."*
[MS-OI29500] §2.1.699 **and** §2.1.700, item b, identical text on both: *"In Office, **only the cell xf record defines the formatting applied to a cell**."*
→ `cellXfs[c/@s]` is **complete and self-contained**. `xfId` is provenance metadata, not an inheritance pointer. **Any resolver that merges `cellStyleXfs[xfId]` under `cellXfs[s]` will disagree with Excel.** This invalidates the mental model the standard sets up, and it is the single most consequential deviation in the format.

🔴 **`apply*` flags are NOT render gates.** [MS-OI29500] §2.1.721 (six items, a–f). In `cellXfs` they are **edit-time propagation/sticky bits**: `applyFont="0"` means "if the named style's font changes later, push it into this record", *not* "ignore this xf's fontId". Defaults are **asymmetric**: `true` in `cellStyleXfs`, `false` in `cellXfs`. The XSD declares no default for any of the six — the defaults exist only in the Microsoft document.

**Precedence is fallback-to-a-single-index, not a merging cascade:** `c/@s` → `row/@s` (**only if `row/@customFormat="1"`**) → `col/@style` (**only for cells "not yet allocated"**) → `cellXfs[0]`. Whichever wins supplies the *whole* format. There is no per-property inheritance anywhere. ⚠️ Row-vs-column precedence when both apply and no `c` exists is **not resolved by any clause** — inferred, not cited.

**The only true overlay layer in SpreadsheetML is `dxf`** (conditional formatting), which ECMA §18.8.15 explicitly defines as differential and applied *"on top of or in addition to"* existing formatting.

Other high-impact: `cm`/`vm` are **one-based in Office, zero-based in the spec**; Excel writes `left`/`right` borders where the spec says `start`/`end` (and Strict has **no** `left`/`right` at all — no border markup is valid in both); `font` children need Excel's **fixed sequence** despite being an `xsd:choice`; built-in `numFmtId` 14 is **`m/d/yyyy` in Excel, `mm-dd-yy` in the spec**; custom number formats cap at **206** and *"Office persists files that contain more than 206 custom formats (which it cannot load)"*.

**`t="s"` vs `t="str"`** is the format's most common misread: `s` means `v` is a **zero-based index into `sst/si`**; `str` means `v` is the string itself. And `sst` rich-text runs split exactly like Word's `w:r` — same mitigation (atomize, diff atoms, re-derive runs).

**Sparse structure:** `dimension` is optional and routinely stale; `row/@spans` is a **16-row block union**, never a per-row bound; `c/@r` absent ⇒ previous column + 1.

**Shared formulas are the corruption mechanism:** followers are *empty* `<f t="shared" si="N"/>` elements. Delete the master and the formulas exist nowhere. `calcChain.xml` is **pure cache and safely deletable** — ECMA §18.6 says the app *"is free to ignore"* it.

### PresentationML + DrawingML (researched 2026-08-17, all [RAW], dual-sourced via MS-OI29500 HTML *and* the DOCX)

**Sizing: DrawingML (444 variations) beats PresentationML (100) by 4.4× — and it is reusable across all three formats.** §21 = 264 (of which **charts alone = 172**), §20 = 180, §19 = 100 (of which **half, 49, are animation/timing** — the slide/shape model everyone asks about is only 18).

🔴 **The PPTX cascade — placeholder correspondence.** Only ONE matching rule is documented anywhere, and it is in MS-OI29500, not ECMA (§19.3.1.36(b)):
- **Slide shape → layout placeholder: match on `@idx`, never `@type`.** `@idx` defaults to 0.
- **Notes slide → notes master: match on `@type`.** Different rule, different part. One matcher for both is wrong for one of them.
- `idx = 0xffffffff` is a sentinel for "no correspondence" — legal `unsignedInt`, so **schema validation cannot catch its misuse**.
- ⚠️ **Layout → master matching is undocumented.** Not in ECMA, not in MS-OI29500. What python-pptx and others do (match `@type`, folding `ctrTitle`→`title`, `subTitle`→`body`) is **observed practice, not spec**.
- A master can only carry `title`, `body`, `dt`, `ftr`, `sldNum` — so it **cannot** be the shape-level ancestor of most layout placeholders. **Geometry inheritance stops at the master shape; `p:txStyles` contributes text only.**

🔴 **`a:xfrm`/`off`/`ext` are all `minOccurs="0"` — absent means INHERIT, not zero.** Writing `<a:off x="0" y="0"/>` as a "default" silently pins the shape to the top-left and severs the cascade. Reported as the #1 way generated decks come out visually wrong.

🔴 **Three disjoint colour alphabets.** `ST_SchemeColorVal` (17 values, what `schemeClr/@val` may say) vs `ST_ColorSchemeIndex` (12, the theme slot names) vs `p:clrMap` *attribute names* (12, the semantic keys). They overlap on `accent1-6`/`hlink`/`folHlink`, which is why the bug hides until a dark master. **The spec's own example makes `<a:schemeClr val="tx1"/>` resolve to `a:clrScheme/a:lt1`** — the *light* colour. `dk1/lt1/dk2/lt2` are the only four that bypass the map.

🔴 **Style-reference indexing is NOT uniform.** `fillRef`/`bgRef`: 0 or 1000 = none, 1–999 = `fillStyleLst`, **1001+ = `bgFillStyleLst`, 1-based** (normative, ECMA §20.1.4.2.10 / §19.3.1.3). `lnRef`/`effectRef`: **no offset** — they address one list each. And **ECMA §20.1.4.2.19 names the wrong list for `lnRef`** (`fillStyleLst`; Office reads `lnStyleLst`). `ST_StyleMatrixColumnIndex` has **no bounds**, so out-of-range `idx` is a semantic check only.

🔴 **`@rot` unit contradiction inside the standard.** ECMA §20.1.7.5 prose says **1/64000** degree; `ST_Angle` (§20.1.10.3), the schema, and Office all say **1/60000**. A 6.25% error that reads as a rendering artefact.

**Group transform** (the "why is my shape in the wrong place" mechanism): `sx = ext.cx / chExt.cx` (1 if `chExt` absent/zero); `child_abs = grp.off + (child.off − grp.chOff) × s`. Resizing a group changes `ext` but not `chExt`, which is how children silently acquire a scale. `p:spTree` **is itself a group**, so the transform applies at the root too. Nests recursively.

**`@lvl` is 0-based; `lvl1pPr`…`lvl9pPr` are 1-based.** `a:lvl1pPr` ↔ `@lvl="0"`. Nine levels, hard cap — exactly like `w:ilvl`.

**`+mj-lt` token grammar is UNDOCUMENTED** — appears exactly once in 5,039 pages, inside an example. `ST_TextTypeface` is an unrestricted string. Word uses a *completely different* syntax (`w:rFonts/@asciiTheme="minorHAnsi"`) against the same theme part.

🔴 **Three of the four inheritance hops are IMPLICIT relationships with zero XML reference** (slide→layout, layout→master, master→theme). You resolve them by opening the part's `.rels` and finding the relationship whose `Type` ends in `/slideLayout`. **`packageIntegrity.ts` cannot catch a missing one — there is no `rId` to dangle.** A slide with no layout relationship loses its entire inheritance chain and renders with defaults, silently. → task #13.

**Better MS-OI29500 ingestion path:** the DOCX flattened to text gives all 1,895 variations in ~34k lines. Entries delimited by `^Part \d Se\w+ [\d.]+, `; sub-items by `^[a-z]\.\s{3}`. Beats 1,895 HTML fetches, and no summarizer in the loop. **~25% of §19/§20/§21 entries are pure cross-references** (`19.3.1.44 spPr → 20.5.2.30(a-c)`) and must be **resolved at ingest**, including sub-letter selection, or they store nothing.

**Preset geometry**: the 187 `ST_ShapeType` definitions are NOT in the PDF — they ship as `OfficeOpenXML-DrawingMLGeometries.zip` (51,672 bytes) inside the ECMA Part 1 ZIP.

### Numbering (top bug source)
`numPr` → `numId` → `w:num` → `abstractNumId` → `abstractNum` → `lvl[@ilvl]`.
Three patterns: direct; style-linked (`lvl/pStyle` back-reference — how Heading auto-numbering works); **named list styles** where `abstractNum` has *no* `lvl` children, only `numStyleLink` → styles.xml → back to a *different* `num`. A resolver expecting `lvl` returns nothing and the list renders unnumbered.
- `numId="0"` is magic: means **remove numbering**, not "look up num 0".
- `numId` shares a *counter*; `abstractNumId` shares a *definition*. Confusing them = "list restarts mid-document" or "second list continues from first".

### MCE (Markup Compatibility) — required preprocessing pass
`mc:AlternateContent` / `mc:Choice Requires` / `mc:Fallback` / `mc:Ignorable` / `mc:ProcessContent`.
Modern Word writes shapes **twice** — DrawingML in `Choice`, VML in `Fallback`. Naive walkers either **double-count** (every textbox twice) or **zero-count** (find neither). VML is Transitional-only; it's the main reason real files aren't Strict.

### Strict vs Transitional (computed from official schemas)
- **Strict is a pure subset** — no type in Strict is absent from Transitional.
- **VML is Transitional-only** (5 vml-*.xsd files).
- **Percentages differ**: Transitional allows `50000` *or* `"50%"`; Strict only `"50%"`. A parser assuming one **mis-scales by 1000×**.
- Booleans: Transitional adds `on`/`off`, `ST_TrueFalseBlank` accepts empty string and capitalized variants.
- **Word writes Transitional by default and always has.** (A `learn.microsoft.com` Q&A page claiming Office 2013 defaults to Strict is **wrong** — good example of corpus poisoning from an authoritative-looking domain.)

### Namespace-year trap
`w15` → URI says `2012/wordml` but ships in Office **2013**. `w16se` → `2015/…` = Office **2016**. `w16cex` → `2018/…` = Office **2021**. **Never date a document from its namespace URIs.**

---

## 5. Architecture findings (post-audit; five claims were retracted — see §8)

### The core inversion
**Do not let the small model plan, route, or choose retrieval. Do it all in TypeScript.** The app already has the parsed DOM, selected node, ancestor chain, and style tree in memory. The model's only job is turning a small, pre-composed, already-correct evidence bundle into prose.

Supporting evidence (verified):
- *LLMs Cannot Self-Correct Reasoning Yet* (ICLR'24, arXiv 2310.01798) — no reliable self-correction without external feedback; sometimes degrades. **Retires the retry-loop design.**
- arXiv 2601.04254 — rule-based pattern matching hits **100% on structured information retrieval** where LLM approaches fail. That is exactly the XSD-validation / cascade-resolution case.
- Same paper: scaffolding gives **"amplification rather than compensation"** — gains only for models that already have the ability. **You cannot orchestrate your way out of a small base model.**

### Retrieval
- BM25/lexical beats dense retrieval by up to **49.9 points** on rare-entity lookups (EMNLP 2021). `w:kinsoku` is exactly that.
- Cascade: exact symbol → fuzzy (typos) → BM25 over prose → embeddings only for paraphrase tail.
- **But**: Sourcegraph production practice is **fan out to all retrievers and merge**, not route to one — retrievers are complementary. Merge with **Reciprocal Rank Fusion** (`rank_constant` default 60) since the three corpora produce non-comparable scores.

### The badge — four tiers, computed from provenance in the orchestrator
1. **Verified** — output of a deterministic checker
2. **Grounded** — claim checked against a specific retrieved span
3. **Unverified** — retrieval missed / check failed
4. **Abstain** — no evidence, high uncertainty

- ALCE (EMNLP'23) verbatim: *"on the ELI5 dataset, even the best models lack complete citation support 50% of the time."*
- arXiv 2605.06635: frontier models keep link validity >94% and relevance >80% but only **39–77% factual accuracy**; fact-check accuracy **drops ~42% as tool calls scale 2→150**. **More retrieval ≠ safer.** Argues for tight evidence bundles.
- **Design rule**: the **minimum** trust tier across cited evidence sets the badge, not the maximum. (Inverse of AWS Bedrock's documented any-chunk-wins behaviour.)
- Cheap win: pass **pre-attributed evidence IDs** into the prompt; reject any ID the model emits that wasn't in the bundle.

### Rejected / demoted
- **Fine-tuning — ruled out.** RAG roughly doubles unsupervised FT on knowledge injection (Ovadia, EMNLP'24). More decisively: a fine-tuned fact has **no source span**, so it can't be cited or checked — structurally destroys the honesty property. Also no adapter hook in Chrome's Prompt API.
- **Semantic entropy — demoted.** Detects *confabulation* (arbitrary inconsistency), explicitly **not** stable wrong beliefs. A confidently-wrong Nano scores low entropy and gets badged grounded. ~10× sampling cost.
- **Vector DB — no.** Few thousand structured records, primary access is exact key.
- **Re-ranking — not yet.** Presupposes multiple candidates; exact lookup returns 0 or 1.

### On-device NLI (for the Grounded tier)
`Xenova/nli-deberta-v3-xsmall` ONNX sizes (from HF API): fp32 284 MB / q4 230 MB / fp16 143 MB / q4f16 121 MB / **int8 90 MB**. Shippable next to Nano's 1.5–2.4 GB. ⚠️ INT8 typically costs 2–5% accuracy and MNLI accuracy for these quantized variants was **not verified**. A degraded entailment checker gating the badge is a real risk.

### Faithful narration of computed diffs
*Faithful Low-Resource Data-to-Text Generation through Cycle Training* (arXiv 2305.14793): human eval on WebNLG, hallucinations **14.84 → 2.57**, factual errors 8.05 → 0.49. Mechanism is a reverse text→data model.
**Our advantage**: the reverse direction needs no model. Diff records are structured and closed, so every element name, style ID and numeric value in a faithful narrative must appear in the input records — a deterministic string check. *(Extrapolation: nobody found this done for diffs specifically.)*

---

## 6. Browser / storage findings

- **Chrome allows up to 60% of total disk per origin.** Capacity is not the constraint. `localStorage` capped at 10 MiB — unusable for the corpus.
- **Eviction is all-or-nothing per origin.** MDN: if an origin stored via IndexedDB *and* Cache API, **both are deleted**. You cannot lose only cold shards.
- **Safari deletes script-created data after 7 days without user interaction.**
- `navigator.storage.persist()`: Firefox prompts; **Chrome silently auto-approves/denies via an undocumented heuristic**. Real gap.
- **Bulk insert**: 1,000 records in ONE transaction ≈ **80 ms**; one transaction each ≈ **2 s** (100×). Bottleneck is transaction handling, not throughput. ✅ `storageService.ts` already does the fast pattern.
- **512 MB V8 string wall** (`ERR_STRING_TOO_LONG`) on single-blob JSON persistence. Sharding is a correctness requirement past a size, not just perf.
- **Sharding pattern (Pagefind, read from source)**: chunk by posting-list entry count (default 20,000); record chunk boundaries as **word-prefix ranges** in a metadata file trimmed to shortest distinguishing prefix; content-hash filenames for free cache invalidation. ⚠️ **Normalize/stem BEFORE selecting the shard** — Pagefind issue #478 was exactly this bug.
- **Versioning**: Workbox manifest of `{url, revision}`; diff client-side, refetch only changed. Linear's `__schemaHash` → IndexedDB version bump → migration.
- **Curated-corpus rot**: steal GNU gettext's **fuzzy** flag. Key each curated entry to a content hash of the generated entry it annotates; on regeneration, mismatch → auto-mark fuzzy → **exclude from retrieval entirely** until a human clears it. Fail-closed.
- **Honest degradation (Notion)**: *"Opening a page and seeing half the content missing would be a worse user experience than not being able to open it at all."* Surface as available only when **fully** resident.

### Prompt API runtime facts
- Read `session.contextUsage` / `session.contextWindow` at runtime. **Never hardcode** — reported values range 6k–9k across Chrome versions and the docs publish no number.
- **Renamed**: `inputUsage`→`contextUsage`, `inputQuota`→`contextWindow`, `measureInputUsage()`→`measureContextUsage()`, `onquotaoverflow`→`oncontextoverflow`. Support both for older Chrome.
- Overflow evicts oldest prompt/response pairs; **system prompt is never removed**. `QuotaExceededError` carries `requested` and `contextWindow`.
- `responseConstraint` accepts a JSON Schema but **consumes context window**.
- **Chrome auto-removes the model when free disk drops below 10 GB.** Local AI is a *state*, not a one-time capability check.

### 🔴 OPEN BUG — `services/geminiService.ts`
```
line 250  analyzeFile:  f.content.slice(0, 8000)      ← per file, inside forEach, NO total cap
line 351  analyzeDiff:  f.original.slice(0, 8000)
line 352  analyzeDiff:  f.modified.slice(0, 8000)
```
Both sit **above** the provider branch — same string reaches `promptLocalModelForJson` (lines 298, 405). At ~3 chars/token for XML: 1-file explain ≈ 2,700 tok; 2-file explain ≈ 5,300; 1-file diff ≈ 5,300; 2-file diff ≈ **10,700** — against a 6k–9k window **shared with output**. `grep` for `contextWindow|measureContextUsage` across repo = **0 hits**. DLP mode forces local-only, so privacy users hit this hardest.

---

## 7. Tooling landscape

- **No tool exists** that says "this markup is legal but won't render how you want." That is the gap the agent fills.
- Open XML SDK `OpenXmlValidator`: 4 error classes (Schema/Semantic/Package/MarkupCompatibility). Semantic tier has **21 constraint types** incl. `ReferenceExistConstraint`, `RelationshipExistConstraint`, `IndexReferenceConstraint`, `UniqueAttributeValueConstraint`. Zero concept of visual outcome. Default ctor targets **Office2007** — footgun. .NET only.
- **Why naive XML diff fails**: (a) `w:rsid*` churn — Eric White wrote a tool in 2008 purely to strip these; a re-saved file diffs as heavily changed. (b) **Run splitting** — same sentence can be 1 or 20 `w:r` with identical formatting.
- **The technique to steal (WmlComparer)**: flatten each document to an array where every item is **one content atom** (a character, a paragraph mark, an image), diff the atom arrays, re-derive runs on output. Run boundaries stop mattering. *(Archived by Microsoft; dormant since 2022 — take the insight, not the dependency.)*
- **Don't use LibreOffice as a diff oracle** — its Compare ignores format-only changes entirely (bold→italic = 0 redlines) and ignores footnote edits.
- Generic XML tree-diff (XyDiff, X-Diff, TED) doesn't scale past a few hundred nodes and has no notion of which differences are *visible*. `document.xml` is 10⁴–10⁵ nodes.
- **Nothing in JS models OOXML semantics** — no style resolver, no numbering resolver, no effective-formatting computation. **We build that.**
- `docxodus` (MIT, WASM PowerTools fork) — only client-side semantic differ; `getRevisions()` returns structured records. But ~9 months old, ~single maintainer, ~3.3 MB brotli, ~1 s cold start, **all claims self-reported**. Prototype behind an interface; don't marry.

---

## 7b. Date-system corrections (2026-08-17)

Two things I had recorded as fact are **not supported by primary sources**:

- ❌ **"Lotus 1-2-3 provenance of the leap-year bug"** — not in ECMA-376 Part 1 (all 5,039 pages grepped) and not in any MS-OI29500 clause opened. Widely repeated, no primary source found. **Do not encode as sourced.**
- ❌ **"29 February 1900" is never named in ECMA-376.** The phantom day's *existence* is provable from serial arithmetic; its *identity* is not, from these sources. ECMA §18.17.7.344 (`WEEKDAY`) points at §18.17.4.1 for *"special handling of certain days in 1900"* — and §18.17.4.1 **contains no such handling**. That is a dangling cross-reference in the standard itself.

✅ **What IS verified** (ECMA §18.17.4.1, verbatim): 1900 base = **1899-12-30** (serial 0); 1904 base = **1904-01-01**; offset **1462 days**, cross-checked two ways. Annex L defines a *third* base, 1900-backward-compat = **1899-12-31**. [MS-OI29500] §18.2.28(j): **Strict ⇒ true 1900 base; Transitional ⇒ compat base**, and *"Excel does not support negative serial numbers"* — so the spec's entire pre-1900 range is unreachable.

⚠️ Also unresolved: ECMA Annex L §L.2.16.9.3 contradicts normative §18.17.4.3 on whether serial 1.5 or 2.5 is 1900-01-01T12:00. Trust the normative clause.

⚠️ The **"numFmtId 0–163 reserved, custom starts at 164" rule has no normative basis** — no such statement exists in ECMA-376 Part 1. Attested only by the spec's own examples and Excel behaviour. **Convention, not law.**

⚠️ The [MS-OI29500] Excel column-width formula **appears to contain an error** (`+ (… MOD 8)` where the intent looks like `−`). Quoted verbatim, not corrected. Needs empirical testing against real Excel before being encoded.

## 8. Honesty ledger — what is contested, unverified, or was retracted

**Retracted during a citation-repair pass** (an agent disclosed it had cited papers it never opened):
1. "80% at 2 hops → 0% at 4 hops" — no source; fabricated by a search summarizer.
2. Four-item citation failure taxonomy — grepped the paper: 0 hits for every term. Invented.
3. "Perplexity / OpenAI Search benchmarked" — paper actually covers 14 LLMs.
4. "GroundEval" — arXiv 2607.01793 is *Safety Testing LLM Agents at Scale*.
5. BEAVER — real paper, but offline distributional bounds, not runtime validation. Circulating numbers don't match its abstract.

**Diagnostic worth remembering:** *every* fabrication entered through a WebFetch/WebSearch call that **succeeded** and returned confident, well-structured, invented detail. Failed fetches were harmless — they announced themselves. **Verify numbers against raw source text (`curl` + `grep`), not summarizer output.**

**Still open:**
- **MS-OI29500 licensing** — grant is purpose-bound ("in order to develop implementations"), with reservation of rights. RAG-to-answer-questions looks like the intended shape; bulk republication doesn't. **Needs counsel. Gates the highest-value asset.**
- **Toggle truth table** — testable against real Word.
- **Per-flag effects of the 65 `w:compat` elements** — need extraction from ECMA-376 §17.15.1.
- **Provenance tracking in RAG** — no credible literature found. Treat the three-corpus design as unprecedented.
- **Chrome `persist()` heuristic** — undocumented.
- **MNLI accuracy of quantized ONNX NLI variants** — unverified.
- ECMA-376 in-document copyright notice — not extracted.

---

## 8b. Stage 0 — COMPLETE (2026-08-17)

Three branches pushed to origin, all off `main`, none merged:

| Branch | Commit | What |
|---|---|---|
| `fix/ai-prompt-context-budget` | `3af80df` | Context-budget fix + `services/promptBudget.ts` + 22 tests |
| `feat/schema-derived-rag-corpus` | `534dd10` | Corpus 29 → **1,521** records via `scripts/ingestSchema.ts` |
| `feat/schema-derived-rag-corpus` | `2ecd45d` | Offline fallback wired up; `services/README.md` corrected |
| `docs/ooxml-expert-agent-research` | `9551657` | This file + the build plan |

**All four defects from §1 are fixed**, plus two found en route: `ragRouter` mapped the `shared`/DrawingML domain onto the *Presentation* SDK namespace, and an IndexedDB rejection propagated out of `getRagContext` instead of degrading to the offline store.

**Corpus now:** docx 603, xlsx 383, shared/DrawingML 323, pptx 212 = 1,521. Only 0.27 MB minified, so the sharding work in §6 is **not yet needed**. 29 records are `curated` (prose + citation); 1,492 are `schema` (structure only, deliberately carrying no definition and no citation).

**Design rule established — *prose from humans, structure from the schema*.** Curated records are authoritative only for `definition`/`citation`/`reviewerNote`; every structural field comes from the schema whenever a schema record exists. Worth remembering why: the first attempt let curation win outright, and the xlsx namespace bug *survived*, because the curated values were precisely the wrong ones.

**Two derivations in `scripts/ingestSchema.ts` that aren't obvious from the source data:** parents are the inverse of each element's `Children` list; and attributes are usually inherited (only 174 of 726 WordprocessingML entries declare them directly, 408 come via `BaseClass`).

**`ragRouter` honesty change:** schema-derived records have no prose, so the prompt now states that no specification description is on file and instructs the model not to cite one. Structure stays grounded, meaning explicitly does not. This is the minimal honest version of the badge question — the full four-tier ladder is still unbuilt.

**Mastra assessed and declined** for ingestion: the pipeline contains no LLM, so a workflow framework adds a dependency and buys nothing. Revisit only if prose generation later needs durable suspend/resume for human review — but note the research ruled out the LLM-judge and self-correction-retry patterns that were the old Mastra pipeline's entire purpose.

## 8c. Stage 1a — package integrity (2026-08-17)

`services/packageIntegrity.ts` + 21 tests, commit `1a7905f`. **The first code in the app that answers a correctness question by computation rather than retrieval** — every finding is derived, so it can be presented as *verified* rather than merely grounded.

Rules: `missing-content-types`, `untyped-part`, `dangling-relationship-id`, `missing-relationship-target`, `orphaned-rels-part`, `malformed-xml`.

Two design points worth not re-deriving:
- **Relationships resolve per part.** An image used by `header1.xml` must be declared in `word/_rels/header1.xml.rels`; `document.xml.rels` does not satisfy it. Word rejects the lenient reading; some readers accept it. There is a test asserting rejection.
- **References found by scanning, not enumerating** — any attribute in the relationships namespace holding an `rId`. Enumerating `r:id`/`r:embed`/`r:link`/… silently misses format-specific ones.

Format-agnostic by construction (packaging is fully shared across the three formats). Pure functions over a `Record<partPath, content>` map, decoupled from JSZip.

✅ **Wired up** (`af11bf5`, Stage 1c). `readPackageParts()` in `zipService.ts` adapts JSZip to the checker's path→content map (binary parts map to `''` — presence is all the checks need). A "Check Package Integrity" action in the Validator reports findings into the existing log console, styled apart from the AI actions because nothing on that path is generated, retrieved, or networked. Verified against a real 107-part package.

**Stage 1b — MCE preprocessing** shipped (`2ad039f`). `services/markupCompatibility.ts` + 18 tests.
- Resolves `mc:AlternateContent` to exactly one branch. Without it a reader double-counts (every shape twice, since Word writes DrawingML in `Choice` and VML in `Fallback`) or zero-counts. Both fail silently.
- Understood-namespace set is a **parameter**, with `LEGACY_CONSUMER_NAMESPACES` and `MODERN_CONSUMER_NAMESPACES` presets. Resolving both ways answers "why does this look different in an older Word?" — the Validator reports when they diverge.
- Namespace URIs taken from the SDK's `namespaces.json`, not memory. `w15` = `.../2012/wordml` but ships in Office **2013**; a test pins it.
- `Requires` prefixes resolve against **in-scope** declarations, all-not-any, first match wins. Nesting iterates to a fixed point, bounded.
- **Integrity checks deliberately run against UNRESOLVED markup** and are reported separately: a broken relationship inside an `mc:Fallback` is still broken for whoever takes that branch.
- ⚠️ Real-file check found `mc:Ignorable="w14 w15 wp14"` but **zero** `AlternateContent` (a text-only spec document). So the `AlternateContent` paths are unit-tested on synthetic fixtures only — **not yet exercised against a real file containing shapes.** Worth doing with a real .docx that has a text box. Naive walkers double-count every textbox, because Word writes shapes twice — DrawingML in `mc:Choice`, VML in `mc:Fallback`. Required before the Stage 2 resolvers. The seam is clean: `packageIntegrity` consumes a path→content map, so MCE slots in as a transform over that map.

## 8d. Stage 2 — Word cascade COMPLETE (2026-08-17)

All six layers of ECMA-376 §17.7.2 now supported.

| Module | Commit | Covers |
|---|---|---|
| `services/wordStyleResolver.ts` | `e8aa678` | Layers 1, 4, 5, 6 + `basedOn` roll-up with 4 merge semantics + 12 toggles |
| `services/wordNumbering.ts` | `ab6d091` | Layer 3 |
| `services/wordTableStyles.ts` | `ab6d091` | Layer 2 |

Layers 2 and 3 are **supplied by the caller**, not resolved inside the resolver — each needs context it doesn't have (cell position in the table; the numbering part). A caller that omits layer 2 for a run that *is* in a table gets a trace note, not a silently incomplete answer.

**Also shipped:** corpus expanded to **1,899 records across 11 namespaces** (`bbfb65a`) including full DrawingML; namespace-blind lookup bug fixed (`ac0ec80`).

### Open conflicts, deliberately not papered over
- **Toggle truth table** — spec prose vs Microsoft's OpenXmlPowerTools disagree in one case. Resolver marks the result `uncertain` rather than guessing.
- **Numbering vs paragraph-style indentation** — MS-OI29500 says `lvl/pPr` indentation *overrides* the paragraph style; the §17.7.2 cascade puts numbering at layer 3 and paragraph styles at layer 4, i.e. the opposite. Cascade order ships (better sourced); trace records both. **Settle against real Word.**

### Word deviations now encoded and tested
`tblStyleRowBandSize` defaults to **0** in Word (not 1) → no banding; Word's conditional order is row banding → col banding → first/last **col** → first/last **row** → corners; `w:tblLook`'s legacy bitmask is read **only** when no named attribute is present; `numId="0"` means *remove* numbering; `ilvl` outside 0–8 makes Word refuse the file.

## 8e. The Verified tier is live (2026-08-17)

The chain now closes end to end: **selected element → cascade resolved over the real package → pre-verified evidence → tier computed from provenance → badge**.

| Piece | Commit |
|---|---|
| `services/wordFormattingAnalysis.ts` — composition layer | `a42c4d1` |
| `selectEvidenceTier` + three-state badge | `cc43407` |
| `AIPanel` builds the evidence | `ba2426a` |

**Tier rule:** minimum across sources *actually present*. A complete computation alone = `verified`; a computation with gaps = `grounded`; a citation = `grounded`; nothing = `unverified`. Absent evidence is deliberately **not** a weak source — a corpus miss must not drag a computed answer down, or the common case (tag outside the curated 29, cascade resolved perfectly) would be mislabelled.

**The locator refuses to guess.** Real documents are full of identical paragraphs; a match must be unambiguous or it returns null and the caller degrades to the ordinary explanation. Same rule one level down for runs.

**Still Word-only and `word/document.xml`-only.** Headers, footers and footnotes are not yet covered, and Excel/PowerPoint have no resolver, so those paths never reach `verified`.

## 8f. Stage 2 — Excel resolver shipped (2026-08-17)

`services/excelStyleResolver.ts` + 31 tests, commit `5874a3c`.

**Built to what Excel does, not what the spec says.** The trace states on every resolution that `xfId` is provenance and is *not* merged in, so nobody "fixes" it back toward the standard later.

- **Fallback chain, never a merge:** `c/@s` → `row/@s` (only when `customFormat="1"`) → `col/@style` (only for unallocated cells) → `cellXfs[0]`. Whichever wins supplies the whole format.
- **`apply*` flags ignored for rendering**, with a note explaining they are edit-time propagation bits with asymmetric defaults that exist only in the Microsoft document.
- **`numFmts` keyed by explicit `numFmtId`**, unlike every other component table, which is positional.
- **Built-in format table carries both readings** where Excel and the standard disagree (14, 22, 37–40, 47). Excel's is reported by default.
- **All three date epochs**, including the 1900 compatibility system's phantom day. Existence is provable from the standard's own arithmetic; identity is never stated in ECMA-376, so the code does not claim one.

✅ **Wired** (`ad0abb0`). `services/excelFormattingAnalysis.ts` composes it over a package and the panel routes both formats through one table, keyed on **which part is open** rather than the file extension — a selection inside `xl/styles.xml` has no cell to resolve.

Value interpretation is where most of the care went, because it is easier to get wrong than the format:
- **`t="s"` means `<v>` is an INDEX** into the shared string table; **`t="str"` means `<v>` IS the string.** One character apart, opposite meanings. Both tested against the same `<v>0</v>`.
- Shared-string runs are concatenated — same run-splitting problem as Word's `w:r`.
- A date is a number *plus a date-shaped format*; nothing in the cell says so.
- A formula with no cached `<v>` is flagged; libraries write these routinely and the cell renders blank everywhere except Excel.

## 8g. Stage 2 COMPLETE — all three formats (2026-08-18)

| Format | Resolver | Composition | Commits |
|---|---|---|---|
| Word | `wordStyleResolver` + `wordNumbering` + `wordTableStyles` | `wordFormattingAnalysis` | `e8aa678`, `ab6d091`, `a42c4d1`, `641434d` |
| Excel | `excelStyleResolver` | `excelFormattingAnalysis` | `5874a3c`, `ad0abb0` |
| PowerPoint | `powerpointResolver` | `powerpointFormattingAnalysis` | `962cc01`, `1717bb3` |

All three reach the **Verified** tier through one table in `AIPanel`, keyed on *which part is open* rather than file extension. Word covers every story — document, headers, footers, footnotes, endnotes, comments — but deliberately **not** sub-folders: `word/glossary/document.xml` has its own `styles.xml`, and resolving it against the main one would report formatting Word never applies.

**PowerPoint's structural difference:** the whole inheritance chain is implicit relationships. Nothing in `slide1.xml` points at its layout. `resolveSlideChain` walks slide → layout → master → theme and names whichever hop broke — the failure that otherwise renders with defaults and no error.

**Three bugs caught by tests during PPTX work**, each a case where the code looked right: geometry reported `master` without checking the master carried a transform; the empty-element normalizer collapsed the wrong tag pair (`<p:spPr/></p:sp>` → name `p:sp` + `Pr/`); and the shape was located in one parsed DOM while the chain re-parsed the slide independently, so identity comparison always failed and every shape read as "not a placeholder".

## 8h. Retrieval — measured, not guessed (2026-08-18)

`services/retrievalMetrics.ts` (`b839eb9`) counts which of six paths answered each lookup, surfaced in the Validator.

**This is the precondition for the embeddings decision, not a step toward it.** Embeddings, vector DB, re-ranking and fine-tuning were all *ruled out* in Parts 1–2 with reasons — they are not pending work. The only genuinely open retrieval item is that the NL fallback still takes the **first substring hit with no ranking**, and the evidence-backed fix is **BM25** (pure JS, works under DLP), not embeddings. Task #21.

Counts only, never query text — a search log is precisely what DLP mode exists to prevent.

## 8i. Both Word conflicts SETTLED (2026-08-18) — and my premises were wrong

Research found both answers already documented. **Prefer [MS-OI29500] over ECMA prose without hesitation** — it is Microsoft's normative statement of what Word does, and on both questions it is corroborated by independent code.

### Toggles — Word RESETS, it does not toggle. ✅ FIXED (`services/wordStyleResolver.ts`)

[MS-OI29500] **§2.1.258** (Part 1 §17.7.8, Paragraph Styles) and **§2.1.246** (§17.7.6, Table Styles), identical wording: *"The standard specifies that the resolved value of the toggle properties will toggle the previous level (True) or leave it unchanged (False) ... Word resets the value of the toggle property to the value specified by the [paragraph|table] style if a value is present."*

**Three premises I had wrong:**
1. The ECMA sentence I cited is from the **withdrawn 2006 1st edition** — removed in the 5th. The current rule is XOR in **§17.7.3**, not the per-element prose.
2. **OpenXmlPowerTools is not a model of Word.** It matches neither — differs from the spec in one cell and from Word in a *different* one.
3. So the code was **uncertain about the cell it got right and confident about the cell it got wrong**.

| higher | lower | ECMA XOR | **Word (reset)** | OXPT |
|---|---|---|---|---|
| T | F | T | **T** | T |
| F | T | T | **F** | F |
| T | T | F | **T** | F ← old code was here, wrong |
| F | F | F | **F** | F |

Now reports a **divergence** ("Word applies on; a strictly conformant consumer would apply off") rather than an uncertainty. ⚠️ Caveat: reset notes cover paragraph and table styles only — no equivalent note for **character** styles, so applying it there is inference.

### Numbering — conditional on provenance. ✅ FIXED (`c1d1969`)

[MS-OI29500] **§2.1.229** (Part 1 §17.7.2, Style Hierarchy): *"Word applies the properties from a paragraph style applied to a paragraph before it applies the properties from a numbering style applied to a paragraph **via numbering properties**."*

That qualifier is load-bearing:
- Numbering via **`w:numPr` on the paragraph** → **numbering wins** (apply AFTER the style chain)
- Numbering via the **paragraph style** → **paragraph style wins** (current ECMA order is correct)

OpenXmlPowerTools `FormattingAssembler.cs` lines 2163–2192 encodes exactly this conditional (`lii.FromParagraph` vs `lii.FromStyle`), which is strong corroboration.

**My §17.9.22 citation was misattributed** — that clause only says direct formatting beats numbering, which both readings already agree on.

**ECMA contradicts itself:** §17.7.2's prose puts numbering at layer 3, but the **figure on the same page** orders it Document Defaults → Table → **Paragraph → Numbering** → Character → Direct. Verified by coordinate extraction *and* by rasterising the page. The figure sides with Word, and the contradiction has been there since 2006.

**Fixed.** `CascadeContext.numberingSource` carries the provenance; `wordFormattingAnalysis` sets it (`paragraph` unless it had to fall back to the style). Also implemented from the same research: numbering-level `pPr` restricted to `jc`/`ind`/`tabs` per §17.9.22(b) — a level setting `keepNext` was being honoured when Word ignores it — and style-hierarchy `w:ind` dropped when `numId="0"`, without which a cancelled list item keeps its list indent and sits out of line.

Tests verified to fail against the old placement: restoring it breaks 2 of the 5.

⚠️ **LibreOffice diverges from Word here** (`SwTextNode::AreListLevelIndentsApplicableImpl`): a paragraph style setting `w:ind` *suppresses* the list level's indent. Known interop bug class.

## 8j. MS-OI29500 licensing — ✅ RESEARCH COMPLETE

**Findings are in [`LICENSING.md`](./LICENSING.md)** — verbatim IPR notice (landing page and DOCX verified word-for-word identical), what is clearly permitted, what is genuinely ambiguous, verified precedent, contact routes, and **four questions a lawyer can be sent as-is**.

### The three things that matter most
1. ✅ **Patents are settled.** `[MS-OI29500]` is expressly covered by the Open Specification Promise — confirmed via the OSP page *and* the machine-readable Patent Map (`ows_Programs=";#OSP;#"`). Not the Community Promise.
2. 🔴 **Copyright is not settled, and the OSP does not touch it.** Microsoft's own OSP FAQ: *"Copyrights in the Covered Specifications are not provided through the OSP."* ⚠️ **Apache POI cites the OSP as authority for a copyright question** — a widely-copied conflation, not a safe basis.
3. 🔴 **The general Learn terms of use flatly prohibit this**, and the IPR notice is the only carve-out — *"Regardless of any other terms…"*. So the scope of that carve-out is load-bearing, not academic.

**Precedent worth showing a lawyer:** Samba ships **52,098 lines** of bulk-scraped Open Specifications tables in a public GPL repo with the notice reproduced verbatim — further than we plan to go. ⚠️ But Samba holds a **separately negotiated PFIF agreement**, so show that fact at the same time. At the other end, python-docx has **zero** references and documents the same facts from its own observation of Word.

**The asymmetry to notice in the notice itself:** schemas, IDLs and code samples get *"with or without modification, **any**"*; the prose gets only *"portions… as necessary"*. A structured transformation of the prose falls in that gap, and Reservation of Rights means silence is no.

*(Original brief retained below in case the research is ever re-run.)*

### What to establish
1. **The verbatim IPR notice** on an actual [MS-OI29500] document (landing page, DOCX or PDF) — Copyrights, Patents, No Trade Secrets, Reservation of Rights, Trademarks, Tools. **Quote it exactly.** Paraphrasing a licence is the failure mode to avoid; a summary of licence terms is worth less than a URL.
2. **The scope of "in order to develop implementations of the technologies described in this documentation"** — the load-bearing phrase. Does Microsoft publish any clarification? Does it distinguish a tool that *implements* a format from one that *teaches* it? Report what the text says and what is genuinely ambiguous; do not resolve the ambiguity.
3. **Distributing portions in an implementation vs. republishing a derived database.** The notice permits distributing portions "in your implementations ... or in your documentation as necessary to properly document the implementation". A RAG knowledge base is arguably a *derived database* rather than a quotation. Anything addressing bulk extraction or structured transformation?
4. **Patents are a separate question from copyright.** Open Specification Promise vs Community Promise, which applies here, and where the per-document mapping is published. Answering one does not settle the other.
5. **Precedent.** Do real projects ingest or systematically reference Open Specifications content — python-docx, docx4j, Apache POI (an ASF project with a formal IP review), LibreOffice, sheetjs? **How do they attribute it, and do they quote or only cite?** Report what is observable in their source or docs.
6. **The route to certainty.** Microsoft publishes a contact path for licensing questions on these documents. Find the current one, plus any public forum or documented process for scope questions.
7. **Fallbacks if the answer is no.** Citing clause numbers without reproducing content (a pointer, not a copy); linking the public page; deriving equivalent facts independently by testing real Office output. Say honestly how much value each preserves.

### Why it matters
[MS-OI29500] is **1,895 clause-keyed entries** of "the standard says X, Office does Y" — the only source of that knowledge. It has already proved decisive twice in this project: it settled the toggle-reset question (§8i) and the Excel `cellXfs` question (§8f), both cases where building faithfully from ECMA would have produced code that disagrees with Office **and passed review**.

### Deliverable
Verbatim terms → clearly permitted → genuinely ambiguous → precedent → route to certainty → fallbacks, ending with **three or four questions a lawyer can be sent as-is**. Those questions are the most reusable output even if everything else is thin.

## 8k. Charts — extraction for translation (2026-08-18)

`services/chartSemantics.ts` + 26 tests, `8dcb923`. Built for **converting** a chart to another format (the stated tviz case), not for rendering — different jobs, and the difference is which properties *mean* something.

**Three converter questions it answers:**
1. **Where the data lives.** Values are a literal list or a reference to a range, and a reference carries a *cache*. The cache is what the producing app wrote at save time, not live data — but a chart in a `.docx`/`.pptx` often ships no workbook, so it is frequently the only data there is. `cacheIsOnlySource` says which case applies.
2. **Structure vs paint.** `PRESENTATIONAL_ELEMENTS` names what is safe to drop. Dropping `spPr` loses styling; dropping `order` silently reorders the series — both are "just properties" in the markup.
3. **What will not survive.** Combination charts, log scaling, reversed axes, undrawn axes that still affect scaling, `sourceLinked` number formats — named up front.

**Traps encoded:** `idx` (identity) vs `order` (display position) coincide in simple charts, which is why conflating them hides; points are **sparse with explicit indices**, so reading in document order shifts later values by one; `sourceLinked="1"` means the format comes from the source cells so `formatCode` alone misleads.

✅ **Wired to the panel** in `d1be6d8` via `computeChartEvidenceForMarkup(parts)` and a format-agnostic `ANALYSIS_TARGETS` entry matching `/charts\/chart[^/]*\.xml$/`. Chart parts are self-contained, so the entry declares no sibling parts.

## 8l. Bookmarks were being deleted by the diff — fixed (2026-08-18)

`d1be6d8`. `services/ooxmlDiff.ts` listed `bookmarkStart`/`bookmarkEnd` in `EPHEMERAL_ELEMENTS`, next to `proofErr` and `rsid`. **They are not noise.** A bookmark is the target of every hyperlink, cross-reference and TOC entry, so a document that lost all of them was reported as having *no semantic changes*.

Two parts to the fix, and the second is the one that is easy to miss:
1. Removed them from the strip list. `NOISE_BOOKMARK_NAMES = new Set(['_GoBack'])` keeps out the one that genuinely is noise — Word writes it to remember the last edit position and rewrites it on every save.
2. **Bookmarks are direct children of `w:p`, not of `w:r`.** The atomizer only walked runs, so they stayed invisible even after surviving normalisation. `ContentAtom.kind` gained `'anchor'` and atomize now walks paragraph direct children.

## 8m. Bookmarks — the id/name split (2026-08-18)

`services/wordBookmarks.ts` + 31 tests, `d0551a7`. Task #26 closed.

**The model.** A bookmark is a pair of empty markers matched by **`@w:id`**, and **only the start carries `@w:name`**. So a lost end does not make a malformed bookmark — it makes *no* bookmark, and every hyperlink, cross-reference and TOC entry aimed at that name resolves to nothing while the file still opens and looks right. Because the end has no name, an unmatched end tells you a bookmark was lost but **not which one** — the module says so rather than inventing an answer.

**The corruption class, and its provenance.** `@w:id` is drawn from a space shared with tracked changes (`w:ins`, `w:del`), permissions (`w:permStart`) and every `*Change` element — 22 element types. A generator numbering its revisions from 1 collides with existing bookmark ids in any document that has bookmarks, and documents routinely have hundreds. **Word rejects the file as corrupt; macOS Preview and most libraries open it fine**, so the bug passes every test and reaches users. `findMarkupIdCollisions` detects it; `nextSafeMarkupId` returns the id to start from — one past the max across *all* element types, never one past the max bookmark id.

⚠️ This is **observed Word behaviour**, sourced to a concrete generator bug report ([anthropics/skills#489](https://github.com/anthropics/skills/issues/489)), **not a rule stated in ECMA-376**. The spec assigns `w:id` per element type without declaring the space shared; Word is stricter than the schema. Recorded as behaviour with its source, in line with §8's honesty ledger.

**Verified against SDK schema data** (`CT_Bookmark`, `CT_MarkupRange`, `CT_MoveBookmark`, `CT_Markup`): `@w:name` and `@w:id` both required on `bookmarkStart`; **name capped at 40 characters**; `@w:id` is `ST_NonNegativeDecimalNumber` unioned with signed numbers ≤ -2, so **-1 is the one integer the union excludes**.

`moveFromRangeStart`/`moveToRangeStart` are separate range kinds (`CT_MoveBookmark`) that must not close a `bookmarkStart` sharing their id.

## 8n. OLE objects — the preview that hides breakage (2026-08-18)

`services/oleObjects.ts` + 24 tests, `f6bfc91`. Task #27 closed. Cross-format.

**The failure.** Every OLE object ships a preview image, because no OOXML consumer can render a foreign binary. Drop the embedding, keep the preview, and the page is **pixel-identical** — Word and PowerPoint open the file without complaint and it breaks months later on a double-click. *"It renders correctly" is not evidence the object survived.* `findSilentlyBrokenOleObjects` lists exactly the objects that look fine and are broken anyway, and deliberately excludes a missing preview (visible) and a missing `progId` (still works).

**The same concept, three incompatible expressions** — directly relevant to the spec-translation use case, because a converter cannot ask one question of all three:

| Format | Embedded-or-linked | Preview |
|---|---|---|
| Word | `o:OLEObject/@Type` = `Embed`\|`Link` — an **attribute** | sibling VML `v:shape > v:imagedata/@r:id` |
| PowerPoint | `p:embed` vs `p:link` — a **child element choice** | `p:oleObj > p:pic`, a real DrawingML picture |
| Excel | `@link` **presence** (a formula reference) | VML shape in the sheet's legacy drawing, reached from the worksheet |

⚠️ **`o:OLEObject` attributes are unprefixed and PascalCase** — `Type`, `ProgID`, `ShapeID`, `DrawAspect`, `ObjectID`, `UpdateMode` — against the lowerCamelCase convention the rest of OOXML follows. A namespaced lowercase lookup finds nothing, twice over. Verified against the SDK schema for `urn:schemas-microsoft-com:office:office`.

`oleDataIsPresent` returns **null** for a linked object rather than false: the target is outside the package by definition, so *"cannot check"* and *"is missing"* stay distinct and only one is a defect. Same discipline as the Verified tier.

## 8o. Comments — three parts, and the one Word may not have written (2026-08-18)

`services/wordComments.ts` + 43 tests, `878c26d`. Task #25 closed. Built by a subagent; facts below were re-verified before merging.

**The anchor** is a range (`commentRangeStart`/`End` matched by `@w:id`) plus a `commentReference` marking the display point — the same paired-marker problem as bookmarks, and `commentRangeText` keeps the same three-state contract as `bookmarkText`: `null` = no answer exists, `''` = anchored to a point, text = the covered span.

**Threading is in a side-car Word may not have written.** `w15:commentEx` in `word/commentsExtended.xml` is keyed on the **`w14:paraId` of the comment's *last* paragraph**, not on the comment id. So *"is this a reply?"* and *"is this resolved?"* are unanswerable from `comments.xml` alone. When the side-car is absent `threadingKnown` is false and `commentThreads()` returns **`null`, not a flat list of roots** — because a flat list of roots is exactly what a threaded document missing its side-car looks like, so returning one would manufacture the confusion the module exists to prevent.

🔴 **My brief to the agent had the namespace URIs wrong, and the agent caught it.** Independently re-verified against `data/namespaces.json`:

| | Correct | What I wrote |
|---|---|---|
| w15 | `http://schemas.microsoft.com/office/word/2012/wordml` | …`/2012/wordml/main` |
| w14 | `http://schemas.microsoft.com/office/word/2010/wordml` | …`/2010/wordml/main` |

**These Microsoft extension namespaces have no `/main` segment**, unlike the ECMA ones. Two mutants confirm it: adding `/main` to either URI fails 15 tests.

Also verified: `w14:paraId` is 4-byte hex on `CT_P`; `w:author` required (max 255), `w:initials` max 9; the comment markers are `CT_MarkupRange`/`CT_Markup`, so they take **the same `@w:id` union as bookmarks** (-1 excluded).

**Left unverified, deliberately:** whether Word *requires* a `commentReference` to display a comment (the schema makes all three markers optional, so the problem is worded as likely non-display, not a violation); whether `w15:paraIdParent` names the thread root or the immediate parent (`commentThreads` walks transitively so both readings agree on membership); and whether Word regenerates a missing `commentsExtended.xml` or discards threading on save.

## 8p. Pivot tables — the three-hop chain (2026-08-18)

`services/excelPivotTables.ts` + 41 tests, `d99046d`. Task #28 closed. Built by a subagent.

**The chain**, and the point of the module is naming *which hop* broke, because none of them is visible from the pivot table part:

```
pivotTable1.xml @cacheId → workbook.xml pivotCache @r:id → cacheDefinition @r:id → cacheRecords
```

**`@cacheId` and `@r:id` are different identifier spaces on the same element**, both required — verified: `CT_PivotCache` declares both with `RequiredValidator`. Conflating them is a real bug and one mutant proves the test catches it.

**Two schema facts that changed the design:**
- `CT_PivotCacheDefinition`'s `@r:id` has **no** RequiredValidator, so **absent cache records are not a schema violation** — a cache set to refresh on load legitimately has none. That is why `PivotProblem` carries a `severity: 'error' | 'note'` the other modules don't need: it keeps a legitimate state off the error list instead of forcing a choice between calling it damage and dropping it.
- **`CT_Worksheet`'s content model contains no pivot child at all**, so worksheet→pivotTable is a purely implicit relationship — nothing in the worksheet XML names it. Hence orphan detection: a pivot part under `xl/pivotTables/` that no relationship points at.

✅ **The "67 variations" figure is now VERIFIED — and the claim attached to it was wrong.** Settled 2026-08-19 by counting clause-keyed entries in [MS-OI29500]'s **public table of contents** (`toc.json`, 1,892 clause-keyed entries). Only headings and clause references were read, never the prose — the distinction `LICENSING.md` §5 draws between "a dataset of URLs and headings" and Microsoft's expression.

- **Part 1 §18.10 (pivot tables): 67.** Exactly as claimed.
- 🔴 **"Second only to formulas" is false.** §18.10 is **third**: §18.17 formulas **218**, §18.3 worksheet **83**, then §18.10 at 67. Formulas deviate more than three times as much. The exaggerated version invites the wrong triage, so the module now states the ranking explicitly.

Corrected in `services/excelPivotTables.ts`; the pivot analyzer's `cannotDetermine` now says the *count* is settled while the per-file impact is not.

**Distribution cross-check** against the numbers recorded earlier in this file — near-identical, with drift of at most one entry, so it is the same corpus at a slightly different revision. Recorded rather than silently overwritten: Part 1 §18 **534** (was 534), §17 **528** (was 529), §21 **264** (264), §20 **179** (was 180), §22 **106** (106), Part 4 §19 **104** (104), Part 1 §19 **100** (100). Total clause-keyed **1,892** (was 1,894/1,895).

Also flagged not-verified: that `pivotFields` runs parallel to `cacheFields` (nothing in the schema ties them; the cache field count is used as the bound for both, and says so), and `@x = -2` as the "values" pseudo-field (`@x` is a plain Int32 with no facet — it is convention, excluded from range checks so multi-measure pivots do not false-positive).

## 8q. ISO Strict was decorative in `oleObjects.ts` — fixed (2026-08-18)

`76b19ab`. Found by the pivot agent while following that file as a pattern.

**Strict does not merely swap the host — it drops the year, which sits in the middle of the URI:**

```
Transitional  http://schemas.openxmlformats.org/spreadsheetml/2006/main
Strict        http://purl.oclc.org/ooxml/spreadsheetml/main
```

Matching a `/spreadsheetml/2006/main` suffix therefore matches **Transitional only, while looking Strict-tolerant**. Every Strict document, workbook and deck reported zero OLE objects — the exact silent-failure shape the module exists to catch, in the module itself.

🔴 **The test that should have caught it asserted a URI that does not exist** (`purl.oclc.org/ooxml/wordprocessingml/2006/main`), so it passed against the broken matcher and certified support the code never had. Replaced with three real Strict cases across all three formats; reverting the matcher now fails all three. The helper takes a *vocabulary name* and builds both spellings itself, so a caller can no longer pin one form by accident.

**Known gap this exposes:** the Word and Excel resolvers (`wordStyleResolver`, `wordBookmarks`, `excelStyleResolver`, …) match element namespaces against **exact Transitional constants**, so they read nothing out of Strict packages either. Unlike `oleObjects` they never *claimed* otherwise, so this is a limitation rather than a false claim — but it is repo-wide and undocumented. Relationship-type matching is fine everywhere (it matches the trailing segment, which Strict preserves).

## 8r. The spine — diff grounded, one Finding type, nothing orphaned (2026-08-18)

An audit before answering "where is this converging" found the real problem was not too many features but **almost none of them connected**: 16 analyzers, ~10,000 lines, 6 reachable from the UI and only on the *explain selected tag* path. Three had no caller at all — including `ooxmlDiff.ts`, 489 lines built for the user's most-wanted case.

**The reframe.** This is not a RAG system. It parses parts, resolves inheritance, walks relationships and emits diagnostics with severities and remediations — it is **a compiler front-end for OOXML** with a model as a narration layer. Reference architecture is ESLint or a language server: a rule registry, one diagnostic type, several consumers. The deliverable is the **engine**; the panel is its first consumer. (User confirmed: *"based on those we can build further agents for our use-cases."*)

⚠️ **Findings are not embeddings, and the distinction is load-bearing.** A finding is a claim with a traceable origin — this code, in this part, about this subject — and that traceability is what computes the Verified badge. A vector has no origin to name. Compressing findings into an embedding space keeps the size saving and throws away the property the whole design rests on. Embeddings may still belong on the *prose corpus* for fuzzy retrieval (§8h, task #21); never on the findings.

**Shipped:**

1. **The diff is grounded** (`829c2eb`). `diffPackages` now runs before the model sees anything, marked authoritative on both provider paths with an explicit instruction not to report differences it does not list. Parts are pivoted into a before-package and an after-package and diffed *together*, so content moved between parts is visible. `unresolved` travels with it, so a partial derivation caps the tier.

2. **One `Finding` type** (`cb26b8d`), in `services/findings.ts`: `code`, `severity`, `part`, `message`, `remediation`, `silent`, `subject`. Codes are namespaced `analyzer/kind` — necessary, not cosmetic: **bookmarks and comments both emit `duplicate-id`**. Prose is now *rendered* from findings (`renderFinding` is the only place a sentence is built) rather than authored as strings, which is what makes the output consumable by another agent.

3. **Severity and silence became explicit judgements.** Moving them into per-analyzer rule tables forced the call once per kind. The finding: **most integrity and bookmark faults are silent** — the package opens, the page renders, navigation is broken. Pivot keeps its severity per-call-site because there it genuinely depends on the occurrence (absent cache records are an `error` normally, a `note` under `refreshOnLoad`).

4. **Nothing is orphaned.** `wordComments` and `excelPivotTables` gained evidence functions and `ANALYSIS_TARGETS` entries. Every analyzer now has a caller.

**Method note.** The cross-analyzer contract tests run all five analyzers over deliberately broken input and assert every record is namespaced, names its part, carries a fix, and survives a JSON round-trip — a shared type is only worth having if nothing quietly opts out. Nine mutants; **one escaped and found a real gap**: the silent-ordering test used codes that already sorted correctly alphabetically, so it passed with the tiebreak deleted.

## 8s. The registry — Phase 3 complete (2026-08-18/19)

`services/analyzers.ts` + `services/coverageGaps.ts`, `402fdf8` and `fab250b`. **799 tests.**

**Two surfaces were running a fraction of the engine, for the same reason** — there was no way to say *"run everything that applies here"*:

- **Validate ran ONE analyzer.** `ValidatorView` imported `checkPackageIntegrity` and nothing else, so *"is this file correct?"* checked content types and relationships while fifteen analyzers sat unused.
- **Compare ran the structural diff and no analyzer at all**, so a file that had lost an OLE embedding or a bookmark's end marker reported the part change without reporting that anything was broken.

**`ANALYSIS_TARGETS` is gone.** The set of markup the engine could explain lived in a table inside a React component — the UI layer defined the engine's capabilities and nothing else could see them. Routing is now registry-driven; the four resolver-only modules (Word cascade, Excel formats, PowerPoint inheritance, charts) are registry entries too.

**Analyzers may validate, explain, or both** — different questions: `analyze` finds what is *wrong*; `explain` describes what a selected element *is* and how it resolves. An explain-only analyzer counts as **skipped** by the validate pass, so the ledger never overstates what was checked.

**The capability ledger.** Every entry declares `cannotDetermine` beside `determines`, and the ledger is computed from the registry — never asserted by a model. The validate surface now says which checks ran, which were skipped, and what those checks cannot see; its clean-run message is narrowed to **"no problems found by the checks that ran"**, not "this file is fine". This extends the honesty property from per-fact to **per-capability**.

**`diffFindings`** answers what a change did to the *health* of a package, which no structural diff can: a dropped embedding is one removed part and looks like any other, but as a finding it says the document still renders correctly and is broken anyway. Identity is **code + part + subject, never the message** — messages interpolate counts and paths, so including them reports the same fault as both fixed and reintroduced. Pre-existing faults are reported as `unchanged` so nothing is blamed on the wrong change.

**The gap log** (`coverageGaps.ts`) is the "keeps improving" half, deliberately the boring half. When a part is opened with no analyzer behind it, that is recorded — a backlog ranked by real usage instead of guesswork. **Counts only, over spec vocabulary alone**: part paths normalised (`word/header3.xml` → `word/header#.xml`, so one gap does not become fifty), element names are ECMA-376 tags. Same DLP discipline as `retrievalMetrics`, and surfaced beside it.

🔴 **What "self-improving" does NOT mean here, recorded so it is not re-proposed:** a model that writes its own analyzers, grades its own output, or learns from conversations. Each puts a model back on the trust path — the reason fine-tuning was ruled out — and a rule the model invented cannot produce a Verified badge, because nothing verified it. The loop closes **through a person**.

**Method note.** `explainersFor`, `siblingsFor`, `explainPart` and `analyzePackage` all take an optional registry, added purely for testability: mutants survived because nothing could inject a stub — no test could force an analyzer to throw, produce a duplicate line, or match its own part as a sibling. **PowerPoint's sibling pattern does match `ppt/slides/slide1.xml` itself**, so that guard is load-bearing, not defensive. Across the two commits, 23 mutants and 7 real gaps found.

## 8t. The report — the engine's output as a document (2026-08-19)

`services/report.ts` + 14 tests, `3c02d3b`. **813 tests.**

The engine's output could previously only be read as a log, which no other program can consume — a problem, given the stated direction that **the engine is the product**.

`reportPackage` / `reportComparison` emit a structured report: sorted findings, counts by severity, how many are **silent**, and the capability ledger travelling alongside so a consumer knows what was checked and what could not be. Exported from the validator as JSON; with both files loaded it exports the comparison instead.

**`schemaVersion` makes it a contract**, with the rules written into the module header: adding a field or a new finding code is MINOR; renaming or repurposing one is MAJOR; **finding codes are part of the contract**. `generatedAt` is the only non-deterministic field, so two runs over one file are byte-identical apart from the timestamp — which is what makes a report diffable in CI.

**Deliberately absent: the explain path.** It describes how markup *resolves* rather than what is wrong with it, and forcing that into findings would mean inventing a fault for every fact. When an agent needs it, it wants a different shape; inventing that shape before there is a caller would be guessing.

## 8u. Which features deserve an analyzer — the criterion, and a measured backlog (2026-08-19)

Asked directly: do shapes, tables, lists, audio/video each need their own analyzer? **No — and the answer is not "it depends".** There is a test.

### The test: does this feature fail INVISIBLY?

Every analyzer that has earned its keep detects something a person looking at the document cannot see. That is the entire pattern:

| Analyzer | What renders fine while broken |
|---|---|
| OLE | preview image intact, embedding gone |
| bookmarks | text all present, every cross-reference target lost |
| comments | margin looks normal, replies flattened to top-level |
| pivots | cells show last-refresh values, chain severed |
| package | file opens, implicit relationship missing |

A feature whose faults are *visible* does not need one. A wrong shape looks wrong; the user reports it without help.

### Three signals that predict invisible failure

1. **Indirection** — resolves through `r:id` or an implicit relationship. Break the link and a fallback renders. *(OLE, media, images, pivots, charts, slide layouts.)*
2. **Paired markers** — two elements matched by id where losing one is silent. *(bookmarks, comments, revisions, permissions, moveFrom/To ranges.)*
3. **Cached or duplicated state** — a stored copy that can drift from its source. *(pivot cache, chart series cache, field results, sharedStrings.)*

If a feature hits none of these, it probably wants better *explanation*, not an analyzer.

### Measured backlog — [MS-OI29500] deviations per clause

Counted from the public TOC (§8p). **Deviation count says where Office and the spec disagree most, i.e. where "valid per spec" ≠ "renders right in Word."** Cross it with the invisibility test:

| Clause | Deviations | Covered? | Invisible failure? |
|---|---:|---|---|
| §18.17 formulas | **218** | ✗ | partly — a stale cached result is invisible |
| §21.2 charts | 172 | partial | cache-only series ✓ |
| §20.1 DrawingML framework | 149 | ✗ | mostly visible |
| §18.3 worksheet | 83 | partial | — |
| §17.3 paragraphs | 77 | ✓ cascade | — |
| §22.1 math | 75 | ✗ | mostly visible |
| **§17.16 fields & hyperlinks** | **72** | ✗ | ✓✓ **strong** — a REF to a lost bookmark shows a stale cached result |
| **§17.4 tables** | **71** | styling only | ✓ Word silently repairs a broken grid |
| §17.15 settings | 70 | ✗ | ✓ but low impact |
| §18.10 pivots | 67 | ✓ | — |
| §17.13 annotations | 46 | ✓ | — |

### Verdict on the four asked about

- **Word tables** — ✅ yes, but *not the styling*: `wordTableStyles.ts` already covers banding and `tblLook`. The gap is the **grid model** — `gridSpan` sums that do not match `tblGrid`, and `vMerge` continuations orphaned with no `restart` above them. Word repairs these silently and renders something plausible; another renderer mangles the layout. **Being built.**
- **PPT audio/video** — ✅ yes, and it is the OLE pattern exactly: the poster frame renders while the media is gone, and `r:link` media lives *outside* the package so a deck that plays on the author's machine is silently not self-contained. Low deviation count, high invisibility. **Being built.**
- **Word lists** — ⚠️ mostly covered. `wordNumbering.ts` already handles the three patterns and the `numStyleLink` double-hop. Marginal gain.
- **Word shapes** — ⚠️ biggest blind spot by deviation count (199 across §20.1 + §21.1) but **mostly visible**, so it ranks lower than the count suggests. The invisible parts are narrow: text inside `wps:txbx` that extraction misses, and dangling `a:blip/@r:embed`.

**Next by this ranking: fields (§17.16).** 72 deviations, and it fails invisibly in the strongest way — a `REF` field caches its last-computed result, so a cross-reference to a deleted bookmark keeps displaying the old text indefinitely. It also interlocks with the bookmark analyzer already built.

## 8v. ISO Strict, properly — normalised at one choke point (2026-08-19)

`services/conformance.ts` + 17 tests, `93b3be2`. §8q fixed one module; this fixes the class.

**Every analyzer compared namespaces by exact equality against a Transitional constant**, so a Strict package produced no findings at all — not an error, just silence. `oleObjects` had been caught doing this while *looking* Strict-tolerant; the other thirteen modules had the same blind spot and had simply never claimed otherwise.

**Fixed at one choke point rather than sixty call sites.** ~60 namespace comparisons across 14 modules funnel through 6 constants. Teaching each of them both spellings would be 60 chances to get one wrong plus a rule every future analyzer author must remember. Mapping the URIs once, before any analyzer sees the markup, keeps the analyzers exact and makes Strict support a property of the pipeline. Rewriting is text-level, not tree-level: a DOM's `namespaceURI` is read-only, and rewriting the base URI fixes relationship `Type` attributes for free since a type is that URI plus a trailing segment.

⚠️ **It is not a converter.** Namespace mapping makes Strict markup *readable* by Transitional-shaped code; it does not reconcile the classes. **Strict forbids VML**, so an embedded object there carries a DrawingML preview rather than the `v:imagedata` one the OLE check looks for — the `conformance` analyzer reports Strict as a **note** (it is valid, usually deliberate) and states that limit, because a clean report would otherwise imply coverage we do not have.

`Analyzer` gained **`readsRawMarkup`** for this one analyzer — found by a test: normalising first left it permanently convinced every package was Transitional.

🔴 **Provenance, and a bug caught by checking it.** Strict URIs come from `pjfanning/ooxml-strict-converter`; **every Transitional target was verified independently against the SDK's `namespaces.json`**. That check found an error in *their* mapping file — it maps wordprocessingml to `.../wordprocessingml/main`, missing the `/2006`. Copying it wholesale would have left every Strict Word document unreadable. A test pins the correct value.

**Honest note on ordering:** longest-key-first replacement is *insurance*, not a live fix — no pair in the current table needs it, because `chart` and `chartDrawing` both gain the same `/2006`. Pinned as an invariant with the reasoning recorded, and a test comment claiming otherwise was corrected.

## 8w. Fields — the text on the page is a cache, and it can lie (2026-08-19)

`services/wordFields.ts` + 34 tests, registered as the `field` analyzer. Top of the §8u backlog: **72 deviations for Part 1 §17.16**, and the strongest invisible failure in the format.

**The headline.** A field is a small program (`REF`, `PAGEREF`, `TOC`, `HYPERLINK`, `PAGE`) whose displayed text is **the result from the last time Word ran it**, stored in the file. So:

> Delete the bookmark a cross-reference points at, and the cross-reference keeps displaying the old text — correctly formatted, indefinitely, until someone presses F9.

Text extraction returns the stale value. A converter copies it. A reviewer proofreads it. Nothing anywhere flags it. This is how *"see section 4.2"* survives in documents whose section 4.2 was deleted years earlier.

**Three states, and only one is safe:**
- `@w:dirty="true"` — Word recalculates on open, so a stale result is temporary. But **anything reading the file without evaluating fields is reading a value Word has already declared out of date.**
- neither flag — **the dangerous case.** The cached text is presented as current and will not recalculate on its own.
- `@w:fldLock="true"` — **worse.** The field will not update even on F9; the stale value is permanent by instruction. Escalated to its own code.

**Why it needed the bookmark analyzer.** The field says which bookmark it targets; only the bookmark index knows whether that bookmark still exists. `crossCheckFieldTargets` is the first finding in this codebase that **neither analyzer could reach alone** — the argument for the registry, demonstrated rather than asserted.

**Traps encoded:**
- **Fields nest** — a `TOC` holds a `PAGEREF` per entry *inside its own result*. Pairing the first `begin` with the first `end` grabs the inner field and mis-reads everything after it, so the walk keeps a stack. A nested field's result is also part of its parent's, because that is what a reader sees.
- **`separate` is optional.** A field that has never been calculated goes `begin → instruction → end`. Treating its absence as a fault would put noise on every clean document. `cachedResult` is `null` for that and `''` for a field that ran and produced nothing.
- **False positives were the real design risk.** `HYPERLINK` targets a bookmark only with `\l`; without it the argument is a URL. `STYLEREF` takes a style name. Reporting either as a dead bookmark would make the report unreadable on a real document, so both are tested explicitly.

Verified against the SDK schema: `w:fldChar/@w:fldCharType` required, exactly `begin|separate|end`; `w:fldSimple/@w:instr` required; both `w:fldChar` and `w:instrText` are children of `w:r`. There are **two** `fldSimple` declarations — `CT_SimpleField` and `CT_SimpleFieldRuby` — which is why matching is on element name, not parent.

Ten mutants; one escaped and found a real gap: no fixture had `instrText` *after* a `separate`, so the guard stopping a field's result leaking into its own instruction could be deleted unnoticed.

## 8x. Table geometry and PowerPoint media (2026-08-19) — and a salvage lesson

Both built by subagents, **both killed mid-task by the session limit**, both recovered. `8448f27`. **1005 tests.**

**`services/wordTableGrid.ts`** — a row whose `gridSpan` values do not sum to the declared `tblGrid`, and `vMerge` continuations with no `restart` above them in the same column. **Word silently repairs a broken table and renders something plausible**, so the document looks fine and a renderer in another environment mangles the layout. 71 documented deviations for Part 1 §17.4. Note `w:vMerge` with **no `w:val` at all means *continue*, not restart** — the opposite of most OOXML defaults, and the trap that makes orphan detection worth having.

**`services/pptMedia.ts`** — the OLE pattern again (the poster frame renders whether or not the media is behind it) plus the case that matters more in practice: **`r:link` media lives outside the package**, so a deck that plays on the author's machine is silently not self-contained. PowerPoint links video by default above a size threshold, so this is common and almost never deliberate.

### 🔴 Process lesson — worktree work is not safe until it is committed

The brief for both agents said *"commit as soon as you have something green rather than saving it all for the end."* One followed it and had a commit to cherry-pick. The other did not, and **57 KB of finished, passing work was sitting uncommitted in a worktree** that would have been discarded with it.

**Recovery is only possible because worktrees survive the agent.** Before removing any agent worktree, always check `git -C <worktree> status --short` for uncommitted files as well as `git log` for commits — the second agent's work was invisible to a log check alone.

**Strengthen future briefs:** tell the agent to commit after the *first* green test run, not after the module is finished, and to keep committing. An agent that treats its commit as a final deliverable is one quota reset away from losing everything.

## 8y. Formulas — the biggest divergence cluster in the format (2026-08-19)

`services/excelFormulas.ts` + 32 tests, registered as the `formula` analyzer. **218 normative variations against Part 1 §18.17 — three times the next-largest SpreadsheetML clause**, and the top of the §8u backlog.

**Same failure class as Word fields, at spreadsheet scale.** A cell stores the program *and* the answer:

```
<c r="B2" t="n"><f>SUM(A1:A10)</f><v>55</v></c>
        the program ──┘              └── the answer, from the last recalculation
```

**Every reader without a calculation engine shows `<v>`** — converters, extractors, dashboards, `openpyxl` by default, and this tool. So a workbook can display numbers its own formulas would no longer produce, and nothing looks wrong.

**Excel signals its own doubt, and those signals are the most useful thing here:**
- `calcPr/@fullCalcOnLoad="1"` — Excel will recalculate on open, so the stored values are **stale by Excel's own declaration**. Anything reading them without calculating is reading numbers Excel has already disowned.
- `calcPr/@calcMode="manual"` — values drift from formulas *by design*, and Excel will not correct them until someone presses F9.
- `calcPr/@calcCompleted="0"` — the last pass did not finish, so values may not be self-consistent.

**The shared-formula trap, which silently loses a formula entirely.** Excel compresses a filled-down column by writing the formula once: the master carries `ref` + `si` + the text, and every follower carries **`si` and nothing else** — its formula exists only as an offset from the master. Delete or fail to write the master (what happens whenever a tool rewrites rows without understanding `si`) and every follower becomes a cell with a cached number and no way to recompute it. Excel repairs it quietly on open; other readers see an empty formula element.

⚠️ **Nothing here evaluates a formula, and the module says so.** It reports what is stored and what is missing; it never claims a cached value is numerically wrong. *"This may be stale"* is supportable; *"this is 55 but should be 60"* would need an engine this does not have. That limit is in `cannotDetermine` and in the panel's `unresolved`.

Verified against the SDK schema: `x:f/@t` is `Normal|Array|DataTable|Shared` and **optional** (absent = normal); `@si` is UInt32; `x:c/@t` is `Boolean|Number|Error|SharedString|String|InlineString|Date`; `x:calcPr` declares `@calcId`, `@calcMode`, `@fullCalcOnLoad`, `@calcCompleted`, `@calcOnSave`, `@forceFullCalc`.

Eight mutants; two escaped and both were instructive. One line turned out to be **provably redundant** (a "skip masters" guard the `si` lookup already covers) and was removed rather than left looking like protection. The other exposed an untested guard: the error check keys on `x:c/@t="e"` as well as the value text, and nothing tested a **string** whose text happens to be `#N/A` — which `IFERROR(x,"#N/A")` produces legitimately. Without the guard that working formula reports as broken.

## 8z. Content controls and hyperlinks (2026-08-19)

**`services/wordContentControls.ts`** + 28 tests. The mechanism behind every template-driven document pipeline, and it fails in a way nobody can see: **a control's content is stored in the file**, so if the binding breaks the content stays. A document whose bindings point at a custom XML part that is gone opens showing the values from the last time it worked. For document generation that is the difference between *"the template populated"* and *"the template printed last month's numbers again"*.

`w:showingPlcHdr` is the second signal — the control is displaying prompt text, not data. Correct in a template; in a supposedly generated document it means the field was never filled.

Store item ids are matched **without braces and case-insensitively**, because generators are inconsistent about both and a binding differing only in case is not broken. ⚠️ **The XPath is deliberately not evaluated** — that needs a namespace-aware engine plus the binding's `prefixMappings`, and getting it subtly wrong would produce confident false reports about working templates. The module checks the *part* exists and says the expression is unchecked.

**`services/hyperlinks.ts`** + 64 tests, built by a subagent. Format-agnostic, and it **corrected the brief again**: 🔴 **`CT_HyperlinkOnClick` does not exist.** The type is `a:CT_Hyperlink` (SDK class `HyperlinkType`), shared by `a:hlinkClick`, `a:hlinkMouseOver` and `a:hlinkHover`, with **no required attributes** — so an action-only link with no relationship is legal. Also found a second `w:hyperlink` schema type (`CT_HyperlinkRuby`) carrying an identical attribute set, and confirmed `x:hyperlink/@ref` is **required** while everything on the Word and DrawingML forms is optional.

External URLs are **never fetched** and produce no finding at all — they go to `unresolved`, and the three-state accessor returns `null`. The `ppaction://` verb vocabulary is explicitly **not verified**: `@action` is an untyped string with no schema enumeration, so the table was cross-checked against python-pptx and LibreOffice rather than against ECMA-376, and unknown verbs are reported verbatim and never judged.

**Its mutation report is the best example yet of the pattern in [[ooxml-explorer-mutation-testing]]:** 22 mutants, three survived, and **two were tests passing for the wrong reason** — a column-arithmetic test whose fixture would have passed under any base, and an external-link test whose verdict was decided by an earlier branch rather than the rule it aimed at. The third survivor was a **bad mutant**, not a gap, and the agent said so rather than inventing a test for it.

## 8aa. Charts made to report, and tracked changes (2026-08-19/20)

**Charts now contribute findings.** A registry audit found something worth remembering as a pattern: **four analyzers were `explain`-only**, and for three of them that is correct (they describe *resolution*, not faults). But **charts — second-largest divergence cluster in the format at 172 variations — contributed nothing to validation or comparison**, despite `readChart` already computing rich `problems` and `translationNotes`. They were plain strings that only ever reached the explain path.

Added `chartFindings`, plus the check the model never made: `c:externalData/@r:id` names the embedded workbook a chart opens when you double-click it. **Lose that part and the chart still draws perfectly from its cache and can never be edited or refreshed again** — the OLE preview problem wearing a different hat. A chart whose series reference ranges while carrying no workbook at all gets its own finding: the formulas name cells that exist nowhere in the package.

Two codes here are deliberately **coarse** (`chart/structural-problem`, `chart/translation-risk`), each carrying many generated messages. That is the ESLint precedent — many messages under one rule id — and it is recorded so nobody later reads it as sloppiness.

⚠️ **A test caught a real bug in that change**: `s.values?.formula` yields `undefined` when a series has no values element, and **`undefined !== null` is true**, so the naive spelling reported every chart with a *missing* data source as reference-based.

**`services/wordRevisions.ts`** + 62 tests, built by a subagent. A document with unaccepted changes has **two different texts**, and anything extracting text picks one without saying which. `compareRevisionOutcomes` returns accepted, rejected, and `naive` — what a tag-blind extractor produces, which with both an insertion and a deletion present **matches neither reading**.

🔴 **The brief was wrong again, and the schema said so: `@w:date` is NOT required** on `CT_TrackChange` or any `*Change` type. It *is* required on `CT_MoveBookmark`. The module reports these as two findings at two severities rather than flattening them. Also corrected: **`w:tblGridChange` declares only `@w:id`** — no author, no date — so without an exemption every valid table-grid edit reports a fault. And move *starts* are `CT_MoveBookmark` with required `@w:name` while move *ends* are `CT_MarkupRange` with **no name at all**, so starts pair to ends by id and the two halves pair to each other by name.

**Its mutation round found a real defect, not just a test gap:** chasing a surviving mutant through what looked like dead code surfaced that text nested inside `w:pPr/w:rPr` (out of schema) was being scored as document content, so **a malformed paragraph mark could inject words into the *accepted* reading**. 25 mutants, three survivors, all three tests decided by an accident of the fixture rather than by the rule they targeted.

## 8ab. ✅ RESOLVED — the salvage worked (kept as a record of the recovery)

Both modules are now tested, fixed and registered; see §8ad. What follows is the state as it was, kept because the recovery procedure is worth reusing.

**Historical:** Both subagents were killed by the quota reset **after writing their modules but before writing a single test**. The modules were recovered from their worktrees (39 KB and 32 KB) and committed, because losing them would mean re-deriving all the schema research. They **typecheck and lint clean**.

**They are NOT registered in `services/analyzers.ts`, and must not be until they are tested.** An untested analyzer that reports findings under a Verified tier is exactly the failure this project exists to prevent.

**What each still needs:**
1. `tests/excelExternalLinks.test.ts` and `tests/pptAnimation.test.ts`.
2. **Mutation testing** — non-negotiable here, since neither module has ever had a test run against it.
3. `pptAnimation.ts` is **incomplete**: `computeAnimationEvidenceForMarkup` was cut off mid-function and is marked with a ⚠️ comment at the exact line. Its `explain` entry must not be registered until that function actually builds its lines. The `analyze` path may be complete — check before trusting it.
4. Registration in `analyzers.ts` with `determines` / `cannotDetermine`.

The original briefs follow, since the modules should be read against them rather than assumed correct.

**`services/excelExternalLinks.ts`** — the other half of the formula analyzer. A workbook linking to another workbook **caches the values it last read**; open it where the source is unreachable and every linked cell shows numbers from an unknown date. Chain: `xl/workbook.xml` `<externalReferences><externalReference r:id>` → `xl/externalLinks/externalLink<N>.xml` `<externalBook r:id>` with `<sheetDataSet>` cached values → `TargetMode="External"` to the real workbook. **The `[1]` in a formula is a 1-based index into the `externalReferences` list** — verify that. Report the external target as *unverifiable*, never broken; never fetch. Check: reference with a missing relationship or absent part; an `externalBook` whose own `r:id` does not resolve (so even the path is unknown); a formula index with no matching reference; externalLink parts nothing references.

**`services/pptAnimation.ts`** — an animation targets a shape **by id**, and deleting the shape makes it **silently never fire**. `p:sld/p:timing/p:tnLst` holds time nodes (`p:par`, `p:seq`, `p:anim`, `p:animEffect`, `p:animMotion`, `p:set`); a node targets through `p:cTn/p:tgtEl/p:spTgt/@spid`, which must match a `p:cNvPr/@id` in the slide's shape tree. Check: `spid` naming a missing shape (the headline); `p:bldP`/`p:bldLst` `@spid` likewise; `p:cond/@evt` on a missing target; `p:txEl`/`p:pRg` selecting paragraph indices the shape does not have; duplicate `p:cNvPr/@id` making any id targeting ambiguous. Report counts — *"3 of 11 animations will never run"* is the sentence someone needs.

Both briefs said: base on `feat/schema-derived-rag-corpus`, touch only their two files, do not edit `analyzers.ts` (registration is done here), verify names against the SDK schema, and **commit after the first green test run**.

## 8ac. Dangling style and format references (2026-08-20)

`services/styleReferences.ts` + 22 tests, registered as `styleRef`. **The top remaining gap, verified as uncovered before building**: nothing reported a `pStyle` naming a missing style, nor a cell `@s` past the end of `cellXfs`.

**Why it is the highest-value one left.** Every formatting system in OOXML is a table plus references into it, and **every one falls back silently**:

- `w:pStyle` naming an undefined style → **Word applies Normal.** A paragraph meant to be Heading 1 renders as body text. Right font, right size, wrong meaning, file opens cleanly.
- A cell `@s` past `cellXfs` → default format.
- A `@numFmtId` with no `numFmt` → General, so a currency column renders as bare numbers.

None of it looks like breakage. It looks like **plainer formatting**, which reads as a design choice — which is exactly why it survives review. It is also the most common defect in *generated* documents, because a generator that writes a reference and forgets the definition produces a file that passes every structural check ever written for it.

**Why a separate analyzer rather than folding into the resolvers.** The resolvers answer *"what does this element resolve to?"* — one element, asked when someone selects it. This asks *"does every reference in the package land?"* — the whole file, asked when nobody is looking at anything in particular. The resolvers were `explain`-only and contributed no findings; this is the missing half.

**Two false-positive traps, both tested:**
- ⚠️ **`w:numId="0"` is NOT a dangling reference** — it means *remove numbering*. Treating it as a lookup fires on every document that has ever had a list removed, which is most of them.
- ⚠️ **Excel number formats 0–163 are built in and declared nowhere.** Only **164+** must appear in `numFmts`; checking all of them reports every ordinary workbook as broken.

One finding per missing style, not per use — forty paragraphs referencing one broken style is one broken style.

Seven mutants; one escaped and it was the classic: the "last valid index" test used a value far from the boundary, so a `<` vs `<=` bound was invisible. Fixed with the index exactly one past the end.

## 8ad. Footnotes, animations, external links — and my recommendation was wrong (2026-08-20)

**`services/wordNotes.ts`** + 16 tests. A reference mark with nothing behind it: the superscript still renders, the reader sees a footnote marker, and there is no note at the bottom of the page. Extraction returns the mark and drops the content.

⚠️ **The trap that would fire on 100% of real documents.** `footnotes.xml` *always* contains `separator` and `continuationSeparator` notes that **nothing in the body references** — they are referenced from `sectPr`, and Word writes them into every document that has ever had a footnote. A naive *"every note must be referenced"* check reports two false positives on every real file. Only `w:type="normal"` (or absent, the default) is a content note. Same family as `numId="0"` and Excel's built-in number formats: **the highest-value work in these analyzers is often knowing what NOT to report.**

**`pptAnimation.ts` and `excelExternalLinks.ts` — the salvaged modules, now finished.** Both agents found **real defects in the code they inherited**, and both defects were the same species:

- `p:pRg` with no `@st` was read as paragraph 0, because **`Number(null)` is 0** — producing confident findings about a range the markup never stated.
- A cached sheet with an empty `@sheetId` was attributed to the *first* sheet of the source workbook, because **`Number('')` is 0**.

🔴 **The animation agent found §8ab's own brief was wrong**: `p:cTn` has six children and `p:tgtEl` is not among them. Behaviours target through **`p:cBhvr/p:tgtEl`**, and there is no `p:condLst` — conditions live in `p:stCondLst`/`p:endCondLst`/etc. A module written from that brief would have found nothing on every real deck. **That is four briefs of mine corrected by agents now.**

### 🔴 Correcting a recommendation I made and got wrong

I suggested letting the **gap log** decide which analyzer to build next. **That was wrong, and I verified why:** the gap log fires only when a part is opened that **no analyzer matches**. `word/footnotes.xml` is matched by eight analyzers, so opening it never records a gap — the log would have stayed silent forever about a missing footnote check.

**The gap log measures PART coverage, not CHECK coverage.** It is the right tool for discovering whole part types nobody anticipated (`customXml`, `docProps`, ink, embedded fonts). It cannot discover a missing check inside a part already covered.

**So: build the remaining ones deliberately, using the §8u criterion.** The stopping rule is that criterion, not the log — stop when the marginal analyzer's failure mode is *visible* (DrawingML geometry, math, VML), because a wrong shape looks wrong and the user reports it without help.

**Remaining by that criterion:** Excel defined names (`#REF!` propagation), Excel tables/ListObjects, document settings (§17.15, 70 deviations), data validation and conditional formatting.

## 8ae. A skill for adding the next analyzer (2026-08-20)

`.agents/skills/add-analyzer/` — `SKILL.md` plus `mutate.py`. Encodes the procedure that
produced the 21 existing analyzers, because every step of it exists to prevent a specific
defect that actually happened here.

**What it encodes that a person would otherwise re-derive:**
- **Step 1 is a gate, not a formality.** The invisible-failure criterion (§8u) with the
  three signals, plus the [MS-OI29500] deviation-count script. **"Do not build this" is a
  documented legitimate outcome** — it is how DrawingML, math and VML got correctly skipped
  despite topping the deviation table.
- **Verify the schema before writing.** With the exact SDK commands. Four briefs on this
  project asserted element paths that do not exist and every one was caught here.
- **The false-positive trap table.** `numId="0"`, footnote separators, Excel `numFmtId`
  < 164, `HYPERLINK` without `\l`, `STYLEREF`, absent `separate`, shared-formula masters,
  Strict dropping the year — plus the two JavaScript traps that caused real defects,
  **`Number(null) === 0`** and **`undefined !== null`**.
- **Mutation testing, with the failure modes named.** Suspect the *test* first: a fixture
  that passes either way, an assertion decided by an earlier branch, an accidental
  tie-break, a provably unreachable guard. Equivalent mutants get recorded, not papered
  over with an invented test.

**`mutate.py`** replaces the bash loop that was hand-written about a dozen times. It
refuses ambiguous or absent patterns rather than mutating the wrong line, restores the
source in a `finally` so a crash cannot leave a mutated file staged, and exits non-zero
when a mutant survives so it can gate a commit. Verified end to end against `wordNotes`:
a real mutation killed 4 tests, a deliberate no-op correctly survived, a bogus pattern was
skipped, and the source came back byte-identical.

⚠️ **Step 5 is deliberately gated and honest.** The user asked for "passing marks from
real-world files"; **this project has no real-file corpus** — every test runs against
hand-written fixtures, and acquiring one is blocked on the licensing question (task #11).
The skill runs against `tests/fixtures/` if it exists and otherwise **says so and records
the analyzer as fixture-verified only**. Building a step that silently passed would have
manufactured exactly the false confidence the whole design exists to prevent.

## 9. Next actions

Stages 0–2 are complete for all three formats and the Verified tier is live. Remaining work is per-subsystem, tracked as tasks #11–#28.

All four subsystems the user named are **DONE**: #26 bookmarks (§8m), #27 OLE (§8n), #25 comments (§8o), #28 pivot tables (§8p).

**All three phases of the §8r plan are done.** The architecture work is finished; what remains is content and cleanup.

1. ~~Source or cut the "67 variations" figure~~ — **DONE**, see §8p. Verified at 67; the "second only to formulas" half was false and is corrected.
2. ~~Carry `Finding[]` to the panel boundary~~ — **DONE for validate and compare**, see §8t. Those paths produce findings natively and now export as JSON. The *explain* path is still prose by design; revisit only when a caller actually needs it. `ComputedEvidence` still lives in `aiService` and is type-imported by `geminiService` to avoid a runtime cycle — a small seam worth tidying if that type ever grows.
3. ~~Decide on Strict support repo-wide~~ — **DONE**, see §8v. Normalised at one choke point; the analyzers still compare exactly, and the `conformance` analyzer reports what the mapping does not cover.
4. **#11** Pin a real `styles.xml` regression fixture — **blocked** on the licensing question (§8j / `LICENSING.md`).
5. ~~**#21** Replace the substring NL fallback with BM25~~ — **DONE** (`00c3bd1`). Worth being precise about what this did and did not settle: the counters exist to decide **embeddings vs lexical**, and that stays open pending usage data. BM25 replaced *"take the first substring hit"*, which lost to lexical scoring under every hypothesis, so it raises the floor the measurement compares against rather than pre-empting it.
6. **New analyzers**, driven by the gap log rather than guesswork. The obvious blind spots from §10: **DrawingML** (shape geometry, effects, theme style matrices), **formulas** (the biggest MS-OI29500 cluster in SpreadsheetML), and **fields** (`w:fldSimple`, `w:instrText` — TOC, cross-references and page numbers all run through them, and they interact directly with bookmarks).

**Method note.** Both modules shipped in §8m/§8n were **mutation-tested**: the implementation was deliberately broken several distinct ways and the tests re-run, to check they fail for the right reasons rather than agreeing with themselves. This found a real gap in the bookmark tests (no case had two range kinds in one document, the only arrangement where kind-matching is observable). Worth doing for every module here — a green suite on first run is not evidence.

**Wired to the panel** in `c0c4b0c`, which also fixed a latent design bug: `buildComputedEvidence` used `ANALYSIS_TARGETS.find`, so only the **first** matching entry ran. That was fine while each part had exactly one analysis and wrong the moment it did not — a `word/document.xml` carries formatting *and* bookmarks *and* possibly OLE objects, which are independent questions. It now runs every match, unions the sibling requests so a part is fetched once, and merges the results; one analysis throwing no longer suppresses the rest. **Anything added to `ANALYSIS_TARGETS` from here on composes rather than shadows.**

**Still open:**
- **#11** Pin a real `styles.xml` regression fixture — **blocked** on the licensing question in §8j / `LICENSING.md`.
- **#21** Replace the substring NL fallback with BM25. Counters are live (§8h), so **measure before building**.

## 10. Known gaps at time of writing

- ~~SpreadsheetML semantics~~ — **DONE**, see §4.
- ~~PresentationML semantics~~ — **DONE**, see §8g; resolver shipped.
- **DrawingML** — the largest remaining blind spot. Shared across all three formats; ~475 element declarations; **264 variations for Part 1 §21 plus 180 for §20.** Charts (§8k) cover one corner of it; shape geometry, effects and the theme style matrices are untouched.
- **Formulas** — the single biggest MS-OI29500 cluster in SpreadsheetML, never researched.
- **Fields** (`w:fldSimple`, `w:instrText`) — TOC, cross-references and page numbers all run through them, and they interact directly with bookmarks (#26).
