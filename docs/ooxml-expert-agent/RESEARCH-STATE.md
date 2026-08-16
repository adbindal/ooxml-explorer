# OOXML Expert Agent — Research State & Resume Point

**Status:** research complete; build plan published; **Stage 0 shipped, Stage 1a shipped** (see §8b).
**Last updated:** 2026-08-17
**Purpose:** durable record so this work can resume after a session ends. If you are picking this up cold, read this file top to bottom before doing anything else.

Published write-ups (private artifacts):
- Part 1 — The 29-Tag Problem: https://claude.ai/code/artifact/e1b1738f-2969-42c6-9437-190b0a29465f
- Part 2 — What the Schema Can't Tell You: https://claude.ai/code/artifact/d3fb80ba-ea48-411a-b04a-848c52418147
- **Part 3 — Staff Engineer in a Box (the build plan): https://claude.ai/code/artifact/a76cb080-e4fd-4be5-b524-9ca490a94470**

**To resume:** read §8b for what shipped, then §9 for next actions. The SpreadsheetML and PresentationML/DrawingML research was never completed (§10) — those are the gaps to re-fill. The build plan's architecture does not depend on them; only Stage 2's per-format detail does.

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
| **[MS-OI29500]** | **1,895 clause-keyed entries**, each a literal *spec-says / Office-does* pair. Verified: toc.json = 2,152 nodes. Distribution: **Part 1 §18 (Spreadsheet) 534, §17 (Word) 529, §21 (DrawingML components) 264, §20 (DrawingML framework) 180, §22 (shared) 106, Part 4 §19 104, §19 (Presentation) 100** | 1.2 MB DOCX | **Needs counsel** |
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

⚠️ **Not wired up.** `checkPackageIntegrity()` has no caller yet — see task #9.

**Remaining in Stage 1:** MCE preprocessing (task #8). Naive walkers double-count every textbox, because Word writes shapes twice — DrawingML in `mc:Choice`, VML in `mc:Fallback`. Required before the Stage 2 resolvers. The seam is clean: `packageIntegrity` consumes a path→content map, so MCE slots in as a transform over that map.

## 9. Next actions

1. ~~Research (Word, architecture, tooling, storage)~~ — done.
2. ~~Write the build plan~~ — done, Part 3 above.
3. ~~Fix the `geminiService.ts` context bug~~ — done, `3af80df`.
4. ~~Stage 0~~ — done, see §8b. **All work is consolidated on the single branch `feat/schema-derived-rag-corpus`; `main` is untouched.**
5. ~~Stage 1a — package integrity~~ — done, `1a7905f`. See §8c.
6. **Stage 1b — MCE preprocessing** (task #8), then **1c — surface findings in the UI** (task #9). Together these finish Stage 1 and make the *Verified* badge tier real.
7. Decide: which format goes first (recommend sequencing, not parallelising — resolvers are genuinely different per format).
8. Decide: MS-OI29500 licensing — ask or not. Gates Stage 3 only.
9. Decide: audience (self vs onboarding engineers). If mentoring is primary, Stage 5 moves up.

## 10. Known gaps at time of writing

Everything researched so far is **WordprocessingML-heavy**. For the tri-format ambition these remain thin:

- **SpreadsheetML semantics** — the `cellXfs`/`cellStyleXfs`/`xfId` resolution chain, `apply*` flags, number-format application, 1900 vs 1904 date systems, the column-width character formula, shared formulas via `@si`/`@ref`, calc chain authority. Note: **534 MS-OI29500 variations for Part 1 §18 — more than Word's 529.**
- **PresentationML semantics** — placeholder inheritance via `p:ph/@type` and `@idx`, `clrMap`/`clrMapOvr` resolution, `p:txStyles` list-style chain, `a:fmtScheme` style references and their indexing convention, group transform math (`chOff`/`chExt`).
- **DrawingML** — the biggest blind spot. Shared across all three formats; ~475 element declarations; **264 variations for Part 1 §21 plus 180 for §20.**

Agents were dispatched on the first two. If their reports are missing from this file, they did not land.
