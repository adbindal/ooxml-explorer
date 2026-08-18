# [MS-OI29500] — licensing findings

**Not legal advice.** This is research to brief a lawyer efficiently. Every quote was fetched and extracted directly; the landing-page HTML and the DOCX were verified **word-for-word identical**.

Sources: [landing page](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/1fd4a662-8623-49c0-82f0-18fa91b413b8) · `[MS-OI29500]-260519.docx` (936 pages, footer `v20260519`, © 2026 Microsoft).

---

## 1. The notice, verbatim

> **Copyrights.** This documentation is covered by Microsoft copyrights. Regardless of any other terms that are contained in the terms of use for the Microsoft website that hosts this documentation, you can make copies of it in order to develop implementations of the technologies that are described in this documentation and can distribute portions of it in your implementations that use these technologies or in your documentation as necessary to properly document the implementation. You can also distribute in your implementation, with or without modification, any schemas, IDLs, or code samples that are included in the documentation. This permission also applies to any documents that are referenced in the Open Specifications documentation.
>
> **No Trade Secrets.** Microsoft does not claim any trade secret rights in this documentation.
>
> **Patents.** Microsoft has patents that might cover your implementations of the technologies described in the Open Specifications documentation. Neither this notice nor Microsoft's delivery of this documentation grants any licenses under those patents or any other Microsoft patents. However, a given Open Specifications document might be covered by the Microsoft Open Specifications Promise or the Microsoft Community Promise. If you would prefer a written license, or if the technologies described in this documentation are not covered by the Open Specifications Promise or Community Promise, as applicable, patent licenses are available by contacting iplg@microsoft.com.
>
> **Reservation of Rights.** All other rights are reserved, and this notice does not grant any rights other than as specifically described above, whether by implication, estoppel, or otherwise.
>
> **Support.** For questions and support, please contact dochelp@microsoft.com.

**"Preliminary Documentation" is NOT present** — verified by grepping the full 2.2M-character extract. Zero hits.

### The term the notice overrides
The Copyrights paragraph opens *"Regardless of any other terms… in the terms of use for the Microsoft website that hosts this documentation"*. It is pointing at [Microsoft Learn TOU](https://learn.microsoft.com/en-us/legal/termsofuse):

> **Personal and Non-Commercial Use Limitation.** …You may not modify, copy, distribute, transmit, publicly display, perform, reproduce, publish, license, create derivative works from, transfer or sell any information… obtained from the Services (except for your own, personal, non-commercial use) without prior written consent from Microsoft.

**So the general TOU flatly prohibits this, and the IPR notice is the only carve-out — exactly as wide as its own words.** That makes the scope question load-bearing rather than academic.

### What one entry looks like
> **Part 1 Section 17.3.1.13, jc (Paragraph Alignment)**
> a. The standard states that no alignment is applied to the paragraph when the paragraph alignment setting is never specified in the style hierarchy.
> Word applies a left alignment to the paragraph under this circumstance.

The whole corpus is that formula. **1,894 entries** under §2.1, plus 19 under §3.

---

## 2. Clearly permitted

- **Copying to develop an implementation** — unambiguous on the face of the text.
- **Distributing schemas, IDLs and code samples**, with or without modification. This is a *separate, broader* sentence with no "as necessary" qualifier. docx4j relies on exactly this.
- **No trade-secret claim** blocks anything.
- ✅ **Patents are settled.** `[MS-OI29500]` is expressly covered by the **Open Specification Promise**, confirmed two independent ways:
  1. The OSP page's "Implementer Notes for Covered Specifications" lists it by name.
  2. The machine-readable [Patent Map](https://officeprotocoldoc.z19.web.core.windows.net/files/public-patents-export/PatentMapList.xml) has exactly one row: `ows_PatentNumbers="None" ows_PatentApplications="None" ows_Programs=";#OSP;#"`.
  - It does **not** appear in the Community Promise page (grepped, zero hits).

🔴 **Patents ≠ copyright, and Microsoft says so.** OSP FAQ: *"Copyrights in the Covered Specifications are not provided through the OSP."* The OSP settles patents completely and copyright **not at all**.

---

## 3. Genuinely ambiguous — flagged, not resolved

1. **Does "in order to develop implementations" reach a tool that *teaches* the format?** Our users are implementers; the tool itself is not an OOXML implementation. Whether the purpose clause attaches to the copier or the copier's users is not addressed anywhere. **No Microsoft clarification or FAQ on this phrase was found.**
2. **Is a RAG knowledge base "your implementation" or "your documentation"?** The distribution right is narrower than the copying right, and "as necessary to properly document the implementation" is doing real work with no published gloss.
3. **Bulk extraction / derived database.** The notice is *silent*, and Reservation of Rights means silence is no. Note the asymmetry the drafters chose: schemas/IDLs/code samples get "with or without modification, **any**"; the prose gets only "portions… as necessary". A structured transformation of the prose falls in that gap.
4. **Copies vs distribution** are different grants in one sentence. A locally-built KB may sit in the first; shipping it lands in the second.
5. **Thin expression.** "The standard states X. Word does Y." is close to bare fact in formulaic phrasing. Lawyer question.

---

## 4. Precedent, verified in source

| Project | What it actually does |
|---|---|
| **Samba** | 🔴 **Goes further than we plan.** `libcli/util/*_err_table.txt` — **52,098 lines** of bulk-scraped Open Specifications tables in a public GPL repo, used as build inputs. Reproduces **the entire IPR notice verbatim** plus source URLs. ⚠️ **But Samba holds a separately negotiated agreement via the PFIF**, so it may stand on ground we do not have. Show the lawyer the files *and* that fact together. |
| **docx4j** | Relies explicitly on the **schema sentence** in its `NOTICE`. For MS-OI29500 itself: **citation only** (`[MS-OI29500] 2.1.87`), no reproduction. |
| **Apache POI** | Pointer citations only, **no NOTICE entry** (grepped, zero hits). ⚠️ Its guidelines cite the **OSP as authority for a copyright question** — a widely-copied conflation the OSP FAQ contradicts. Not a safe basis. |
| **python-docx** | **Zero** OI29500 references. Documents Word behaviour in its own prose from observation — the independent-derivation model. |
| **SheetJS** | One structural citation. Pointer only. |
| **ASF policy** | [resolved.html](https://www.apache.org/legal/resolved.html): must not include unmodifiable-licence material in version control or source releases; build-time download is acceptable. Fielding on LEGAL-120: *"that would not impact your implementation unless you wanted to include the spec itself."* **ASF's consistent shape: implement freely, don't ship the document.** |

---

## 5. Route to certainty

| Purpose | Contact |
|---|---|
| Documentation questions incl. scope | **dochelp@microsoft.com** (the notice's own Support line; Dev Center labels it "Interop Documentation Support Team") |
| Patent licences | **iplg@microsoft.com** |
| IP Licensing Team | **protocol@microsoft.com** |
| General copyright permission | Microsoft Corporation, One Microsoft Way, Redmond, WA 98052-6399 — ⚠️ `[SUMM]` only, page 403'd; **verify before relying** |

**Public forum:** [M365 Open Specifications Q&A](https://learn.microsoft.com/en-us/answers/tags/328/m365-office-open-specs) — 160 questions, active through July 2026, staffed by named Microsoft moderators. Observed: a 2024 question answered substantively in **3 days**, with Microsoft committing to file a bug against MS-OI29500 itself.

⚠️ **Every thread read was technical. No instance found of a licensing or scope question being answered there.** Do not assume the channel handles IP questions.

**Closed door:** the `MicrosoftDocs/open_specs_office` repo referenced in page metadata is **private** (404). No CC-licensed route in.

**Fallback material:** the [TOC JSON](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/toc.json) gives 2,149 stable per-clause URLs. Titles carry the clause reference; hrefs are opaque GUIDs. A link-only fallback needs a GUID↔clause map — **a dataset of URLs and headings, not of Microsoft's prose.** Materially different asset.

---

## 6. Four questions to send a lawyer as-is

1. *Microsoft's Open Specifications notice says I "can make copies of it in order to develop implementations of the technologies that are described in this documentation." I'm building a developer tool that explains OOXML markup to engineers who write OOXML readers and writers. My tool is not itself an OOXML implementation, but its users are building them. Does that purpose clause cover me, or does it only cover someone writing the implementation themselves?*

2. *The same notice lets me "distribute portions of it in your implementations that use these technologies or in your documentation as necessary to properly document the implementation." I want to load about 1,900 of its entries into a searchable knowledge base that an AI assistant answers from — typically a short paraphrase plus a clause number, occasionally a sentence quoted directly. Is that "distributing portions," is it building a derived database that the notice doesn't cover, or is it something else? Does it change the answer if the knowledge base ships with the product versus being built locally by each user?*

3. *Each entry is two sentences in a fixed formula: "The standard states [X]. Word does [Y]." Given how thin and formulaic that expression is, how much of it is protected at all — and does paraphrasing it meaningfully reduce my exposure, or does copying 1,900 of them create a compilation problem regardless of how I word each one?*

4. *Microsoft's general Learn terms of use prohibit reproducing or creating derivative works from the site for anything but personal, non-commercial use, but the Open Specifications notice starts "Regardless of any other terms… in the terms of use for the Microsoft website that hosts this documentation." If my use falls outside what the Open Specifications notice permits, does the restrictive general terms-of-use clause snap back into effect and govern instead? And is it worth writing to dochelp@microsoft.com for a scope answer before we build, or does asking create problems of its own?*

---

## 7. Not verified

- **Microsoft guidance on the "develop implementations" phrase** — searched, none found. Cannot prove none exists.
- **The copyright-permissions postal address** — summarizer only; page returned 403.
- **LibreOffice** — grep returned no match but verification did not complete. **Do not cite either way.**
- **Whether the OSP "Implementer Notes" table pairs [MS-OI29500] with specific 29500 versions** — the rendered HTML is single-column and the pairing is not unambiguously recoverable.
- **Any instance of Microsoft answering a licensing question, or objecting to bulk extraction.** Both negatives, unestablished.
- **The comparative value assessment of fallback options** — raw material gathered, analysis not written. Main gap.
