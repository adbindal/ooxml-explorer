# 1-Pager: OOXML RAG Pipeline Robustness & Accuracy Plan

## 1. The Problem Statement
Our current evaluation strategy assumes that the manually-created "Golden" dataset (`public/rag-data.json`) is the absolute source of truth. However, recent diff analyses revealed that:
1.  **Golden Data Inaccuracies**: The Golden dataset contains human errors (e.g., `<w:document>` citing section `17.3.1.10` instead of `17.2.3`).
2.  **LLM Accuracy**: The Mastra pipeline actually produced more accurate citations (`17.2.3`) and attributes (`conformance`) than the Golden reference.
3.  **Namespace Inconsistencies**: The LLM output mixed short prefixes (`w`) with full XML URIs.

If the baseline contains errors, we cannot reliably use it for automated calibration. We need a robust, scalable pipeline that ensures the final RAG database is 100% accurate with high confidence.

---

## 2. Proposed Architecture: Hybrid Verification

We will split the RAG data generation into three distinct layers of responsibility:

```
                  ┌──────────────────────────────┐
                  │   Core Ingestion Pipeline    │
                  └──────────────┬───────────────┘
                                 │
         ┌───────────────────────┴───────────────────────┐
         ▼                                               ▼
┌────────────────────────────────┐             ┌────────────────────────────────┐
│   Deterministic Layer (XSD)    │             │    Generative Layer (LLM)      │
├────────────────────────────────┤             ├────────────────────────────────┤
│ • Namespaces                   │             │ • Natural Language Definitions │
│ • Valid Attributes             │             │ • ECMA-376 Section Citations   │
│ • Valid Parent Elements        │             │ • Open XML SDK Class Mappings  │
└────────────────┬───────────────┘             └────────────────┬───────────────┘
                 │                                              │
                 └───────────────────────┬──────────────────────┘
                                         │
                                         ▼
                        ┌────────────────────────────────┐
                        │   Verification Layer (HITL)    │
                        ├────────────────────────────────┤
                        │ • Compare against Golden       │
                        │ • Ingestion-Driven Correction  │
                        └────────────────────────────────┘
```

---

## 3. Step-by-Step Execution Plan

### Phase 1: Namespace Decoupling & Prompt Hardening (Immediate)
*   **Action**: Decouple short prefixes from full XML URIs at the application layer.
*   **Implementation**:
    1.  Add a `NAMESPACE_MAP` lookup table in the codebase.
    2.  Harden the Mastra workflow prompt to strictly return short prefixes (`w`, `r`, `p`).
*   **Outcome**: Resolves all namespace mismatches and ensures consistency.

### Phase 2: Ingestion-Driven Golden Dataset Correction (Immediate)
*   **Action**: Shift from a "match-the-golden-file" mindset to an "ingestion-driven correction" workflow.
*   **Implementation**:
    1.  Run the ingestion pipeline to generate `generated-rag.json`.
    2.  Run `diff_rag.ts` to identify mismatches.
    3.  Perform a quick human-in-the-loop review of the mismatches. 
    4.  If the LLM output is correct (e.g., `17.2.3` for `w:document`), merge/promote the LLM output into `public/rag-data.json`.
*   **Outcome**: The Golden dataset becomes a living, high-quality, verified dataset.

### Phase 3: XSD-Backed Schema Grounding (Medium-Term)
*   **Action**: Remove LLM dependency for deterministic schema fields (attributes, parents, namespaces).
*   **Implementation**:
    1.  Download the official ECMA-376 XML Schema Definition (`.xsd`) files.
    2.  Create a script that parses these `.xsd` files to extract the exact schema structure.
    3.  Integrate this script as a step in the Mastra pipeline to automatically populate the `attributes`, `parents`, and `namespace` fields.
*   **Outcome**: 100% accuracy for structural metadata. The LLM is only responsible for writing the human-readable definition and finding the citation.

### Phase 4: Cross-Model Consensus Verification (Long-Term)
*   **Action**: Automate citation verification.
*   **Implementation**:
    1.  Add a step in the Mastra workflow that queries two independent LLM providers (e.g., Gemini and Claude).
    2.  Compare their generated citations. If they agree but differ from the Golden dataset, flag the tag for developer review.
*   **Outcome**: Automated detection of potential errors in both the LLM outputs and the Golden dataset.
