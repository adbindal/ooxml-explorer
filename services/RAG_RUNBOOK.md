# OOXML RAG Ingestion & Model Hardening Runbook

This runbook outlines the process for calibrating, hardening, and verifying the OOXML RAG schema database. It serves as a guide for the team to understand how we improve model output, handle hallucinations, and maintain a high-precision knowledge base.

---

## 1. RAG Calibration Architecture

The RAG database is populated using a two-phase loop: **Ingestion** (automated LLM queries) and **Calibration** (verification against a golden reference dataset).

```
 [Ingest Script] ──> [Mastra Workflow] ──> [LLM (Gemini)] ──> [generated-rag.json]
                                                                     │
                                                           (diff_rag.ts Calibration)
                                                                     ▼
 [storageService] <── [IndexedDB] <── [rag-data.json] <─── [Human Review & Promote]
```

*   **Golden Dataset (`public/rag-data.json`)**: The verified source of truth containing exact ECMA-376 citations, correct namespaces, and curated attributes.
*   **Generated Output (`public/generated-rag.json`)**: Raw outputs from the Mastra workflow pipeline.
*   **Static Fallback (`services/staticKnowledgeBase.ts`)**: Auto-compiled subset of the golden dataset for offline/fast-path execution.

---

## 2. Ingestion Calibration & Diff Analysis

When running the diff tool (`npx tsx scripts/diff_rag.ts`), you may see mismatches. The table below outlines how to evaluate these mismatches:

| Field | Golden Reference (Correct) | LLM Generated (Needs Hardening) | Mitigation / Why Golden is Correct |
| :--- | :--- | :--- | :--- |
| **Namespace** | Short prefix (e.g., `w`, `r`, `p`) | Full XML URI (e.g., `http://schemas.openxmlformats.org/...`) | **Enforced in Prompt**: The app's UI and tree-query logic query elements by their short prefix (e.g., `w:document`). Storing the full URI breaks the query matching. |
| **Citation** | Exact element section heading (e.g., `17.3.1.10` for `document`) | High-level section or slightly off (e.g., `17.2.3` for `document`) | **Enforced in Prompt**: The LLM tends to generalize (picking the general parent section of the spec) or slightly hallucinate the section numbers. The golden dataset is verified against the actual ECMA-376 PDF specification. |
| **SDK Class** | Matches C# Open XML SDK (e.g., `Document`) | Sometimes generic or `N/A` | **Enforced in Prompt**: The C# SDK class name is required for code generation/SDK mapping in the editor. |
| **Attributes & Parents** | Highly curated, minimal subset | Highly comprehensive (includes all schema attributes) | **Curated in Golden**: The Golden dataset is intentionally kept minimal to avoid token bloat during RAG prompt injection. The Generated dataset is technically more complete. |

---

## 3. Prompt Hardening Guidelines

To prevent the LLM from hallucinating citations or returning incorrect namespaces, the system prompt in [src/mastra/index.ts](file:///Users/adbindal/code/exp/ooxml-explorer/src/mastra/index.ts) is hardened with the following strict rules:

### A. Namespace Prefix Hardening
Instruct the model to avoid full XML namespace URIs and only return the shorthand prefix:
```
- For 'namespace', you MUST return the short prefix (e.g., 'w' for WordprocessingML, 'r' for SpreadsheetML/Relationships, 'p' for PresentationML) instead of the full XML namespace URI.
```

### B. Citation Precision Hardening
Prevent the model from generalizing or guessing section numbers. Force it to target the exact section heading for the element:
```
- For 'citation', you MUST reference the most specific section number in the ECMA-376 specification that defines this specific element. 
- Do NOT use high-level parent section numbers. For example, the element `document` is defined in `ECMA-376 Part 1, Section 17.3.1.10`, not the general `Section 17.2.3`.
- In case of multiple citations, pick the one that is the section heading and talks about that particular node only.
```

### C. Attribute Curation
If you want the model to generate a curated subset of attributes rather than the entire schema, add constraints:
```
- For 'attributes', only include the most common or structural attributes that directly configure the element's behavior (e.g., identifiers, values, dimensions). Avoid listing internal revision tracking attributes (like rsid*) unless they are critical.
```

---

## 4. Operational Workflow

1.  **Run Ingestion**:
    ```bash
    npx tsx scripts/ingest_rag.ts
    ```
2.  **Run Calibration Diff**:
    ```bash
    npx tsx scripts/diff_rag.ts
    ```
3.  **Review & Promote**:
    Review the diff. If the generated output is correct and contains new tags, merge them into `public/rag-data.json` and commit.
