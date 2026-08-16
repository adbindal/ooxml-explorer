# OOXML Explorer Services & RAG Architecture

This directory contains the core business logic, database, and AI service engines for the OOXML Explorer application.

---

## RAG Data Flow (Runtime)

To provide accurate explanations of XML elements with official ECMA-376 citations and Microsoft Open XML SDK mappings, the application uses a retrieval-augmented generation (RAG) architecture grounded in a pre-generated schema reference dataset.

```
   public/rag-data.json (Golden Reference Dataset)
               |
               v
   services/storageService.ts
   (loads into IndexedDB on app start)
               |
               v
   services/ragRouter.ts
   (looks up a tag; falls back to keyword
    search, then to staticKnowledgeBase.ts)
               |
               v
   services/aiService.ts / services/geminiService.ts
   (injects the retrieved schema as grounding
    context into the AI prompt)
```

### The Datasets
*   `public/rag-data.json`: The **Golden Reference Dataset**. Contains verified, structured OOXML schema information (definitions, attributes, parents, official ECMA-376 citations, and Open XML SDK class mappings) for the full set of covered tags. This is the source of truth for the RAG engine, generated and calibrated offline.
*   `services/staticKnowledgeBase.ts`: The **Static Offline Fallback Knowledge Base**. An auto-generated TypeScript file compiled from the golden dataset, containing only the most common/critical tags. Bundled directly with the application to provide immediate offline access before IndexedDB is initialized.

### Runtime Execution
*   **[services/storageService.ts](./storageService.ts)**: On application start, initializes IndexedDB (`ooxml_explorer_db`) and populates the `rag_schemas` store by fetching `/rag-data.json`.
*   **[services/ragRouter.ts](./ragRouter.ts)**: Intercepts queries for XML elements, retrieves their schema/citations from IndexedDB (falling back to `staticKnowledgeBase.ts` if not found), and injects it as grounding context into the AI request. Also supports runtime "self-healing" overrides recorded via `logRagFeedback`.

### Regenerating the dataset
`public/rag-data.json` and `services/staticKnowledgeBase.ts` are produced by an offline ingestion pipeline that is not part of this application repository. Treat both files as vendored data: update them by re-running that external pipeline, not by hand-editing.

---

## AI Provider Routing

*   **[services/aiProvider.ts](./aiProvider.ts)**: Single source of truth for choosing between Chrome's built-in local model (`chrome-local`) and the Gemini Cloud API (`gemini-cloud`), and for enforcing DLP Mode (see below). Both `aiService.ts` and `geminiService.ts` call into this instead of each maintaining their own provider-detection logic.
*   **DLP Mode** (`ui.dlpMode`, default on): when enabled, no request may reach the cloud API — if local AI isn't available, the call fails with a `DLP_BLOCK` error rather than silently falling back to the cloud. This applies uniformly to every AI action in the app (whole-file explain/technical analysis, diff explain, and the selected-tag explainer).
*   **[utils/guardrails.ts](../utils/guardrails.ts)**: Lightweight prompt-injection guardrails applied to AI input and output regardless of provider.
