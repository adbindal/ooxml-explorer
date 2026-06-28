# OOXML RAG Pipeline Robustness & Accuracy Plan

## 1. The Problem Statement
Our RAG pipeline must ensure that the generated RAG reference database (`public/rag-data.json`) is 100% accurate and conforms strictly to the ECMA-376 specification. 

However, running LLM generation sequentially on dozens or hundreds of tags is extremely slow, and relying purely on LLM memory leads to hallucinations in citations, attributes, and parent elements. 

To solve this, we are transitionining the pipeline from a simple sequential script to a **parallelized, agentic workflow** using Mastra.

---

## 2. Proposed Architecture: Collaborative Agentic Loop

We decompose the pipeline into specialized **Mastra Agents** and **deterministic code steps** composed using Mastra's native workflow engine.

```
                  ┌──────────────────────────────────────────────┐
                  │                 [Tag Input]                  │
                  └──────────────────────┬───────────────────────┘
                                         │
                                         ▼
                  ┌──────────────────────────────────────────────┐
                  │          1. informationRetrieverAgent        │
                  │       (Fetches spec text via Web Search)     │
                  └──────────────────────┬───────────────────────┘
                                         │
                                         ▼
                  ┌──────────────────────────────────────────────┐
                  │            2. schemaGeneratorAgent           │
                  │       (Synthesizes reference schema)         │
                  └──────────────────────┬───────────────────────┘
                                         │
                                         ▼
 ┌───────────────────────────────────────┴───────────────────────────────────────┐
 │                                                                               │
 │                ┌──────────────────────────────────────────────┐               │
 │                │             3. xsdGroundingStep              │               │
 │                │        (Extracts attributes & parents        │               │
 │                │          deterministically from XSD)         │               │
 │                └──────────────────────┬───────────────────────┘               │
 │                                       │ (Validation Report)                   │
 │                                       ▼                                       │
 │                ┌──────────────────────────────────────────────┐               │
 │                │              4. schemaAuditorAgent           │               │
 │                │       (LLM-as-a-Judge - Reviews Schema)      │               │
 │                └──────────────────────┬───────────────────────┘               │
 │                                       │                                       │
 │          ┌────────────────────────────┼────────────────────────────┐          │
 │          │ (ACCEPT_GENERATED)         │ (REJECT_GENERATED)         │ (SUSPEND)│
 │          ▼                            ▼                            ▼          │
 │  ┌───────────────┐            ┌───────────────┐            ┌───────────────┐  │
 │  │ Auto-Approve  │            │  Retry < 3?   ├────────────┼───────────────┼──┘
 │  │ & Skip Retry  │            └───┬───────┬───┘            │    Pause      │
 │  └───────┬───────┘                │ (Yes) │ (No)           │   Workflow    │
 │          │                        ▼       ▼                └───────┬───────┘
 │          │                  [Self-Heal]  [Log Defect to            │
 │          │                               CALIBRATION_FEEDBACK.md]  │
 │          │                                │                        │
 └──────────┼────────────────────────────────┴────────────────────────┼──────────┘
            │                                                         │
            └────────────────────────────┬────────────────────────────┘
                                         │
                                         ▼
                  ┌──────────────────────────────────────────────┐
                  │           5. commitApprovedDataStep          │
                  │       (Writes Golden & Compiles High-    │
                  │            Priority Static KB)               │
                  └──────────────────────────────────────────────┘
```

### The 4 Core Steps:
1.  **`informationRetrieverAgent`**:
    *   **Type**: Mastra Agent + Web Search & Scraper tools.
    *   **Task**: Searches for and retrieves the exact online spec text for the target tag, providing "open-book" context to the generator.
2.  **`schemaGeneratorAgent`**:
    *   **Type**: Mastra Agent.
    *   **Task**: Takes the retrieved spec text and any human `reviewerNote` (from the golden doc), and synthesizes a clean RAG reference schema.
3.  **`xsdGroundingStep`**:
    *   **Type**: Deterministic Code Step.
    *   **Task**: Parses the official ECMA-376 `.xsd` schema files to extract the 100% correct list of valid attributes and parent elements for the tag. It generates a `validationReport` indicating any structural mismatches.
4.  **`schemaAuditorAgent`** (The Judge):
    *   **Type**: Mastra Agent.
    *   **Task**: Compares the generated schema against the existing Golden doc, the validation report, and the retrieved spec text. It makes one of three decisions:
        *   `ACCEPT_GENERATED`: The schema is correct, acceptable, or an upgrade. Phrasing variations are accepted.
        *   `REJECT_GENERATED`: The schema has actual errors/hallucinations. This triggers the **Autonomous Self-Correction Loop** (up to 3 retries, feeding the Judge's reasoning back to the Generator).
        *   `SUSPEND_FOR_REVIEW`: A major conflict requiring human intervention.

---

## 3. Workflow Parallelization & Concurrency

To solve the execution speed issue, we use Mastra's native **`.foreach()`** operator in the workflow definition. 

Instead of sequential loops inside a single step, the workflow splits the array of tags and runs them in parallel:

```typescript
ingestionWorkflow
  // 1. Retrieve information for all tags in parallel (concurrency: 5)
  .foreach(retrieveInfoStep, { concurrency: 5 })
  
  // 2. Generate schemas in parallel (concurrency: 5)
  .foreach(generateSchemaStep, { concurrency: 5 })
  
  // 3. Evaluate, validate, and self-correct in parallel (concurrency: 5)
  .foreach(evaluateAndAuditStep, { concurrency: 5 })
  
  // 4. Commit all results in a single final step
  .then(commitApprovedDataStep)
  .commit();
```

This guarantees a **5x speedup** while keeping the execution safe from API rate limits.

---

## 4. Scalable Static KB Compilation (Bundle Optimization)

Compiling thousands of OOXML tags into a static frontend TypeScript file would bloat the bundle size. 

We split the storage strategy:
1.  **Full Database (`public/rag-data.json`)**: Contains every single ingested tag. The frontend can lazy-load this file on demand via `fetch` when a user inspects a less common tag.
2.  **Static KB (`services/staticKnowledgeBase.ts`)**: Contains only **high-priority** tags (e.g., the 33 core calibration tags, or tags explicitly marked as `priority: "high"` by a reviewer). Only this small, optimized subset is bundled into the frontend.

---

## 5. Phase-by-Phase Implementation Plan

### Phase 1: Clean Calibration & 3-Decision Model (Completed)
*   Implemented the robust `extractJson` parser to prevent LLM formatting crashes.
*   Simplified the Judge decisions to `ACCEPT_GENERATED`, `REJECT_GENERATED`, and `SUSPEND_FOR_REVIEW` to eliminate false-positive defect logs.
*   Implemented the human `reviewerNote` injection.
*   Implemented the priority-based `staticKnowledgeBase` compiler.

### Phase 2: Mastra Agent Refactoring (Immediate)
*   Refactor the workflow to define first-class `Agent` instances for the **Generator** and **Auditor**.
*   Remove the `execSync(jetski)` shell wrapper, running the agents in-process.
*   Validate the new agent execution in Mastra Studio.

### Phase 3: Parallelization (`.foreach()`) (Immediate)
*   Refactor the workflow steps to focus on a single tag.
*   Hook them up using Mastra's `.foreach(..., { concurrency: 5 })` to enable parallel execution.

### Phase 4: Deterministic XSD Grounding & Retriever Agent (Medium-Term)
*   Write the XSD parser helper to extract attributes and parents.
*   Equip the Retriever Agent with web search tools to fetch the spec text before generation.
