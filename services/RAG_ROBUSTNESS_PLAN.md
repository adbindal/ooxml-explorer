# OOXML RAG Pipeline Robustness & Accuracy Plan

## 1. The Problem Statement
Our current evaluation strategy assumes that the manually-created "Golden" dataset (`public/rag-data.json`) is the absolute source of truth. However, recent diff analyses revealed that:
1.  **Golden Data Inaccuracies**: The Golden dataset contains human errors (e.g., `<w:document>` citing section `17.3.1.10` instead of `17.2.3`).
2.  **LLM Accuracy**: The Mastra pipeline actually produced more accurate citations (`17.2.3`) and attributes (`conformance`) than the Golden reference.
3.  **Namespace Inconsistencies**: The LLM output mixed short prefixes (`w`) with full XML URIs.

If the baseline contains errors, we cannot reliably use it for automated calibration. We need a robust, scalable pipeline that ensures the final RAG database is 100% accurate with high confidence.

---

## 2. Proposed Architecture: Self-Healing Feedback Loop

To solve this, we introduce an **LLM-as-a-Judge** layer coupled with an **Autonomous Self-Correction Loop** and a **Systemic Feedback Backlog**. 

This architecture handles differences at two distinct timescales:
1.  **Short-Term (Autonomous)**: If the generator produces an output that the Judge deems worse than the Golden reference (`KEEP_GOLDEN`), the pipeline automatically retries the generator (up to 3 times) by feeding it the Judge's exact error analysis.
2.  **Long-Term (Systemic)**: If the generator fails to correct itself after 3 retries, the pipeline logs the failure and the Judge's reasoning to a local markdown file (`services/CALIBRATION_FEEDBACK.md`), creating a structured backlog for developer prompt engineering.

```
                      ┌──────────────────────────────┐
                      │   1. generateSchemaStep      │ ◄──────────────────────────┐
                      │  (Generator LLM / Consensus) │                            │
                      └──────────────┬───────────────┘                            │
                                     │                                            │ (Retry with
                                     ▼                                            │  Feedback)
                      ┌──────────────────────────────┐                            │
                      │   2. xsdValidationStep       │                            │
                      │ (Deterministic Schema Check) │                            │
                      └──────────────┬───────────────┘                            │
                                     │ (Passes Validation Report)                 │
                                     ▼                                            │
                      ┌──────────────────────────────┐                            │
                      │   3. llmJudgeStep            │                            │
                      │ (LLM-as-a-Judge - Pro Model) │                            │
                      └──────────────┬───────────────┘                            │
                                     │                                            │
         ┌───────────────────────────┼───────────────────────────┐                │
         │ (UPGRADE_GOLDEN)          │ (KEEP_GOLDEN)             │ (SUSPEND)      │
         ▼                           ▼                           ▼                │
┌──────────────────┐        ┌──────────────────┐        ┌──────────────────┐      │
│ Auto-Approve &   │        │   Retry < 3?     ├────────┼──────────────────┼──────┘
│ Update Golden    │        └──┬─────────────┬─┘        │ Pauses Workflow  │
└────────┬─────────┘           │ (Yes)       │ (No)     │ Awaits Human     │
         │                     │             │          └────────┬─────────┘
         │                     ▼             ▼                   │
         │              [Auto-Correct]  [Log Defect to           │
         │                              CALIBRATION_FEEDBACK.md] │
         │                                   │                   │
         └───────────────────────────┬───────┴───────────────────┘
                                     │
                                     ▼
                      ┌──────────────────────────────┐
                      │   4. commitApprovedDataStep  │
                      │   (Write & Compile KB)       │
                      └──────────────────────────────┘
```

---

## 3. Step-by-Step Execution Plan

### Phase 1: Namespace Decoupling & Prompt Hardening (Completed)
*   **Action**: Decouple short prefixes from full XML URIs at the application layer.
*   **Implementation**:
    1.  Added a `NAMESPACE_MAP` lookup table and helper in the codebase.
    2.  Hardened the Mastra workflow prompt to strictly return short prefixes (`w`, `r`, `p`).
*   **Outcome**: Resolves all namespace mismatches and ensures consistency.

### Phase 2: LLM-as-a-Judge & Autonomous Self-Healing (Immediate)
*   **Action**: Implement the `llmJudgeStep` and the self-correction retry loop.
*   **Implementation**:
    1.  **Define the Judge Prompt**: Instruct a high-capacity model (Gemini Pro) to act as a pedantic OOXML specification auditor.
    2.  **Judge Input**: Receives the `generated` schema, the `golden` schema, and a `validationReport`.
    3.  **Judge Output**: Returns a JSON object with:
        *   `decision`: `'UPGRADE_GOLDEN' | 'KEEP_GOLDEN' | 'SUSPEND_FOR_REVIEW'`
        *   `reasoning`: The exact reason for the decision (e.g., why the generated version was worse or what was incorrect).
        *   `resolvedDoc`: The final corrected reference document.
    4.  **Autonomous Retry Loop**: If the Judge returns `KEEP_GOLDEN`, re-run `generateSchemaStep` for that tag up to 3 times, appending the Judge's `reasoning` as feedback in the prompt.
    5.  **Defect Logging**: If retries are exhausted and the output is still rejected, append the failure details and Judge's reasoning to `services/CALIBRATION_FEEDBACK.md`.
    6.  **Suspension**: If the Judge returns `SUSPEND_FOR_REVIEW`, the workflow suspends, displaying the Judge's reasoning and the conflict in Mastra Studio.
*   **Outcome**: Simple errors are resolved autonomously; systemic prompt/model gaps are logged as actionable developer tasks.

### Phase 3: XSD-Backed Schema Grounding (Medium-Term)
*   **Action**: Remove LLM dependency for deterministic schema fields (attributes, parents, namespaces).
*   **Implementation**:
    1.  Download the official ECMA-376 XML Schema Definition (`.xsd`) files.
    2.  Write a script to parse these files and generate a local schema catalog.
    3.  Integrate this catalog into the `xsdValidationStep` to automatically populate and validate the `attributes`, `parents`, and `namespace` fields before the Judge step.
*   **Outcome**: 100% accuracy for structural metadata. The LLM is only responsible for writing the human-readable definition and finding the citation.

### Phase 4: Multi-Model Consensus (Long-Term)
*   **Action**: Automate citation verification.
*   **Implementation**:
    1.  Modify `generateSchemaStep` to query two independent LLM providers (e.g., Gemini and Claude).
    2.  Pass both outputs to the `llmJudgeStep`.
    3.  The Judge compares both outputs against the Golden dataset, resolving minor differences and flagging major ones.
*   **Outcome**: Extremely high confidence in natural-language definitions and citations.
