import { Mastra } from '@mastra/core';
import { Workflow, createStep } from '@mastra/core/workflows';
import { LibSQLStore } from '@mastra/libsql';
import { execSync } from 'child_process';
import { z } from 'zod';
import * as fs from 'fs';
import * as path from 'path';
import { ReferenceDoc } from '../../services/staticKnowledgeBase';

// 1. Generate Schema Step (Ingestion)
const generateSchemaStep = createStep({
  id: 'generateSchema',
  inputSchema: z.object({
    tags: z.array(z.object({
      tag: z.string(),
      namespace: z.string(),
      domain: z.string()
    }))
  }),
  outputSchema: z.object({
    results: z.array(z.object({
      tag: z.string(),
      namespace: z.string(),
      domain: z.string(),
      definition: z.string(),
      attributes: z.array(z.string()),
      parents: z.array(z.string()),
      citation: z.string(),
      sdkClass: z.string()
    }))
  }),
  execute: async ({ inputData }) => {
    const { tags } = inputData;
    const results = [];

    for (const item of tags) {
      const { tag, namespace, domain } = item;
      console.log(`[Workflow] Processing <${namespace}:${tag}>...`);
      
      const prompt = `
You are an expert in the ECMA-376 Office Open XML (OOXML) specification.
Generate a structured RAG reference document for the following XML tag:
- Tag Name: "${tag}"
- Namespace Prefix: "${namespace}"
- Domain: "${domain}"

You must return a JSON object conforming exactly to this JSON schema:
{
  "type": "object",
  "properties": {
    "tag": { "type": "string" },
    "namespace": { "type": "string" },
    "domain": { "type": "string" },
    "definition": { "type": "string" },
    "attributes": { "type": "array", "items": { "type": "string" } },
    "parents": { "type": "array", "items": { "type": "string" } },
    "citation": { "type": "string" },
    "sdkClass": { "type": "string" }
  },
  "required": ["tag", "namespace", "domain", "definition", "attributes", "parents", "citation", "sdkClass"]
}

Guidelines:
1. "namespace": You MUST return the short prefix (e.g., "w" for WordprocessingML, "r" for SpreadsheetML/Relationships, "p" for PresentationML) instead of the full XML namespace URI.
2. "definition": Provide a clear, precise explanation of what this element configures and its role. Keep it descriptive but concise (2-3 sentences).
3. "citation": You MUST reference the most specific section number in the ECMA-376 specification that defines this specific element. Do NOT use high-level parent section numbers. For example, the element 'document' is defined in 'ECMA-376 Part 1, Section 17.2.3', not the general 'Section 17.2' or the incorrect 'Section 17.3.1.10'. In case of multiple citations, pick the one that is the section heading and talks about that particular element only. Must match the format: "ECMA-376 Part X, Section Y.Z" (e.g. "ECMA-376 Part 1, Section 17.3.1.22"). Verify the exact section number.
4. "sdkClass": Provide the corresponding Microsoft Open XML SDK class name (e.g. "Paragraph" or "TableCell").

Return ONLY the raw JSON block. No markdown wrapper, no explanations.
`;

      // Escape prompt for CLI execution
      const escapedPrompt = prompt.replace(/"/g, '\\"').replace(/`/g, '\\`');
      const command = `jetski --print "${escapedPrompt}"`;
      
      try {
        const stdout = execSync(command, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] });
        
        // Extract JSON block from output
        const jsonMatch = stdout.match(/\{[\s\S]*\}/);
        if (!jsonMatch) {
          throw new Error(`Could not find JSON in jetski output: ${stdout}`);
        }
        
        const doc = JSON.parse(jsonMatch[0].trim());
        results.push(doc);
        console.log(`[Workflow] Successfully processed <${namespace}:${tag}>`);
      } catch (error) {
        console.error(`[Workflow] Failed to process <${namespace}:${tag}>:`, error);
      }

      // Avoid rate limits
      await new Promise(resolve => setTimeout(resolve, 2000));
    }

    return { results };
  }
});

// 2. Evaluate and Diff Step (LLM-as-a-Judge & HITL Suspension)
const evaluateAndDiffStep = createStep({
  id: 'evaluateAndDiff',
  inputSchema: z.object({
    results: z.array(z.object({
      tag: z.string(),
      namespace: z.string(),
      domain: z.string(),
      definition: z.string(),
      attributes: z.array(z.string()),
      parents: z.array(z.string()),
      citation: z.string(),
      sdkClass: z.string()
    }))
  }),
  outputSchema: z.object({
    approved: z.array(z.object({
      tag: z.string(),
      namespace: z.string(),
      domain: z.string(),
      definition: z.string(),
      attributes: z.array(z.string()),
      parents: z.array(z.string()),
      citation: z.string(),
      sdkClass: z.string()
    }))
  }),
  execute: async ({ inputData, suspend, resumeData }) => {
    const { results } = inputData;
    
    // Load golden dataset
    const goldenPath = path.join(process.cwd(), 'public/rag-data.json');
    let golden: ReferenceDoc[] = [];
    if (fs.existsSync(goldenPath)) {
      golden = JSON.parse(fs.readFileSync(goldenPath, 'utf8'));
    }

    const approved: ReferenceDoc[] = [];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const conflicts: any[] = [];

    for (const genDoc of results) {
      const goldDoc = golden.find(g => g.tag === genDoc.tag && g.domain === genDoc.domain);

      if (!goldDoc) {
        // New tag - auto approve
        approved.push(genDoc);
        continue;
      }

      // Basic deterministic validation
      const validationReport = {
        namespaceIsValid: ['w', 'r', 'p', 'rel', 'contentTypes'].includes(genDoc.namespace),
        citationIsValidFormat: /^ECMA-376 Part \d+, Section \d+(\.\d+)*$/.test(genDoc.citation),
        attributesAreArrays: Array.isArray(genDoc.attributes),
        parentsAreArrays: Array.isArray(genDoc.parents)
      };

      // Check if there are any differences at all
      const hasDifferences = 
        goldDoc.citation !== genDoc.citation ||
        goldDoc.namespace !== genDoc.namespace ||
        goldDoc.definition !== genDoc.definition ||
        JSON.stringify(goldDoc.attributes.sort()) !== JSON.stringify(genDoc.attributes.sort()) ||
        JSON.stringify(goldDoc.parents.sort()) !== JSON.stringify(genDoc.parents.sort());

      if (!hasDifferences) {
        // No differences - keep golden
        approved.push(goldDoc);
        continue;
      }

      // If there are differences, query the LLM Judge to evaluate
      console.log(`[Workflow] Differences found for <${genDoc.namespace}:${genDoc.tag}>. Invoking LLM Judge...`);

      const judgePrompt = `
You are an expert auditor of the ECMA-376 Office Open XML (OOXML) specification.
Your task is to act as a Judge and compare a newly GENERATED schema reference document against the existing GOLDEN schema reference document for the XML tag:
- Tag Name: "${genDoc.tag}"
- Domain: "${genDoc.domain}"

GOLDEN DOCUMENT:
${JSON.stringify(goldDoc, null, 2)}

GENERATED DOCUMENT:
${JSON.stringify(genDoc, null, 2)}

VALIDATION REPORT:
${JSON.stringify(validationReport, null, 2)}

You must evaluate the differences and make an authoritative decision.
Return a JSON object conforming exactly to this JSON schema:
{
  "type": "object",
  "properties": {
    "decision": { "type": "string", "enum": ["UPGRADE_GOLDEN", "KEEP_GOLDEN", "SUSPEND_FOR_REVIEW"] },
    "reasoning": { "type": "string" },
    "resolvedDoc": {
      "type": "object",
      "properties": {
        "tag": { "type": "string" },
        "namespace": { "type": "string" },
        "domain": { "type": "string" },
        "definition": { "type": "string" },
        "attributes": { "type": "array", "items": { "type": "string" } },
        "parents": { "type": "array", "items": { "type": "string" } },
        "citation": { "type": "string" },
        "sdkClass": { "type": "string" }
      },
      "required": ["tag", "namespace", "domain", "definition", "attributes", "parents", "citation", "sdkClass"]
    }
  },
  "required": ["decision", "reasoning", "resolvedDoc"]
}

Guidelines for your Decision:
1. "UPGRADE_GOLDEN": Choose this if the GENERATED document is more accurate or complete than the GOLDEN document. 
   - Example: The GENERATED citation is more specific or correct (e.g., '17.2.3' instead of '17.3.1.10' for w:document).
   - Example: The GENERATED document contains valid attributes or parents that were missing in the GOLDEN document.
   - If you choose UPGRADE_GOLDEN, the "resolvedDoc" should be the GENERATED document.

2. "KEEP_GOLDEN": Choose this if the GENERATED document contains hallucinations, incorrect citations, or is less accurate than the GOLDEN document.
   - If you choose KEEP_GOLDEN, the "resolvedDoc" should be the GOLDEN document.

3. "SUSPEND_FOR_REVIEW": Choose this if there is a major conflict, or if you are unsure which citation or definition is correct and require a human expert to decide.
   - If you choose SUSPEND_FOR_REVIEW, the "resolvedDoc" can be your best-effort merge, but the workflow will pause for human review.

Return ONLY the raw JSON block. No markdown wrapper, no explanations.
`;

      const escapedJudgePrompt = judgePrompt.replace(/"/g, '\\"').replace(/`/g, '\\`');
      const command = `jetski --print "${escapedJudgePrompt}"`;

      try {
        const stdout = execSync(command, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] });
        const jsonMatch = stdout.match(/\{[\s\S]*\}/);
        if (!jsonMatch) {
          throw new Error(`Could not find JSON in Judge output: ${stdout}`);
        }

        const judgeResult = JSON.parse(jsonMatch[0].trim());
        console.log(`[Workflow] Judge Decision for <${genDoc.namespace}:${genDoc.tag}>: ${judgeResult.decision}. Reasoning: ${judgeResult.reasoning}`);

        if (judgeResult.decision === 'UPGRADE_GOLDEN') {
          approved.push(judgeResult.resolvedDoc);
        } else if (judgeResult.decision === 'KEEP_GOLDEN') {
          approved.push(goldDoc); // Keep the original golden
        } else {
          // Suspend for review
          conflicts.push({
            tag: genDoc.tag,
            domain: genDoc.domain,
            namespace: goldDoc.namespace,
            reasoning: judgeResult.reasoning,
            golden: goldDoc,
            generated: genDoc,
            resolvedDoc: judgeResult.resolvedDoc
          });
        }
      } catch (error) {
        console.error(`[Workflow] Judge failed for <${genDoc.namespace}:${genDoc.tag}>, falling back to suspension:`, error);
        conflicts.push({
          tag: genDoc.tag,
          domain: genDoc.domain,
          namespace: goldDoc.namespace,
          reasoning: 'Judge execution failed.',
          golden: goldDoc,
          generated: genDoc
        });
      }

      // Avoid rate limits
      await new Promise(resolve => setTimeout(resolve, 1000));
    }

    if (conflicts.length > 0 && !resumeData) {
      console.log(`[Workflow] Found ${conflicts.length} conflicts requiring human review. Suspending workflow...`);
      return suspend({ conflicts });
    }

    // Resolve conflicts using resumeData
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const resolutions = resumeData ? (resumeData as any).resolutions : [];
    const finalApproved = [...approved, ...resolutions];

    return { approved: finalApproved };
  }
});

// 3. Commit Approved Data Step
const commitApprovedDataStep = createStep({
  id: 'commitApprovedData',
  inputSchema: z.object({
    approved: z.array(z.object({
      tag: z.string(),
      namespace: z.string(),
      domain: z.string(),
      definition: z.string(),
      attributes: z.array(z.string()),
      parents: z.array(z.string()),
      citation: z.string(),
      sdkClass: z.string()
    }))
  }),
  outputSchema: z.object({
    success: z.boolean(),
    count: z.number()
  }),
  execute: async ({ inputData }) => {
    const { approved } = inputData;
    
    const goldenPath = path.join(process.cwd(), 'public/rag-data.json');
    let golden: ReferenceDoc[] = [];
    if (fs.existsSync(goldenPath)) {
      golden = JSON.parse(fs.readFileSync(goldenPath, 'utf8'));
    }

    // Merge approved into golden
    for (const appDoc of approved) {
      const idx = golden.findIndex(g => g.tag === appDoc.tag && g.domain === appDoc.domain);
      if (idx > -1) {
        golden[idx] = appDoc as ReferenceDoc;
      } else {
        golden.push(appDoc as ReferenceDoc);
      }
    }

    // Write back to public/rag-data.json
    fs.writeFileSync(goldenPath, JSON.stringify(golden, null, 2), 'utf8');
    console.log(`[Workflow] Successfully updated golden dataset at: ${goldenPath}`);

    // Also write to services/staticKnowledgeBase.ts
    const kbPath = path.join(process.cwd(), 'services/staticKnowledgeBase.ts');
    const codeContent = `// This file is auto-generated by compileKB step in ooxml-rag-ingestion workflow.
// Do not edit this file manually.

export interface ReferenceDoc {
  tag: string;
  namespace: string;
  domain: 'docx' | 'xlsx' | 'pptx' | 'shared';
  definition: string;
  attributes: string[];
  parents: string[];
  citation?: string;
  sdkClass?: string;
}

export const KNOWLEDGE_BASE: ReferenceDoc[] = ${JSON.stringify(golden, null, 2)};
`;
    fs.writeFileSync(kbPath, codeContent, 'utf8');
    console.log(`[Workflow] Compiled static knowledge base to: ${kbPath}`);

    return { success: true, count: approved.length };
  }
});

// Define the Mastra Workflow for RAG Ingestion
export const ingestionWorkflow = new Workflow({
  id: 'ooxml-rag-ingestion',
  inputSchema: z.object({
    tags: z.array(z.object({
      tag: z.string(),
      namespace: z.string(),
      domain: z.string()
    }))
  }),
  outputSchema: z.object({
    success: z.boolean(),
    count: z.number()
  })
});

// Chain the steps and commit the workflow
ingestionWorkflow
  .then(generateSchemaStep)
  .then(evaluateAndDiffStep)
  .then(commitApprovedDataStep)
  .commit();

// Ensure the .mastra directory exists before initializing LibSQL
const dbDir = path.join(process.cwd(), '.mastra');
if (!fs.existsSync(dbDir)) {
  fs.mkdirSync(dbDir, { recursive: true });
}

// Initialize Mastra and export it
export const mastra = new Mastra({
  storage: new LibSQLStore({
    id: 'ooxml-explorer-store',
    url: `file:${path.join(dbDir, 'mastra.db')}`
  }),
  workflows: {
    'ooxml-rag-ingestion': ingestionWorkflow
  }
});
