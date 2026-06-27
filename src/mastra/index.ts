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

// 2. Evaluate and Diff Step (Decision Engine & HITL Suspension)
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

    const approved = [];
    const conflicts = [];

    for (const genDoc of results) {
      const goldDoc = golden.find(g => g.tag === genDoc.tag && g.domain === genDoc.domain);

      if (!goldDoc) {
        // New tag - auto approve
        approved.push(genDoc);
        continue;
      }

      // Compare citations
      const citationMatch = goldDoc.citation === genDoc.citation;

      if (citationMatch) {
        // Citations match, we can auto-merge attributes and parents
        const mergedAttributes = Array.from(new Set([...goldDoc.attributes, ...genDoc.attributes]));
        const mergedParents = Array.from(new Set([...goldDoc.parents, ...genDoc.parents]));
        
        approved.push({
          ...genDoc,
          namespace: goldDoc.namespace, // Prefer golden prefix
          attributes: mergedAttributes,
          parents: mergedParents
        });
      } else {
        // Citation mismatch - needs human review
        conflicts.push({
          tag: genDoc.tag,
          domain: genDoc.domain,
          namespace: goldDoc.namespace,
          golden: {
            definition: goldDoc.definition,
            citation: goldDoc.citation,
            sdkClass: goldDoc.sdkClass,
            attributes: goldDoc.attributes,
            parents: goldDoc.parents
          },
          generated: {
            definition: genDoc.definition,
            citation: genDoc.citation,
            sdkClass: genDoc.sdkClass,
            attributes: genDoc.attributes,
            parents: genDoc.parents
          }
        });
      }
    }

    if (conflicts.length > 0 && !resumeData) {
      console.log(`[Workflow] Found ${conflicts.length} citation conflicts. Suspending workflow for review...`);
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
