import { Mastra } from '@mastra/core';
import { Workflow, createStep } from '@mastra/core/workflows';
import { LibSQLStore } from '@mastra/libsql';
import { execSync } from 'child_process';
import { z } from 'zod';
import * as fs from 'fs';
import * as path from 'path';
import { ReferenceDoc } from 'web/services/staticKnowledgeBase';
import { Agent } from '@mastra/core/agent';

// 1. Schema Generator Agent
export const schemaGeneratorAgent = new Agent({
  id: 'schema-generator',
  name: 'OOXML Schema Generator',
  instructions: `You are an expert in the ECMA-376 Office Open XML (OOXML) specification.
Your task is to generate a structured RAG reference document for the requested XML tag.
Conform exactly to the requested JSON schema.

Guidelines:
1. "namespace": You MUST return the short prefix (e.g., "w" for WordprocessingML, "r" for SpreadsheetML/Relationships, "p" for PresentationML) instead of the full XML namespace URI.
2. "definition": Provide a clear, precise explanation of what this element configures and its role. Keep it descriptive but concise (2-3 sentences).
3. "citation": You MUST reference the most specific section number in the ECMA-376 specification that defines this specific element. Do NOT use high-level parent section numbers. For example, the element 'document' is defined in 'ECMA-376 Part 1, Section 17.2.3', not the general 'Section 17.2' or the incorrect 'Section 17.3.1.10'. In case of multiple citations, pick the one that is the section heading and talks about that particular element only. Must match the format: "ECMA-376 Part X, Section Y.Z" (e.g. "ECMA-376 Part 1, Section 17.3.1.22"). Verify the exact section number.
4. "sdkClass": Provide the corresponding Microsoft Open XML SDK class name (e.g. "Paragraph" or "TableCell").`,
  model: {
    id: 'google/gemini-3.5-flash',
  },
});

// 2. Schema Auditor Agent (The Judge)
export const schemaAuditorAgent = new Agent({
  id: 'schema-auditor',
  name: 'OOXML Schema Auditor',
  instructions: `You are an expert auditor of the ECMA-376 Office Open XML (OOXML) specification.
Your task is to act as a Judge and compare a newly GENERATED schema reference document against the existing GOLDEN schema reference document for the requested XML tag.

You must evaluate the differences and make an authoritative decision.

Guidelines for your Decision:
1. "ACCEPT_GENERATED": Choose this if the GENERATED document is correct, acceptable, or an upgrade to the GOLDEN document. Phrasing variations or identical schemas MUST be accepted. The "resolvedDoc" should be either the generated document (if upgraded) or the golden document (if identical).
2. "REJECT_GENERATED": Choose this if the GENERATED document contains actual errors, hallucinations, or is incorrect compared to the GOLDEN document. It needs to be corrected.
3. "SUSPEND_FOR_REVIEW": Choose this if there is a major conflict requiring human review.`,
  model: {
    id: 'google/gemini-3.5-flash', // Use gemini-3.5-flash for auditing
  },
});


function getProjectRoot(): string {
  let currentDir = process.cwd();
  let metaUrl = '';
  try {
    metaUrl = import.meta.url;
    const filePath = new URL(import.meta.url).pathname;
    currentDir = path.dirname(filePath);
  } catch {
    // Ignore URL parsing errors
  }

  const startDir = currentDir;
  while (currentDir && currentDir !== '/') {
    if (fs.existsSync(path.join(currentDir, 'pnpm-workspace.yaml'))) {
      console.log(`[ProjectRoot] Resolved to monorepo root: ${currentDir} (from startDir: ${startDir})`);
      return currentDir;
    }
    currentDir = path.dirname(currentDir);
  }
  console.log(`[ProjectRoot] Fallback to cwd: ${process.cwd()}`);
  return process.cwd();
}

const PROJECT_ROOT = getProjectRoot();

function extractJson(text: string): string {
  const firstBrace = text.indexOf('{');
  if (firstBrace === -1) {
    throw new Error('No JSON object found in text');
  }

  let depth = 0;
  let inString = false;
  let escapeNext = false;

  for (let i = firstBrace; i < text.length; i++) {
    const char = text[i];

    if (escapeNext) {
      escapeNext = false;
      continue;
    }

    if (char === '\\') {
      escapeNext = true;
      continue;
    }

    if (char === '"') {
      inString = !inString;
      continue;
    }

    if (!inString) {
      if (char === '{') {
        depth++;
      } else if (char === '}') {
        depth--;
        if (depth === 0) {
          return text.substring(firstBrace, i + 1);
        }
      }
    }
  }
  throw new Error('Unbalanced braces in JSON text');
}



// Helper to log calibration defects to a markdown backlog
function logCalibrationDefect(
  tag: string,
  domain: string,
  golden: ReferenceDoc,
  generated: ReferenceDoc,
  reasoning: string
) {
  const feedbackPath = path.join(PROJECT_ROOT, 'apps/agents/CALIBRATION_FEEDBACK.md');
  const timestamp = new Date().toISOString().split('T')[0];
  
  const entry = `
## [${timestamp}] Ingestion Defect: <${golden.namespace}:${tag}> (${domain})
- **Judge Analysis**: ${reasoning}
- **Golden Citation**: \`${golden.citation || 'N/A'}\`
- **Generated Citation**: \`${generated.citation || 'N/A'}\`
- **Golden Definition**: ${golden.definition}
- **Generated Definition**: ${generated.definition}
- **Golden Attributes**: \`${JSON.stringify(golden.attributes)}\`
- **Generated Attributes**: \`${JSON.stringify(generated.attributes)}\`
---
`;

  fs.appendFileSync(feedbackPath, entry, 'utf8');
  console.log(`[Workflow] Logged calibration defect to: ${feedbackPath}`);
}

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

    // Load golden dataset to check for human reviewer notes
    const goldenPath = path.join(PROJECT_ROOT, 'apps/web/public/rag-data.json');
    let golden: ReferenceDoc[] = [];
    if (fs.existsSync(goldenPath)) {
      golden = JSON.parse(fs.readFileSync(goldenPath, 'utf8'));
    }

    for (const item of tags) {
      const { tag, namespace, domain } = item;
      console.log(`[Workflow] Processing <${namespace}:${tag}>...`);

      const goldDoc = golden.find(g => g.tag === tag && g.domain === domain);
      
      let prompt = `Generate a structured RAG reference document for the XML tag "${tag}" (namespace prefix: "${namespace}", domain: "${domain}").`;
      if (goldDoc?.reviewerNote) {
        prompt += `\n\nCRITICAL REVIEWER NOTE / CORRECTION GUIDE:\n"${goldDoc.reviewerNote}"\nYou MUST strictly follow this note when generating the definition, parents, attributes, and citation. Do not override this instruction under any circumstances.`;
      }

      try {
        const response = await schemaGeneratorAgent.generate(prompt, {
          structuredOutput: {
            schema: z.object({
              tag: z.string(),
              namespace: z.string(),
              domain: z.string(),
              definition: z.string(),
              attributes: z.array(z.string()),
              parents: z.array(z.string()),
              citation: z.string(),
              sdkClass: z.string()
            })
          }
        });

        if (response.object) {
          results.push(response.object);
          console.log(`[Workflow] Successfully processed <${namespace}:${tag}> via Agent`);
        } else {
          throw new Error('Agent failed to return structured object');
        }
      } catch (error) {
        console.error(`[Workflow] Failed to process <${namespace}:${tag}> via Agent:`, error);
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
    const goldenPath = path.join(PROJECT_ROOT, 'public/rag-data.json');
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
        approved.push(genDoc as ReferenceDoc);
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
      console.log(`[Workflow] Differences found for <${genDoc.namespace}:${genDoc.tag}>. Invoking LLM Judge Agent...`);

      const judgePrompt = `
Compare this newly GENERATED schema reference against the GOLDEN schema reference for XML tag "${genDoc.tag}" (domain: "${genDoc.domain}").

GOLDEN DOCUMENT:
${JSON.stringify(goldDoc, null, 2)}

GENERATED DOCUMENT:
${JSON.stringify(genDoc, null, 2)}

VALIDATION REPORT:
${JSON.stringify(validationReport, null, 2)}
`;

      try {
        const response = await schemaAuditorAgent.generate(judgePrompt, {
          structuredOutput: {
            schema: z.object({
              decision: z.enum(['ACCEPT_GENERATED', 'REJECT_GENERATED', 'SUSPEND_FOR_REVIEW']),
              reasoning: z.string(),
              resolvedDoc: z.object({
                tag: z.string(),
                namespace: z.string(),
                domain: z.string(),
                definition: z.string(),
                attributes: z.array(z.string()),
                parents: z.array(z.string()),
                citation: z.string(),
                sdkClass: z.string()
              })
            })
          }
        });

        const judgeResult = response.object;
        if (!judgeResult) {
          throw new Error('Auditor Agent failed to return structured decision');
        }

        console.log(`[Workflow] Judge Decision for <${genDoc.namespace}:${genDoc.tag}>: ${judgeResult.decision}. Reasoning: ${judgeResult.reasoning}`);

        if (judgeResult.decision === 'ACCEPT_GENERATED') {
          approved.push(judgeResult.resolvedDoc as ReferenceDoc);
        } else if (judgeResult.decision === 'REJECT_GENERATED') {
          // --- AUTONOMOUS RETRY LOOP (Phase 2) ---
          let attempts = 1;
          let currentGenDoc = genDoc;
          let currentJudgeResult = judgeResult;
          let healed = false;

          while (attempts <= 3) {
            console.log(`[Workflow] 🔄 Self-Correction Attempt ${attempts}/3 for <${genDoc.namespace}:${genDoc.tag}>...`);
            
            const feedbackPrompt = `
Your previous generated schema reference document for the tag "${genDoc.tag}" was REJECTED by the auditor.

REJECTION REASON:
"${currentJudgeResult.reasoning}"

PREVIOUS INCORRECT OUTPUT:
${JSON.stringify(currentGenDoc, null, 2)}

GOLDEN REFERENCE DOCUMENT (Correct Standard):
${JSON.stringify(goldDoc, null, 2)}

Please review the rejection reason, compare it with the Golden reference document, and regenerate a corrected schema reference document.
`;

            try {
              const retryResponse = await schemaGeneratorAgent.generate(feedbackPrompt, {
                structuredOutput: {
                  schema: z.object({
                    tag: z.string(),
                    namespace: z.string(),
                    domain: z.string(),
                    definition: z.string(),
                    attributes: z.array(z.string()),
                    parents: z.array(z.string()),
                    citation: z.string(),
                    sdkClass: z.string()
                  })
                }
              });

              if (!retryResponse.object) {
                throw new Error('Retry generation failed to return structured object');
              }
              
              currentGenDoc = retryResponse.object;
              
              // Re-evaluate with Judge
              console.log(`[Workflow] Re-evaluating retry attempt ${attempts} with Judge...`);
              const reJudgePrompt = `
Compare this newly GENERATED schema reference against the GOLDEN schema reference for XML tag "${genDoc.tag}" (domain: "${genDoc.domain}").

GOLDEN DOCUMENT:
${JSON.stringify(goldDoc, null, 2)}

GENERATED DOCUMENT:
${JSON.stringify(currentGenDoc, null, 2)}

VALIDATION REPORT:
${JSON.stringify(validationReport, null, 2)}
`;

              const reJudgeResponse = await schemaAuditorAgent.generate(reJudgePrompt, {
                structuredOutput: {
                  schema: z.object({
                    decision: z.enum(['ACCEPT_GENERATED', 'REJECT_GENERATED', 'SUSPEND_FOR_REVIEW']),
                    reasoning: z.string(),
                    resolvedDoc: z.object({
                      tag: z.string(),
                      namespace: z.string(),
                      domain: z.string(),
                      definition: z.string(),
                      attributes: z.array(z.string()),
                      parents: z.array(z.string()),
                      citation: z.string(),
                      sdkClass: z.string()
                    })
                  })
                }
              });

              const newJudgeResult = reJudgeResponse.object;
              if (!newJudgeResult) {
                throw new Error('Retry auditing failed to return structured decision');
              }

              console.log(`[Workflow] Retry ${attempts} Judge Decision: ${newJudgeResult.decision}. Reasoning: ${newJudgeResult.reasoning}`);
                  
              if (newJudgeResult.decision === 'ACCEPT_GENERATED') {
                approved.push(newJudgeResult.resolvedDoc as ReferenceDoc);
                healed = true;
                console.log(`[Workflow] ✅ Self-correction succeeded on attempt ${attempts}!`);
                break;
              }
              currentJudgeResult = newJudgeResult;
            } catch (err) {
              console.error(`[Workflow] Retry attempt ${attempts} failed:`, err);
            }
            
            attempts++;
            await new Promise(resolve => setTimeout(resolve, 2000));
          }

          if (!healed) {
            console.log(`[Workflow] ❌ Retries exhausted for <${genDoc.namespace}:${genDoc.tag}>. Keeping golden and logging defect.`);
            approved.push(goldDoc);
            
            // Log defect to CALIBRATION_FEEDBACK.md
            logCalibrationDefect(genDoc.tag, genDoc.domain, goldDoc, currentGenDoc as unknown as ReferenceDoc, currentJudgeResult.reasoning);
          }
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
        console.error(`[Workflow] Judge auditing failed for <${genDoc.namespace}:${genDoc.tag}>, falling back to suspension:`, error);
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
    
    const goldenPath = path.join(PROJECT_ROOT, 'apps/web/public/rag-data.json');
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
    const kbPath = path.join(PROJECT_ROOT, 'apps/web/services/staticKnowledgeBase.ts');
    
    // Filter golden to only keep high priority tags in staticKnowledgeBase.ts to keep bundle size small.
    // The rest of the tags are loaded on-demand in the app via public/rag-data.json.
    const highPriorityKB = golden.filter(doc => {
      return doc.priority === 'high' || [
        'document', 'body', 'p', 'r', 't', 'tbl', 'tr', 'tc', 'cantSplit', 'tblHeader',
        'worksheet', 'sheetData', 'row', 'c', 'v', 'f', 'sst', 'si',
        'presentation', 'sld', 'sldLayout', 'sldMaster', 'sp', 'txBody',
        'Relationships', 'Relationship', 'Types', 'Override', 'Default'
      ].includes(doc.tag);
    });

    const codeContent = `// This file is auto-generated by the ooxml-rag-ingestion workflow.
// To regenerate this file, run: pnpm --filter mastra ingest
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
  reviewerNote?: string;
  priority?: 'high' | 'low';
}

export const KNOWLEDGE_BASE: ReferenceDoc[] = ${JSON.stringify(highPriorityKB, null, 2)};
`;
    fs.writeFileSync(kbPath, codeContent, 'utf8');
    console.log(`[Workflow] Compiled ${highPriorityKB.length} high-priority tags to: ${kbPath}`);

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
const dbDir = path.join(PROJECT_ROOT, '.mastra');
if (!fs.existsSync(dbDir)) {
  fs.mkdirSync(dbDir, { recursive: true });
}

// Initialize Mastra and export it
export const mastra = new Mastra({
  storage: new LibSQLStore({
    id: 'ooxml-explorer-store',
    url: `file:${path.join(dbDir, 'mastra.db')}`
  }),
  agents: {
    'schema-generator': schemaGeneratorAgent,
    'schema-auditor': schemaAuditorAgent
  },
  workflows: {
    'ooxml-rag-ingestion': ingestionWorkflow
  }
});

export default mastra;
