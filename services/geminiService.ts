import { GoogleGenAI } from "@google/genai";
import { z } from "zod";
import { getActiveAIProvider } from "./aiProvider";
// Type-only: aiService imports getApiKey from this module, so a value import here
// would close a runtime cycle. `import type` is erased at compile time, so it cannot.
// ComputedEvidence wants a home of its own once the Finding type lands.
import type { ComputedEvidence } from "./aiService";
import {
  CLOUD_CONTENT_BUDGET_CHARS,
  allocateContentBudget,
  getLocalContentBudgetChars,
  renderContentSnippet
} from "./promptBudget";

const cleanAndParseJson = (text: string) => {
  let cleaned = text.trim();
  if (cleaned.startsWith('```')) {
    cleaned = cleaned.replace(/^```(?:json)?\n?/, '').replace(/\n?```$/, '').trim();
  }
  return JSON.parse(cleaned);
};

/**
 * Prompts the on-device model for a JSON response and validates it against `schema`.
 *
 * Chrome's local Prompt API is not a constrained-decoding API like Gemini Cloud's
 * `responseSchema` - it just follows instructions in the prompt text. Describing the
 * desired shape as an abstract JSON Schema (as the cloud path does) confuses small
 * on-device models into echoing the schema definition back instead of filling it in,
 * so the prompt must show a concrete filled-in example instead. Even so, on-device
 * models occasionally wrap the response in prose or partially miss the shape, so this
 * makes one corrective retry (reusing the same session) before giving up.
 *
 * The prompt is supplied as a builder rather than a string because the amount of file
 * content that fits can only be known once the session exists - the window size varies
 * by Chrome version and the system prompt has already consumed part of it. See
 * ./promptBudget.
 */
const promptLocalModelForJson = async <T>(
  systemInstruction: string,
  buildUserPrompt: (contentBudgetChars: number) => string,
  schema: z.ZodType<T>
): Promise<T> => {
  const session = await window.LanguageModel!.create({ systemPrompt: systemInstruction });
  try {
    const userPrompt = buildUserPrompt(getLocalContentBudgetChars(session));
    const resultText = await session.prompt(userPrompt);
    try {
      return schema.parse(cleanAndParseJson(resultText));
    } catch {
      const retryText = await session.prompt(
        "That response was not valid. Reply again with ONLY the JSON data object itself - not the schema, not markdown, not any explanation."
      );
      try {
        return schema.parse(cleanAndParseJson(retryText));
      } catch {
        throw new Error("Local AI returned a response that didn't match the expected format. Try again, or switch to Cloud AI in Settings.");
      }
    }
  } finally {
    session.destroy();
  }
};

const STORAGE_KEY = 'ooxml_explorer_api_key';

export const getApiKey = (): string | undefined => {
  // 1. Process Env (Build time / Server injected)
  if (typeof process !== 'undefined' && process.env && process.env.API_KEY) {
    const envKey = process.env.API_KEY;
    if (envKey !== 'undefined' && envKey !== '') {
      return envKey;
    }
  }
  // 2. Local Storage (User provided)
  const key = localStorage.getItem(STORAGE_KEY);
  if (key === null || key === 'undefined' || key === '') {
    return undefined;
  }
  return key;
};

export const setApiKey = (key: string) => {
    localStorage.setItem(STORAGE_KEY, key);
};

export const clearApiKey = () => {
    localStorage.removeItem(STORAGE_KEY);
};

const getAI = () => {
  const apiKey = getApiKey();
  
  if (!apiKey) {
    throw new Error("API_KEY_MISSING");
  }
  return new GoogleGenAI({ apiKey });
};

export const testConnection = async (): Promise<{ success: boolean; message: string }> => {
    try {
        const apiKey = getApiKey();
        if (!apiKey) {
            return { success: false, message: "API Key not found in storage or env." };
        }
        
        const ai = new GoogleGenAI({ apiKey });
        // Lightweight call to verify authentication and quota
        await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: 'Test connection',
        });
        
        return { success: true, message: "Authenticated successfully." };
    } catch (e: unknown) {
        const error = e as Error;
        return { success: false, message: error.message || "Connection failed." };
    }
};

// --- EDITOR MODE TYPES & SCHEMAS ---
// Request shape is Zod-derived (not a hand-written interface) so the same schema
// validates the request at the function boundary and stays in sync with the type.
// See AGENTS.md "AI Request/Response Validation" for the project-wide rule.
export const EditorFileContextSchema = z.object({
    fileName: z.string(),
    content: z.string()
});
export type EditorFileContext = z.infer<typeof EditorFileContextSchema>;

export const AIAnalysisSchema = z.object({
  summary: z.string(),
  criticalIssues: z.array(z.object({
    issue: z.string(),
    impact: z.string(),
    remediation: z.string()
  })),
  keyElements: z.array(z.object({
    tag: z.string(),
    purpose: z.string()
  }))
});

export type AIAnalysis = z.infer<typeof AIAnalysisSchema>;

const AI_ANALYSIS_RESPONSE_SCHEMA = {
  type: 'OBJECT',
  properties: {
    summary: {
      type: 'STRING',
      description: 'A concise explanation of the file\'s purpose and role in the OOXML package.'
    },
    criticalIssues: {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        properties: {
          issue: { type: 'STRING', description: 'The problem or warning found in the XML structure.' },
          impact: { type: 'STRING', description: 'The functional impact on Microsoft Office (e.g. file corruption, formatting loss).' },
          remediation: { type: 'STRING', description: 'How to fix this issue in the XML.' }
        },
        required: ['issue', 'impact', 'remediation']
      },
      description: 'List of critical compliance, namespace, or structural issues found. Empty array if none.'
    },
    keyElements: {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        properties: {
          tag: { type: 'STRING', description: 'The XML tag name.' },
          purpose: { type: 'STRING', description: 'What this specific element configures in plain English.' }
        },
        required: ['tag', 'purpose']
      },
      description: 'Key OOXML elements found in this file with their purpose.'
    }
  },
  required: ['summary', 'criticalIssues', 'keyElements']
};

// Concrete filled-in example for local-model prompting (see promptLocalModelForJson).
// The example's content is unrelated to any real file on purpose, so the model can't
// mistake it for the answer.
const AI_ANALYSIS_LOCAL_EXAMPLE = JSON.stringify({
  summary: "This file defines the styles used throughout the document, such as heading and paragraph formatting.",
  criticalIssues: [
    { issue: "Missing xml:space attribute", impact: "Leading or trailing whitespace in text runs may be trimmed unexpectedly.", remediation: "Add xml:space=\"preserve\" to the affected <w:t> element." }
  ],
  keyElements: [
    { tag: "w:style", purpose: "Defines a single named style (e.g. Heading1) and its formatting properties." }
  ]
} satisfies AIAnalysis, null, 2);

// --- DIFF MODE TYPES & SCHEMAS ---
export const DiffFileContextSchema = z.object({
    fileName: z.string(),
    original: z.string().nullable(),
    modified: z.string().nullable()
});
export type DiffFileContext = z.infer<typeof DiffFileContextSchema>;

export const AIDiffSchema = z.object({
  summary: z.string(),
  changesList: z.array(z.object({
    element: z.string(),
    changeType: z.enum(['added', 'modified', 'deleted']),
    description: z.string(),
    visualImpact: z.string()
  }))
});

export type AIDiffAnalysis = z.infer<typeof AIDiffSchema>;

const AI_DIFF_RESPONSE_SCHEMA = {
  type: 'OBJECT',
  properties: {
    summary: {
      type: 'STRING',
      description: 'A user-friendly plain English summary of the functional impact of the changes.'
    },
    changesList: {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        properties: {
          element: { type: 'STRING', description: 'The XML element or setting that was modified.' },
          changeType: { type: 'STRING', enum: ['added', 'modified', 'deleted'], description: 'The type of modification.' },
          description: { type: 'STRING', description: 'A clear description of what changed and its technical role.' },
          visualImpact: { type: 'STRING', description: 'What the user will see differently in Word/Excel/PowerPoint.' }
        },
        required: ['element', 'changeType', 'description', 'visualImpact']
      },
      description: 'Structured list of key modifications.'
    }
  },
  required: ['summary', 'changesList']
};

// Concrete filled-in example for local-model prompting (see promptLocalModelForJson).
const AI_DIFF_LOCAL_EXAMPLE = JSON.stringify({
  summary: "The default paragraph spacing was increased, making the document look less dense.",
  changesList: [
    { element: "w:spacing", changeType: "modified", description: "The 'after' spacing value on the default paragraph style was increased from 100 to 200 twips.", visualImpact: "Paragraphs will have more vertical space between them." }
  ]
} satisfies AIDiffAnalysis, null, 2);

/**
 * Analyzes a single file (with optional context files) in Editor Mode.
 */
export const analyzeFile = async (
    files: EditorFileContext[],
    mode: 'explain' | 'technical'
): Promise<AIAnalysis> => {
  const parsedFiles = z.array(EditorFileContextSchema).min(1, "No files provided for analysis.").safeParse(files);
  if (!parsedFiles.success) {
    throw new Error("No files provided for analysis.");
  }
  files = parsedFiles.data;
  try {
    const activeProvider = await getActiveAIProvider();

    // Built against a *total* budget shared across every file, not a per-file cap -
    // otherwise N files means N times the limit, which overflows the on-device window.
    const buildFilesContext = (budgetChars: number): string => {
      const limits = allocateContentBudget(files.map(f => f.content.length), budgetChars);
      return files.map((f, index) => `
        --- FILE ${index + 1}: ${f.fileName} ---
        \`\`\`xml
        ${renderContentSnippet(f.content, limits[index])}
        \`\`\`
        \n`).join('');
    };

    const mainFileName = files[0].fileName;
    const relatedCount = files.length - 1;

    let systemInstruction = "";
    let prompt = "";

    if (mode === 'explain') {
        systemInstruction = "You are a helpful Document Assistant with expertise in the ECMA-376 specification. Your goal is to explain the purpose of OOXML files to a user.";
        prompt = `
        Analyze the provided OOXML file(s). The primary file being viewed is "${mainFileName}"${relatedCount > 0 ? `, with ${relatedCount} related file(s) for context` : ''}.
        
        1. **Purpose/Summary**: Explain the role of this file in the OOXML package in clear, accessible language.
        2. **Critical Issues**: Audit the XML structure. Identify if there are any invalid tags, missing required namespaces, broken schema references, or elements that could cause file corruption or formatting loss.
        3. **Key Elements**: Identify the most important XML tags in the file and explain their functional purpose.
        `;
    } else if (mode === 'technical') {
        systemInstruction = "You are a Senior OOXML Engineer with mastery of the ECMA-376 specification. Your goal is to debug, validate, and explain XML structures technically.";
        prompt = `
        Perform a technical deep dive audit on "${mainFileName}"${relatedCount > 0 ? ` using ${relatedCount} related file(s) for reference` : ''}.
        
        1. **Summary/Purpose**: Explain the technical role of this file in the package schema.
        2. **Critical Issues**: Run a strict validation check for ECMA-376 compliance (namespaces, required attributes, nesting rules, relationship consistency). Citing exact issues.
        3. **Key Elements**: Extract the main structural elements, listing their tag names and technical configurations.
        `;
    }

    if (activeProvider === 'chrome-local') {
      return await promptLocalModelForJson(systemInstruction, (budgetChars) => `
      ${prompt}

      Respond with ONLY a single JSON object - no markdown formatting, no backticks, no commentary
      before or after. Here is an EXAMPLE of a correctly formatted response for a different file
      (its content is unrelated - write your own analysis for the actual file below):
      ${AI_ANALYSIS_LOCAL_EXAMPLE}

      Here is the input file content:
      ${buildFilesContext(budgetChars)}
      `, AIAnalysisSchema);
    }

    const ai = getAI();

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      config: { 
          systemInstruction,
          responseMimeType: 'application/json',
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          responseSchema: AI_ANALYSIS_RESPONSE_SCHEMA as any
      },
      contents: `
      ${prompt}

      ${buildFilesContext(CLOUD_CONTENT_BUDGET_CHARS)}
      `
    });

    const text = response.text || "{}";
    const rawData = JSON.parse(text);

    // Validate with Zod to guarantee runtime type safety
    return AIAnalysisSchema.parse(rawData);
  } catch (error: unknown) {
    const err = error as Error;
    if (err.message === 'API_KEY_MISSING') {
        throw err;
    }
    console.error("Gemini Error in analyzeFile:", error);
    throw new Error(err.message || "Failed to analyze file.");
  }
};

/**
 * Analyzes file differences in Diff Mode.
 *
 * `computed` carries the semantic diff produced by `services/ooxmlDiff.ts` — an actual
 * derivation over both packages, not a retrieval. It matters most here of anywhere in
 * the app: a raw XML diff of two saves of the same document is mostly rewritten
 * revision ids and text redistributed across runs, and a model shown only the raw text
 * will describe that noise as change. The computed block states plainly which
 * differences are real, and whether the two files are equivalent despite them.
 */
export const analyzeDiff = async (
    files: DiffFileContext[],
    mode: 'summary' | 'technical',
    computed?: ComputedEvidence | null
): Promise<AIDiffAnalysis> => {
  const parsedFiles = z.array(DiffFileContextSchema).min(1, "No files provided for diff analysis.").safeParse(files);
  if (!parsedFiles.success) {
    throw new Error("No files provided for diff analysis.");
  }
  files = parsedFiles.data;
  try {
    const activeProvider = await getActiveAIProvider();

    // Same contract as the element explainer: computed facts outrank the model's
    // reading of the raw XML, and what could not be established is named so it cannot
    // be filled in with a plausible guess.
    const computedBlock = computed && computed.lines.length > 0
      ? `
      [COMPUTED DIFF - derived from both packages, authoritative]
      ${computed.lines.join('\n      ')}
      ${computed.unresolved.length > 0
        ? `\n      The following could NOT be established. Do not state or guess them:\n      ${computed.unresolved.join('\n      ')}`
        : ''}

      Treat the computed diff above as authoritative. Where it disagrees with your own
      reading of the raw XML below, the computed diff is correct. In particular, do not
      report a difference it does not list - the raw text differs in ways that carry no
      meaning, and those have already been filtered out.
      `
      : '';

    // A diff carries two versions per file, so the budget is split across 2N bodies -
    // the previous per-side cap let a single two-file diff reach four times its limit.
    const buildFilesContext = (budgetChars: number): string => {
      const sides = files.flatMap(f => [f.original?.length ?? 0, f.modified?.length ?? 0]);
      const limits = allocateContentBudget(sides, budgetChars);
      return files.map((f, index) => {
        const snipA = f.original
          ? renderContentSnippet(f.original, limits[index * 2])
          : "(File did not exist)";
        const snipB = f.modified
          ? renderContentSnippet(f.modified, limits[index * 2 + 1])
          : "(File deleted)";
        return `
        --- FILE ${index + 1}: ${f.fileName} ---
        [ORIGINAL VERSION]:
        \`\`\`xml
        ${snipA}
        \`\`\`

        [MODIFIED VERSION]:
        \`\`\`xml
        ${snipB}
        \`\`\`
        \n`;
      }).join('');
    };

    let systemInstruction = "";
    let prompt = "";

    const mainFileName = files[0].fileName;
    const relatedCount = files.length - 1;

    if (mode === 'summary') {
        systemInstruction = "You are a helpful Document Assistant with expertise in the ECMA-376 specification. Your goal is to translate technical XML structures into simple, user-friendly explanations.";
        prompt = `
        Compare the provided versions of the OOXML file(s). The primary file being viewed is "${mainFileName}"${relatedCount > 0 ? `, with ${relatedCount} related file(s) for context` : ''}.
        
        1. **Summary**: Provide a plain English summary of the functional impact of these changes on a non-technical user.
        2. **Changes List**: Identify each key modified XML element. List the element tag, describe the modification (added/modified/deleted), explain what it technically configures, and detail the visual/functional impact.
        `;
    } else if (mode === 'technical') {
        systemInstruction = "You are a Senior OOXML Engineer with mastery of the ECMA-376 specification. Your goal is to debug and validate XML changes against the standard.";
        prompt = `
        Perform a technical deep dive audit on the diff for "${mainFileName}"${relatedCount > 0 ? ` and ${relatedCount} related files` : ''}.
        
        1. **Summary**: Summarize the technical purpose of these modifications.
        2. **Changes List**: Extract all key attribute or tag modifications. For each, describe the change type, provide the exact technical explanation, and check for schema compliance and side effects.
        `;
    }

    if (activeProvider === 'chrome-local') {
      return await promptLocalModelForJson(systemInstruction, (budgetChars) => `
      ${prompt}
      ${computedBlock}

      Respond with ONLY a single JSON object - no markdown formatting, no backticks, no commentary
      before or after. Here is an EXAMPLE of a correctly formatted response for a different diff
      (its content is unrelated - write your own analysis for the actual diff below):
      ${AI_DIFF_LOCAL_EXAMPLE}

      Here are the original/modified files:
      ${buildFilesContext(budgetChars)}
      `, AIDiffSchema);
    }

    const ai = getAI();

    const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash', 
        config: { 
            systemInstruction,
            responseMimeType: 'application/json',
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            responseSchema: AI_DIFF_RESPONSE_SCHEMA as any
        },
        contents: `
        ${prompt}
        ${computedBlock}

        ${buildFilesContext(CLOUD_CONTENT_BUDGET_CHARS)}
        `
    });

    const text = response.text || "{}";
    const rawData = JSON.parse(text);
    
    // Validate with Zod to guarantee runtime type safety
    return AIDiffSchema.parse(rawData);
  } catch (error: unknown) {
    const err = error as Error;
    if (err.message === 'API_KEY_MISSING') {
        throw err;
    }
    console.error("Gemini Error in analyzeDiff:", error);
    throw new Error(err.message || "Failed to analyze diff.");
  }
};