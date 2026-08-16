import { GoogleGenAI } from "@google/genai";
import { z } from "zod";
import { getActiveAIProvider } from "./aiProvider";

const cleanAndParseJson = (text: string) => {
  let cleaned = text.trim();
  if (cleaned.startsWith('```')) {
    cleaned = cleaned.replace(/^```(?:json)?\n?/, '').replace(/\n?```$/, '').trim();
  }
  return JSON.parse(cleaned);
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
export interface EditorFileContext {
    fileName: string;
    content: string;
}

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

// --- DIFF MODE TYPES & SCHEMAS ---
export interface DiffFileContext {
    fileName: string;
    original: string | null;
    modified: string | null;
}

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

/**
 * Analyzes a single file (with optional context files) in Editor Mode.
 */
export const analyzeFile = async (
    files: EditorFileContext[],
    mode: 'explain' | 'technical'
): Promise<AIAnalysis> => {
  if (!files || files.length === 0) {
    throw new Error("No files provided for analysis.");
  }
  try {
    const activeProvider = await getActiveAIProvider();
    
    // Construct the prompt content dynamically based on multiple files
    let filesContext = "";
    files.forEach((f, index) => {
        const snippet = f.content.slice(0, 8000);
        filesContext += `
        --- FILE ${index + 1}: ${f.fileName} ---
        \`\`\`xml
        ${snippet}
        \`\`\`
        \n`;
    });

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
      const localPrompt = `
      ${prompt}

      You MUST respond ONLY with a valid JSON string matching this exact JSON schema:
      ${JSON.stringify(AI_ANALYSIS_RESPONSE_SCHEMA, null, 2)}

      Do NOT include any markdown formatting, backticks, or extra text outside of the JSON block.

      Here is the input file content:
      ${filesContext}
      `;

      const session = await window.LanguageModel!.create({
        systemPrompt: systemInstruction
      });
      try {
        const resultText = await session.prompt(localPrompt);
        const parsed = cleanAndParseJson(resultText);
        return AIAnalysisSchema.parse(parsed);
      } finally {
        session.destroy();
      }
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

      ${filesContext}
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
 */
export const analyzeDiff = async (
    files: DiffFileContext[], 
    mode: 'summary' | 'technical'
): Promise<AIDiffAnalysis> => {
  if (!files || files.length === 0) {
    throw new Error("No files provided for diff analysis.");
  }
  try {
    const activeProvider = await getActiveAIProvider();
    
    // Construct the prompt content dynamically based on multiple files
    let filesContext = "";
    files.forEach((f, index) => {
        const snipA = f.original ? f.original.slice(0, 8000) : "(File did not exist)";
        const snipB = f.modified ? f.modified.slice(0, 8000) : "(File deleted)";
        
        filesContext += `
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
    });

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
      const localPrompt = `
      ${prompt}

      You MUST respond ONLY with a valid JSON string matching this exact JSON schema:
      ${JSON.stringify(AI_DIFF_RESPONSE_SCHEMA, null, 2)}

      Do NOT include any markdown formatting, backticks, or extra text outside of the JSON block.

      Here are the original/modified files:
      ${filesContext}
      `;

      const session = await window.LanguageModel!.create({
        systemPrompt: systemInstruction
      });
      try {
        const resultText = await session.prompt(localPrompt);
        const parsed = cleanAndParseJson(resultText);
        return AIDiffSchema.parse(parsed);
      } finally {
        session.destroy();
      }
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

        ${filesContext}
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