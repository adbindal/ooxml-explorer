import { GoogleGenAI } from "@google/genai";

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
            model: 'gemini-3-flash-preview',
            contents: 'Test connection',
        });
        
        return { success: true, message: "Authenticated successfully." };
    } catch (e: unknown) {
        const error = e as Error;
        return { success: false, message: error.message || "Connection failed." };
    }
};

// --- EDITOR MODE TYPES ---
export interface EditorFileContext {
    fileName: string;
    content: string;
}

// --- DIFF MODE TYPES ---
export interface DiffFileContext {
    fileName: string;
    original: string | null;
    modified: string | null;
}

/**
 * Analyzes a single file (with optional context files) in Editor Mode.
 */
export const analyzeFile = async (
    files: EditorFileContext[],
    mode: 'explain' | 'technical'
): Promise<string> => {
  if (!files || files.length === 0) {
    return "No files provided for analysis.";
  }
  try {
    const ai = getAI();
    
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
        
        1. **Purpose**: What is this file responsible for in the OOXML package? (Reference ECMA-376 concepts but keep it accessible).
        2. **Key Data**: Summarize the important configuration or data it holds.
        3. **Context**: If related files are provided, explain how this file interacts with them.
        
        Return the response in concise Markdown. Use bullet points.
        `;
    } else if (mode === 'technical') {
        systemInstruction = "You are a Senior OOXML Engineer with mastery of the ECMA-376 specification. Your goal is to debug, validate, and explain XML structures technically.";
        prompt = `
        Perform a **technical deep dive** on "${mainFileName}"${relatedCount > 0 ? ` using ${relatedCount} related file(s) for reference` : ''}.
        
        1. **Structure Analysis**: Analyze the XML tags and attributes. Are they standard ECMA-376 elements?
        2. **Validation**: Check for common issues (namespaces, required attributes, nesting).
        3. **Relationships**: If referencing IDs (rId), do they look consistent?
        
        Provide a detailed technical breakdown suitable for a developer.
        `;
    }

    const response = await ai.models.generateContent({
      model: 'gemini-3.1-pro-preview',
      config: { systemInstruction },
      contents: `
      ${prompt}

      ${filesContext}
      `
    });
    
    return response.text || "No analysis generated.";
  } catch (error: unknown) {
    const err = error as Error;
    if (err.message === 'API_KEY_MISSING') {
        throw err;
    }
    console.error("Gemini Error:", error);
    return "Error analyzing content. Please check your network connection or API quota.";
  }
};

/**
 * Analyzes file differences in Diff Mode.
 */
export const analyzeDiff = async (
    files: DiffFileContext[], 
    mode: 'summary' | 'technical'
): Promise<string> => {
  if (!files || files.length === 0) {
    return "No files provided for diff analysis.";
  }
  try {
    const ai = getAI();
    
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
        
        Use your knowledge of ECMA-376 to understand the structure, but focus the output on **functional impact** for a non-technical user:
        - What would the user see differently in Word/Excel/PowerPoint?
        - Filter out minor XML noise (like random ID updates) unless they affect document behavior.
        - Correlate changes across files if provided (e.g. a style change in styles.xml affecting document.xml).
        - Summarize the intent of the change in plain English.

        Keep it concise and easy to read. Use bullet points.
        `;
    } else if (mode === 'technical') {
        systemInstruction = "You are a Senior OOXML Engineer with mastery of the ECMA-376 specification. Your goal is to debug and validate XML changes against the standard.";
        prompt = `
        Perform a **technical deep dive** on the diff for "${mainFileName}"${relatedCount > 0 ? ` and ${relatedCount} related files` : ''}.
        
        Analyze the specific XML mechanics citing ECMA-376 concepts where relevant:
        - Which attributes or tags were modified and what is their role in the standard?
        - Check for cross-file integrity (e.g. if a Relationship ID is used in document.xml, does it exist in .rels?).
        - Are there potential side effects?
        - Does this change look safe and compliant?
        
        Provide a detailed technical breakdown suitable for a developer.
        `;
    }

    const response = await ai.models.generateContent({
        model: 'gemini-3.1-pro-preview', 
        config: { systemInstruction },
        contents: `
        ${prompt}

        ${filesContext}
        `
    });

    return response.text || "No analysis generated.";
  } catch (error: unknown) {
    const err = error as Error;
    if (err.message === 'API_KEY_MISSING') {
        throw err;
    }
    console.error("Gemini Error:", error);
    return "Error analyzing diff. Please check your network connection or API quota.";
  }
}