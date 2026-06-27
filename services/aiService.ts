import { GoogleGenAI } from "@google/genai";
import { useAppStore } from "../store/appStore";
import { getApiKey } from "./geminiService";

/**
 * Detects whether the local AI provider (Chrome Gemini Nano) is available and ready.
 * Falls back to cloud if unavailable, unsupported, or still downloading.
 */
export const getActiveAIProvider = async (): Promise<'local' | 'cloud'> => {
  const preferredProvider = useAppStore.getState().ui.aiProvider;
  
  if (preferredProvider === 'chrome-local') {
    if (typeof window !== 'undefined' && window.LanguageModel) {
      try {
        const availability = await window.LanguageModel.availability();
        if (availability === 'available') {
          return 'local';
        }
      } catch (e) {
        console.warn("[AI Service] Error checking local AI availability, falling back to cloud:", e);
      }
    }
    return 'cloud'; // Fallback
  }
  
  return 'cloud';
};

let cachedAi: GoogleGenAI | null = null;
let cachedApiKey: string | null = null;

export const getAiClient = (apiKey: string): GoogleGenAI => {
  if (!cachedAi || cachedApiKey !== apiKey) {
    cachedAi = new GoogleGenAI({ apiKey });
    cachedApiKey = apiKey;
  }
  return cachedAi;
};

/**
 * Layer 1: Core Explainer AI Service.
 * Explains the purpose of an XML element using the active provider.
 */
export const explainElement = async (
  tagName: string,
  rawXml: string,
  fileType: 'docx' | 'xlsx' | 'pptx'
): Promise<string> => {
  const provider = await getActiveAIProvider();
  
  // Clean tag name and XML
  const cleanTagName = tagName.trim();
  const cleanXml = rawXml.trim();

  // Simple prompt for Layer 1 (to be enhanced with RAG in Layer 3)
  const systemInstruction = "You are a helpful Document Assistant with expertise in the ECMA-376 specification. Explain the purpose of XML tags to the user.";
  const prompt = `
  Explain the purpose of the XML tag "${cleanTagName}" in the context of a ${fileType.toUpperCase()} document.
  
  Here is the raw XML snippet:
  \`\`\`xml
  ${cleanXml}
  \`\`\`
  
  Explain:
  1. What this tag configures in plain English.
  2. Its role and importance in the document.
  `;

  if (provider === 'local') {
    console.log(`[AI Service] Running Explainer locally for <${cleanTagName}>`);
    const session = await window.LanguageModel!.create({
      systemPrompt: systemInstruction
    });
    try {
      const response = await session.prompt(prompt);
      return response.trim();
    } finally {
      session.destroy();
    }
  } else {
    console.log(`[AI Service] Running Explainer in cloud for <${cleanTagName}>`);
    const apiKey = getApiKey();
    if (!apiKey) {
      throw new Error("API_KEY_MISSING");
    }
    const ai = getAiClient(apiKey);
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      config: {
        systemInstruction
      },
      contents: prompt
    });
    return (response.text || "").trim();
  }
};
