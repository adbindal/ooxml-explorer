import { ReferenceDoc } from './ragDb';
import { querySchemaFromStorage, searchSchemasInStorage } from './storageService';

/**
 * Uses the local Gemini Nano model (if available) to extract 1-2 search keywords
 * from a natural language query.
 */
const extractKeywordsWithLocalAI = async (query: string): Promise<string> => {
  if (typeof window !== 'undefined' && window.LanguageModel) {
    try {
      const availability = await window.LanguageModel.availability();
      if (availability === 'available') {
        const session = await window.LanguageModel.create({
          systemPrompt: 'You are a technical keyword extractor. Extract 1-2 primary XML tag names or schema terms from the query. Output ONLY the raw keywords separated by spaces, no punctuation.'
        });
        const result = await session.prompt(query);
        session.destroy();
        return result.trim();
      }
    } catch (e) {
      console.warn('[RAG Router] Local AI keyword extraction failed:', e);
    }
  }
  return query;
};

/**
 * RAG Context Retrieval & Router.
 * Queries the client-side IndexedDB database for a tag and filters by the active editor context.
 * Falls back to LLM-assisted keyword search for natural language queries.
 * Supports runtime self-healing overrides via localStorage.
 */
export const getRagContext = async (
  tagName: string,
  context: { fileType: 'docx' | 'xlsx' | 'pptx'; namespace: string }
): Promise<string> => {
  const cleanTagName = tagName.trim();
  
  // 1. Check for runtime self-healing overrides in localStorage first
  if (typeof window !== 'undefined' && typeof localStorage !== 'undefined') {
    const overrideKey = `ooxml_rag_override_${context.fileType}_${cleanTagName}`;
    const localOverride = localStorage.getItem(overrideKey);
    if (localOverride) {
      console.log(`[RAG Router] Applying local self-healing override for <${cleanTagName}>`);
      return `
[ECMA-376 Specification Context (Self-Healed Override)]
- Tag Name: <${context.namespace}:${cleanTagName}>
- Schema Domain: ${context.fileType.toUpperCase()}
- Definition: ${localOverride}
- Source: Runtime Local Patch
`;
    }
  }

  // 2. Try direct lookup in IndexedDB
  let match: ReferenceDoc | null = await querySchemaFromStorage(cleanTagName, context.fileType);

  // 3. Fallback: If not found and looks like a natural language query, run LLM keyword search
  if (!match && (cleanTagName.includes(' ') || cleanTagName.length > 15)) {
    console.log(`[RAG Router] Tag <${cleanTagName}> not found. Running LLM-assisted keyword search...`);
    const keywords = await extractKeywordsWithLocalAI(cleanTagName);
    const searchResults = await searchSchemasInStorage(keywords, context.fileType);
    if (searchResults.length > 0) {
      match = searchResults[0]; // Pick the best keyword match
    }
  }

  if (!match) {
    return `No official ECMA-376 definitions found locally in the RAG database for tag <${context.namespace}:${cleanTagName}>.`;
  }

  // 4. Return enriched grounding context with official citations and SDK mappings
  return `
[ECMA-376 Specification Context]
- Tag Name: <${match.namespace}:${match.tag}>
- Schema Domain: ${match.domain.toUpperCase()}
- Definition: ${match.definition}
- Official Citation: ${match.citation || 'ECMA-376 Standard'}
- Microsoft Open XML SDK Class: DocumentFormat.OpenXml.${match.domain === 'docx' ? 'Wordprocessing' : match.domain === 'xlsx' ? 'Spreadsheet' : 'Presentation'}.${match.sdkClass || match.tag}
- Supported Attributes: ${match.attributes.join(', ') || 'None'}
- Valid Parent Elements: ${match.parents.join(', ') || 'None'}
`;
};

/**
 * Logs user feedback or corrections for a specific tag to localStorage.
 * Used for developer audits and to patch the knowledge base.
 */
export const logRagFeedback = (
  tagName: string,
  fileType: 'docx' | 'xlsx' | 'pptx',
  feedback: string
): void => {
  if (typeof window !== 'undefined' && typeof localStorage !== 'undefined') {
    const feedbackKey = `ooxml_rag_feedback_${fileType}_${tagName.trim()}`;
    localStorage.setItem(feedbackKey, feedback.trim());
    console.log(`[RAG Feedback] Logged feedback for <${tagName}>: ${feedback}`);
  }
};
