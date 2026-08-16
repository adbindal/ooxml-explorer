import { KNOWLEDGE_BASE, ReferenceDoc } from './staticKnowledgeBase';
import { querySchemaFromStorage, searchSchemasInStorage } from './storageService';

/**
 * Looks a tag up in the bundled fallback knowledge base.
 *
 * IndexedDB is the primary store, but it is empty until `/rag-data.json` has been
 * fetched and can be unavailable outright (private browsing modes, storage pressure,
 * a failed fetch). Without this, the common tags would come back ungrounded in exactly
 * the offline situations the app is meant to handle - and under DLP mode, where no
 * network call is permitted, that is the situation.
 */
const queryStaticKnowledgeBase = (
  tag: string,
  domain: 'docx' | 'xlsx' | 'pptx'
): ReferenceDoc | null =>
  KNOWLEDGE_BASE.find(
    doc => doc.tag === tag && (doc.domain === domain || doc.domain === 'shared')
  ) ?? null;

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

export interface RagContextResult {
  /** The text block to inject into the AI prompt. */
  context: string;
  /**
   * Whether `context` carries an actual citation (official DB match or a human-provided
   * runtime override) versus the "nothing found" fallback. The RAG database only covers
   * a curated subset of ECMA-376 - most tags will not be grounded - so callers MUST use
   * this to avoid instructing the model to cite a source that doesn't exist.
   */
  grounded: boolean;
}

/**
 * RAG Context Retrieval & Router.
 * Queries the client-side IndexedDB database for a tag and filters by the active editor context.
 * Falls back to LLM-assisted keyword search for natural language queries.
 * Supports runtime self-healing overrides via localStorage.
 */
export const getRagContext = async (
  tagName: string,
  context: { fileType: 'docx' | 'xlsx' | 'pptx'; namespace: string }
): Promise<RagContextResult> => {
  const cleanTagName = tagName.trim();

  // 1. Check for runtime self-healing overrides in localStorage first
  if (typeof window !== 'undefined' && typeof localStorage !== 'undefined') {
    const overrideKey = `ooxml_rag_override_${context.fileType}_${cleanTagName}`;
    const localOverride = localStorage.getItem(overrideKey);
    if (localOverride) {
      console.log(`[RAG Router] Applying local self-healing override for <${cleanTagName}>`);
      return {
        grounded: true,
        context: `
[ECMA-376 Specification Context (Self-Healed Override)]
- Tag Name: <${context.namespace}:${cleanTagName}>
- Schema Domain: ${context.fileType.toUpperCase()}
- Definition: ${localOverride}
- Source: Runtime Local Patch
`
      };
    }
  }

  // 2. Try direct lookup in IndexedDB, then the bundled fallback.
  //
  // Storage access is wrapped because IndexedDB can be unavailable rather than merely
  // empty; letting that reject would fail the whole explanation instead of degrading
  // to the offline knowledge base.
  let match: ReferenceDoc | null = null;
  try {
    match = await querySchemaFromStorage(cleanTagName, context.fileType);
  } catch (e) {
    console.warn('[RAG Router] IndexedDB lookup failed; using bundled knowledge base:', e);
  }
  if (!match) {
    match = queryStaticKnowledgeBase(cleanTagName, context.fileType);
  }

  // 3. Fallback: If not found and looks like a natural language query, run LLM keyword search
  if (!match && (cleanTagName.includes(' ') || cleanTagName.length > 15)) {
    console.log(`[RAG Router] Tag <${cleanTagName}> not found. Running LLM-assisted keyword search...`);
    const keywords = await extractKeywordsWithLocalAI(cleanTagName);
    try {
      const searchResults = await searchSchemasInStorage(keywords, context.fileType);
      if (searchResults.length > 0) {
        match = searchResults[0]; // Pick the best keyword match
      }
    } catch (e) {
      console.warn('[RAG Router] IndexedDB keyword search failed:', e);
    }
  }

  if (!match) {
    // The curated RAG database only covers a small subset of ECMA-376. This is the
    // common case, not an edge case - be explicit that there is nothing to cite here,
    // so the prompt (see aiService.ts) can instruct the model not to invent one.
    return {
      grounded: false,
      context: `No official ECMA-376 definitions found locally in the RAG database for tag <${context.namespace}:${cleanTagName}>. No citation or SDK class name is available for this tag.`
    };
  }

  // 4. Return enriched grounding context with official citations and SDK mappings.
  //
  // Most of the database is generated from the Open XML SDK's schema metadata, which
  // gives verified structure (attributes, parents, class name) but no prose. Those
  // records are genuinely grounded about structure and genuinely silent about meaning,
  // and the prompt has to say so - interpolating an absent definition would print
  // "undefined" under a "Grounded" badge, and paraphrasing one would invent authority
  // the schema never conferred.
  const sdkNamespace =
    match.domain === 'docx' ? 'Wordprocessing'
    : match.domain === 'xlsx' ? 'Spreadsheet'
    : match.domain === 'pptx' ? 'Presentation'
    : 'Drawing';

  const definitionLine = match.definition
    ? `- Definition: ${match.definition}`
    : `- Definition: NOT AVAILABLE. The schema database covers this element's structure but carries no description for it. Explain the element from your own knowledge and say plainly that the explanation is not from the specification.`;

  const citationLine = match.citation
    ? `- Official Citation: ${match.citation}`
    : `- Official Citation: none on file. Do not cite a specification section for this element.`;

  return {
    grounded: true,
    context: `
[ECMA-376 Specification Context]
- Tag Name: <${match.namespace}:${match.tag}>
- Schema Domain: ${match.domain.toUpperCase()}
${definitionLine}
${citationLine}
- Microsoft Open XML SDK Class: DocumentFormat.OpenXml.${sdkNamespace}.${match.sdkClass || match.tag}
- Supported Attributes: ${match.attributes.join(', ') || 'None'}
- Valid Parent Elements: ${match.parents.join(', ') || 'None'}
`
  };
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
