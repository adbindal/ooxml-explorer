import MiniSearch from 'minisearch';
import { KNOWLEDGE_BASE, ReferenceDoc } from './ragDb';

// Initialize MiniSearch indexer
const searchIndex = new MiniSearch<ReferenceDoc>({
  fields: ['tag', 'definition'],
  storeFields: ['tag', 'namespace', 'domain', 'definition', 'attributes', 'parents'],
  idField: 'tag' // Tags are unique keys within their respective domains
});

// Load the knowledge base into the indexer
searchIndex.addAll(KNOWLEDGE_BASE);

/**
 * RAG Context Retrieval & Router.
 * Queries the client-side database for a tag and filters by the active editor context (DOCX, XLSX, PPTX).
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

  // 2. Search the local index using MiniSearch
  const results = searchIndex.search(cleanTagName);
  
  // 3. Route the query (filter by active editor domain or 'shared')
  const filtered = results.filter(res => 
    res.domain === context.fileType || res.domain === 'shared'
  ) as unknown as ReferenceDoc[];

  if (filtered.length === 0) {
    return `No official ECMA-376 definitions found locally in the RAG database for tag <${context.namespace}:${cleanTagName}>.`;
  }

  // Pick the best match (MiniSearch ranks by BM25 score)
  const bestMatch = filtered[0];
  
  return `
[ECMA-376 Specification Context]
- Tag Name: <${bestMatch.namespace}:${bestMatch.tag}>
- Schema Domain: ${bestMatch.domain.toUpperCase()}
- Definition: ${bestMatch.definition}
- Supported Attributes: ${bestMatch.attributes.join(', ') || 'None'}
- Valid Parent Elements: ${bestMatch.parents.join(', ') || 'None'}
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
