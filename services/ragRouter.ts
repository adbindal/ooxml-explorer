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
 */
export const getRagContext = async (
  tagName: string,
  context: { fileType: 'docx' | 'xlsx' | 'pptx'; namespace: string }
): Promise<string> => {
  const cleanTagName = tagName.trim();
  
  // 1. Search the local index using MiniSearch
  const results = searchIndex.search(cleanTagName);
  
  // 2. Route the query (filter by active editor domain or 'shared')
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
