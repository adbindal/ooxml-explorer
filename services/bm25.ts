/**
 * BM25 ranking for the natural-language fallback.
 *
 * The primary retrieval path is an exact `(tag, namespace, domain)` lookup — a
 * dictionary key, not a search problem. This module serves the tail: someone types
 * *"how do I stop a table row splitting across pages"* instead of clicking `w:cantSplit`.
 *
 * WHAT THIS REPLACES.
 *
 * `searchSchemasInStorage` returns every record whose tag or definition *contains* the
 * keyword, and the router took `searchResults[0]` — **the first substring hit, in
 * whatever order IndexedDB happened to yield it, with no ranking at all**. A search for
 * "table" returned whichever of the hundred table-related records the cursor reached
 * first. That is not retrieval; it is a coin toss with extra steps.
 *
 * WHY BM25 AND NOT EMBEDDINGS.
 *
 * This is not the decision the retrieval counters were added to settle. That question is
 * **embeddings versus lexical**, and it stays open until there is usage data, because
 * embeddings cost a model, memory, and a learned component sitting near the trust path.
 * BM25 needs none of those: it is arithmetic over 1,899 records, it runs in the browser
 * in milliseconds, and it beats "first substring hit" under every hypothesis. Choosing it
 * now does not pre-empt the measurement — it raises the floor the measurement compares
 * against.
 *
 * Lexical scoring also happens to suit this corpus unusually well. Retrieval research is
 * consistent that BM25 beats dense retrieval on rare structured identifiers, and this
 * corpus is *made of* them: `kinsoku`, `bidiVisual`, `oleObject`. An embedding of
 * `w:cantSplit` is a guess about a token the model may never have seen; a term match is
 * not.
 *
 * PARAMETERS. `k1 = 1.2` and `b = 0.75` are the standard defaults from the TREC work
 * BM25 comes out of, kept deliberately rather than tuned: tuning them against no
 * relevance judgements would be fitting noise, and the honest position is a documented
 * default until there is data to beat it.
 */

import type { ReferenceDoc } from './staticKnowledgeBase';

/** Term-frequency saturation. Higher means repeated terms keep adding weight. */
export const K1 = 1.2;
/** Length normalisation. 0 ignores document length; 1 divides fully by it. */
export const B = 0.75;

/**
 * Splits identifiers the way this corpus is actually written.
 *
 * OOXML names are camelCase and prefixed, so `w:bookmarkStart` has to yield `bookmark`
 * and `start` as well as `bookmarkstart` — a user typing "bookmark" must reach it, and
 * one typing the whole tag must reach it more strongly. Emitting both the split parts
 * and the whole token is what gives that ordering for free: the exact tag matches on
 * every term, a partial word matches on one.
 */
export const tokenise = (text: string): string[] => {
  const tokens: string[] = [];
  // Split on anything that is not a letter or digit; the namespace colon, hyphens,
  // slashes and punctuation are all separators here.
  for (const chunk of text.split(/[^A-Za-z0-9]+/)) {
    if (!chunk) continue;
    const lower = chunk.toLowerCase();
    tokens.push(lower);
    // camelCase / PascalCase boundaries, plus letter-digit transitions (ST_OnOff1).
    const parts = chunk.split(/(?<=[a-z0-9])(?=[A-Z])|(?<=[A-Za-z])(?=[0-9])/);
    if (parts.length > 1) {
      for (const part of parts) tokens.push(part.toLowerCase());
    }
  }
  return tokens;
};

/**
 * The searchable text of a record.
 *
 * The tag is included twice, which is the one piece of weighting here: a term in the
 * element's own name is better evidence than the same term somewhere in its prose.
 *
 * ⚠️ It is a MILD preference, not a decisive one — worth stating because the obvious
 * reading of "weighted" is stronger than what this does. Doubling gives the tag the
 * pull of two mentions, so a definition repeating a term three times still outranks it.
 * That is acceptable here because definitions are a sentence long and do not repeat
 * terms; if that ever stops being true, this needs a real per-field weighting rather
 * than a thumb on the scale, and inventing one now would be tuning against nothing.
 */
const documentText = (doc: ReferenceDoc): string =>
  [doc.tag, doc.tag, doc.definition ?? '', doc.attributes.join(' '), doc.sdkClass ?? ''].join(' ');

interface IndexedDoc {
  doc: ReferenceDoc;
  /** term → count within this document. */
  frequencies: Map<string, number>;
  length: number;
}

export interface Bm25Index {
  documents: IndexedDoc[];
  /** term → how many documents contain it. */
  documentFrequency: Map<string, number>;
  averageLength: number;
}

export const buildIndex = (docs: readonly ReferenceDoc[]): Bm25Index => {
  const documents: IndexedDoc[] = [];
  const documentFrequency = new Map<string, number>();
  let totalLength = 0;

  for (const doc of docs) {
    const terms = tokenise(documentText(doc));
    const frequencies = new Map<string, number>();
    for (const term of terms) frequencies.set(term, (frequencies.get(term) ?? 0) + 1);
    for (const term of frequencies.keys()) {
      documentFrequency.set(term, (documentFrequency.get(term) ?? 0) + 1);
    }
    documents.push({ doc, frequencies, length: terms.length });
    totalLength += terms.length;
  }

  return {
    documents,
    documentFrequency,
    averageLength: documents.length > 0 ? totalLength / documents.length : 0
  };
};

/**
 * Inverse document frequency, in the form that cannot go negative.
 *
 * The textbook BM25 IDF turns negative for a term appearing in more than half the
 * corpus, which would make a common term actively *reduce* a document's score — so a
 * record matching both a common and a rare term could rank below one matching only the
 * rare one. The `+1` variant used here is the standard fix and keeps every contribution
 * non-negative.
 */
const idf = (index: Bm25Index, term: string): number => {
  const n = index.documentFrequency.get(term) ?? 0;
  return Math.log(1 + (index.documents.length - n + 0.5) / (n + 0.5));
};

export interface ScoredDoc {
  doc: ReferenceDoc;
  score: number;
  /** Query terms this document actually matched, for explaining a result. */
  matched: string[];
}

/**
 * Ranks the corpus against a query, best first.
 *
 * Documents matching nothing are dropped rather than returned with a zero score: a
 * result that matched no term is not a weak answer, it is not an answer, and passing it
 * on is exactly how the previous implementation produced confident nonsense.
 */
export const search = (index: Bm25Index, query: string, limit = 10): ScoredDoc[] => {
  const terms = [...new Set(tokenise(query))];
  if (terms.length === 0) return [];

  const scored: ScoredDoc[] = [];
  for (const entry of index.documents) {
    let score = 0;
    const matched: string[] = [];
    for (const term of terms) {
      const frequency = entry.frequencies.get(term);
      if (!frequency) continue;
      matched.push(term);
      // No zero-guard needed, and adding one would imply a case that cannot occur:
      // this line is only reached when a document matched a term, which means it has
      // at least one token, which means averageLength is above zero.
      const normalised = entry.length / index.averageLength;
      score += idf(index, term) * ((frequency * (K1 + 1)) / (frequency + K1 * (1 - B + B * normalised)));
    }
    if (matched.length > 0) scored.push({ doc: entry.doc, score, matched });
  }

  // Ties broken by tag so the order is stable between runs; an unstable ranking makes
  // two identical questions produce two different citations.
  scored.sort((a, b) => b.score - a.score || a.doc.tag.localeCompare(b.doc.tag));
  return scored.slice(0, limit);
};

/**
 * The single best match, or null when nothing is convincing.
 *
 * `minScore` exists because BM25 always ranks *something* first if any term matched at
 * all. A query sharing one incidental word with a record is not grounding, and the
 * evidence tier depends on this refusing to answer rather than answering weakly — the
 * same refuse-to-guess rule the element locators follow.
 */
export const bestMatch = (index: Bm25Index, query: string, minScore = 1): ReferenceDoc | null => {
  const [top] = search(index, query, 1);
  return top && top.score >= minScore ? top.doc : null;
};
