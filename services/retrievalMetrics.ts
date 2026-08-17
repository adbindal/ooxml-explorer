/**
 * Counts which retrieval path answered each lookup.
 *
 * This exists to settle a question that has been open since the retrieval design was
 * first reviewed and that nothing could answer: **how often does the natural-language
 * fallback actually fire?**
 *
 * It matters because it is the only thing that would justify semantic search here. The
 * primary path is an exact lookup of a tag the user clicked — a dictionary key, not a
 * similarity problem — and lexical retrieval beats dense retrieval decisively on rare
 * structured identifiers like `w:kinsoku`. Embeddings would only ever serve the
 * paraphrase tail. If that tail is 2% of lookups, better keyword matching is enough; if
 * it is 40%, the calculation changes. Guessing either way is how projects acquire a
 * vector database they did not need.
 *
 * **Counts only, never query text.** The app has a DLP mode whose whole promise is that
 * nothing leaves the device, and a log of what someone searched for is exactly the kind
 * of thing that promise is about. Six integers in `localStorage` answer the question
 * without recording anything about the user.
 */

const STORAGE_KEY = 'ooxml_retrieval_metrics';

export interface RetrievalMetrics {
  /** Exact `(tag, namespace, domain)` hit in IndexedDB. The path that should dominate. */
  exactHit: number;
  /** Exact hit served from the bundled offline knowledge base instead. */
  offlineHit: number;
  /** The natural-language fallback was entered. */
  naturalLanguageAttempts: number;
  /** …and returned something. */
  naturalLanguageHits: number;
  /** A runtime self-healing override answered it. */
  overrideHit: number;
  /** Nothing matched; the answer went out ungrounded. */
  miss: number;
}

const EMPTY: RetrievalMetrics = {
  exactHit: 0,
  offlineHit: 0,
  naturalLanguageAttempts: 0,
  naturalLanguageHits: 0,
  overrideHit: 0,
  miss: 0
};

const available = (): boolean =>
  typeof window !== 'undefined' && typeof localStorage !== 'undefined';

export const readRetrievalMetrics = (): RetrievalMetrics => {
  if (!available()) return { ...EMPTY };
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { ...EMPTY };
    const parsed = JSON.parse(raw) as Partial<RetrievalMetrics>;
    // Merge over EMPTY so a counter added later reads as 0 rather than undefined.
    return { ...EMPTY, ...parsed };
  } catch {
    return { ...EMPTY };
  }
};

export const recordRetrieval = (event: keyof RetrievalMetrics): void => {
  if (!available()) return;
  try {
    const metrics = readRetrievalMetrics();
    metrics[event] += 1;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(metrics));
  } catch {
    // Metrics are diagnostic. Never let a storage failure break a lookup.
  }
};

export const resetRetrievalMetrics = (): void => {
  if (!available()) return;
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch {
    // ignore
  }
};

/**
 * Summarises the counts, including the figure the decision actually turns on.
 */
export const summariseRetrieval = (metrics: RetrievalMetrics): {
  total: number;
  naturalLanguageShare: number;
  lines: string[];
} => {
  const total =
    metrics.exactHit + metrics.offlineHit + metrics.overrideHit +
    metrics.naturalLanguageAttempts + metrics.miss;

  const share = total === 0 ? 0 : metrics.naturalLanguageAttempts / total;

  const pct = (n: number) => (total === 0 ? '0.0' : ((n / total) * 100).toFixed(1));

  return {
    total,
    naturalLanguageShare: share,
    lines: total === 0
      ? ['No lookups recorded yet.']
      : [
          `${total} lookups recorded.`,
          `  exact hit (IndexedDB):    ${metrics.exactHit} (${pct(metrics.exactHit)}%)`,
          `  exact hit (offline KB):   ${metrics.offlineHit} (${pct(metrics.offlineHit)}%)`,
          `  self-healing override:    ${metrics.overrideHit} (${pct(metrics.overrideHit)}%)`,
          `  natural-language attempt: ${metrics.naturalLanguageAttempts} (${pct(metrics.naturalLanguageAttempts)}%), of which ${metrics.naturalLanguageHits} found something`,
          `  no match:                 ${metrics.miss} (${pct(metrics.miss)}%)`,
          '',
          share < 0.05
            ? 'The natural-language path is rare. Better keyword matching would be sufficient; semantic search is not justified by this data.'
            : 'The natural-language path is used enough to be worth improving. BM25 first - it is pure JS and works offline. Reach for embeddings only if BM25 measurably falls short.'
        ]
  };
};
