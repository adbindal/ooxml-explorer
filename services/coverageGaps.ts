/**
 * Records where the engine had nothing to say.
 *
 * This is the "keeps improving" half of the design, and it is deliberately the boring
 * half. The engine does not learn, rewrite its own rules, or grade its own output —
 * every one of those puts a model back on the trust path, which is the reason
 * fine-tuning was ruled out and the reason the Verified badge is computed rather than
 * claimed. What it does instead is **notice what it could not answer and write that
 * down**, producing a backlog ranked by real usage rather than by guesswork.
 *
 * The loop closes through a person: the log says which part or element keeps coming up
 * with no analyzer behind it, and someone decides whether to build one. That is slower
 * than a self-modifying system and it is the only version that keeps the badge honest.
 *
 * Same discipline as `retrievalMetrics`: **counts only, and only over spec vocabulary.**
 * Part paths are normalised (`word/header3.xml` → `word/header#.xml`) and element names
 * are ECMA-376 tag names. Neither is user content, and nothing else is stored — the app
 * has a DLP mode whose promise is that nothing leaves the device, and a log of what
 * someone was working on is exactly what that promise is about.
 */

const STORAGE_KEY = 'ooxml_coverage_gaps';

/** How many distinct keys to keep. Bounded so the log cannot grow without limit. */
const MAX_KEYS = 200;

export interface CoverageGaps {
  /** Normalised part path → times it was opened with no analyzer covering it. */
  parts: Record<string, number>;
  /** Element local name → times it was selected with no analyzer covering it. */
  elements: Record<string, number>;
}

const available = (): boolean =>
  typeof window !== 'undefined' && typeof localStorage !== 'undefined';

/**
 * Collapses the numbering that distinguishes instances of the same kind of part.
 *
 * `word/header1.xml` and `word/header7.xml` are the same gap; keeping them apart would
 * fill the log with one entry per document and hide the pattern that matters.
 */
export const normalisePartPath = (path: string): string => path.replace(/\d+/g, '#');

export const readCoverageGaps = (): CoverageGaps => {
  if (!available()) return { parts: {}, elements: {} };
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { parts: {}, elements: {} };
    const parsed = JSON.parse(raw) as Partial<CoverageGaps>;
    return { parts: { ...parsed.parts }, elements: { ...parsed.elements } };
  } catch {
    return { parts: {}, elements: {} };
  }
};

const write = (gaps: CoverageGaps): void => {
  if (!available()) return;
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(gaps));
  } catch {
    // Storage full or blocked. A metrics write must never break the feature it measures.
  }
};

const bump = (bucket: Record<string, number>, key: string): void => {
  // Once the cap is reached, keep counting keys already tracked but stop adding new
  // ones. Truncating instead would keep discarding the long tail this exists to find.
  if (bucket[key] === undefined && Object.keys(bucket).length >= MAX_KEYS) return;
  bucket[key] = (bucket[key] ?? 0) + 1;
};

/**
 * Records that a part was opened and no analyzer covered it.
 *
 * `elementName` is the tag the user had selected, when there was one — the more
 * specific signal, since a part may be well covered in general while one element in it
 * is not.
 */
export const recordCoverageGap = (partPath: string, elementName?: string): void => {
  const gaps = readCoverageGaps();
  bump(gaps.parts, normalisePartPath(partPath));
  if (elementName) bump(gaps.elements, elementName);
  write(gaps);
};

export const resetCoverageGaps = (): void => {
  if (!available()) return;
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch {
    // Nothing to do; the next read falls back to empty.
  }
};

/**
 * The backlog, most-requested first.
 *
 * Reads as a to-do list on purpose: these are the analyzers that would have earned
 * their keep, ordered by how often someone actually needed them.
 */
export const summariseCoverageGaps = (
  gaps: CoverageGaps = readCoverageGaps()
): { lines: string[]; total: number } => {
  const rank = (bucket: Record<string, number>) =>
    Object.entries(bucket).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));

  const parts = rank(gaps.parts);
  const elements = rank(gaps.elements);
  const total = parts.reduce((n, [, count]) => n + count, 0);

  if (total === 0) {
    return { lines: ['No coverage gaps recorded — every part opened so far had an analyzer behind it.'], total: 0 };
  }

  const lines = [`${total} request(s) landed on markup no analyzer covers.`, 'Parts, most requested first:'];
  for (const [path, count] of parts.slice(0, 10)) lines.push(`  ${count} × ${path}`);
  if (elements.length > 0) {
    lines.push('Elements, most requested first:');
    for (const [name, count] of elements.slice(0, 10)) lines.push(`  ${count} × ${name}`);
  }
  return { lines, total };
};
