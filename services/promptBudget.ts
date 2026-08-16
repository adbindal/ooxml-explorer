/**
 * Prompt content budgeting.
 *
 * The on-device Prompt API and the cloud API have context windows that differ by
 * roughly two orders of magnitude, so a single hard-coded per-file slice cannot serve
 * both. Worse, a per-file limit is unbounded in aggregate: N files means N times the
 * limit, and the on-device window is shared between the prompt and the model's own
 * response. This module computes a *total* budget and divides it across the files.
 */

/**
 * Rough characters-per-token for OOXML content. XML is token-dense - angle brackets,
 * namespace prefixes and attribute names all fragment - so this is deliberately
 * pessimistic. Overestimating tokens truncates early, which is safe; underestimating
 * overflows the window, which is not.
 */
const XML_CHARS_PER_TOKEN = 3;

/**
 * Tokens held back for the model's own JSON response. On-device, the window is shared
 * between input and output, so a prompt that fills it leaves no room to answer.
 */
const LOCAL_OUTPUT_RESERVE_TOKENS = 768;

/**
 * Used when the Prompt API doesn't report a window - older Chrome builds that predate
 * the `contextWindow` property, or a mocked session under test. Real reported values
 * have ranged from about 6k to 9k tokens across Chrome versions and the API docs
 * publish no number, so this stays below the low end rather than guessing high.
 */
const LOCAL_CONTEXT_FALLBACK_TOKENS = 4096;

/**
 * The cloud model's window is around a million tokens, so this cap exists to bound
 * request size and cost rather than to fit a context window.
 */
export const CLOUD_CONTENT_BUDGET_CHARS = 120_000;

/**
 * The subset of the Prompt API session we read for budgeting.
 *
 * The specification renamed these part-way through: `inputQuota` became
 * `contextWindow` and `inputUsage` became `contextUsage`. Both spellings are declared
 * so a session from either vintage of Chrome can be measured.
 */
interface ContextReportingSession {
  contextWindow?: unknown;
  contextUsage?: unknown;
  inputQuota?: unknown;
  inputUsage?: unknown;
}

const asPositiveNumber = (value: unknown): number | undefined =>
  typeof value === 'number' && Number.isFinite(value) && value >= 0 ? value : undefined;

/**
 * Reads a live on-device session's remaining context and converts it to a character
 * budget for file content.
 *
 * Never hard-code the window instead of calling this: it varies by Chrome version and
 * the system prompt has already consumed part of it by the time the session exists.
 */
export const getLocalContentBudgetChars = (session: unknown): number => {
  const reported = (session ?? {}) as ContextReportingSession;

  const windowTokens =
    asPositiveNumber(reported.contextWindow) ??
    asPositiveNumber(reported.inputQuota) ??
    LOCAL_CONTEXT_FALLBACK_TOKENS;

  const usedTokens =
    asPositiveNumber(reported.contextUsage) ??
    asPositiveNumber(reported.inputUsage) ??
    0;

  const freeTokens = windowTokens - usedTokens - LOCAL_OUTPUT_RESERVE_TOKENS;
  return Math.max(0, Math.floor(freeTokens * XML_CHARS_PER_TOKEN));
};

/**
 * Divides `budgetChars` across items of the given sizes, returning a per-item limit.
 *
 * Uses water-filling rather than an equal split so that nothing is truncated
 * unnecessarily: items smaller than an equal share keep all of their content and
 * donate what they don't use to the larger ones. With a 9,000 character budget and
 * files of 500 and 40,000 characters, the small file is kept whole and the large one
 * gets the remaining 8,500 - where an equal split would have needlessly cut the small
 * file to 4,500.
 *
 * The returned limits always sum to at most `budgetChars`.
 */
export const allocateContentBudget = (sizes: number[], budgetChars: number): number[] => {
  const limits = new Array<number>(sizes.length).fill(0);
  let remaining = Math.max(0, Math.floor(budgetChars));

  // Smallest first, so each item can only ever claim its fair share of what is left
  // after the items that needed less than their share have been satisfied.
  const bySizeAscending = sizes
    .map((size, index) => ({ size: Math.max(0, size), index }))
    .sort((a, b) => a.size - b.size);

  let unallocated = bySizeAscending.length;
  for (const { size, index } of bySizeAscending) {
    const fairShare = Math.floor(remaining / unallocated);
    const granted = Math.min(size, fairShare);
    limits[index] = granted;
    remaining -= granted;
    unallocated -= 1;
  }

  return limits;
};

/**
 * Truncates content to `limitChars`, leaving an explicit marker when it does.
 *
 * The marker matters beyond tidiness: without it the model cannot tell a file that
 * genuinely ends there from one that was cut, and would describe a partial file as if
 * it were complete. Since the app distinguishes grounded from unverified claims, a
 * silent truncation would let it assert something it never actually saw.
 */
export const renderContentSnippet = (content: string, limitChars: number): string => {
  if (content.length <= limitChars) {
    return content;
  }
  const kept = content.slice(0, Math.max(0, limitChars));
  return `${kept}\n<!-- TRUNCATED: this is the first ${kept.length} of ${content.length} characters. The rest of the file was not provided - do not describe it as complete. -->`;
};
