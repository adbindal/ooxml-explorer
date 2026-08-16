import { describe, it, expect } from 'vitest';
import {
  allocateContentBudget,
  getLocalContentBudgetChars,
  renderContentSnippet,
  CLOUD_CONTENT_BUDGET_CHARS
} from '../services/promptBudget';

describe('allocateContentBudget', () => {
  it('never hands out more than the budget in total', () => {
    const sizes = [50_000, 40_000, 30_000];
    const limits = allocateContentBudget(sizes, 9_000);
    expect(limits.reduce((a, b) => a + b, 0)).toBeLessThanOrEqual(9_000);
  });

  it('is the fix for the original bug: N files do not multiply the cap', () => {
    // The previous code sliced each file to 8000 chars independently, so ten large
    // files produced an 80,000 character prompt against a window that fits ~9,000.
    const tenLargeFiles = new Array(10).fill(200_000);
    const limits = allocateContentBudget(tenLargeFiles, 9_000);
    expect(limits.reduce((a, b) => a + b, 0)).toBeLessThanOrEqual(9_000);
  });

  it('leaves content untouched when everything fits', () => {
    expect(allocateContentBudget([100, 200, 300], 10_000)).toEqual([100, 200, 300]);
  });

  it('lets small files keep everything and donates the remainder to large ones', () => {
    // An equal split would cut the 500-char file to 4,500 for no reason.
    const [small, large] = allocateContentBudget([500, 40_000], 9_000);
    expect(small).toBe(500);
    expect(large).toBe(8_500);
  });

  it('shares evenly when every file exceeds its share', () => {
    const limits = allocateContentBudget([50_000, 50_000], 9_000);
    expect(limits).toEqual([4_500, 4_500]);
  });

  it('preserves input order regardless of size ordering', () => {
    // Allocation sorts internally; the returned limits must still line up positionally.
    const limits = allocateContentBudget([40_000, 500], 9_000);
    expect(limits).toEqual([8_500, 500]);
  });

  it('handles a zero budget without producing negative limits', () => {
    expect(allocateContentBudget([1_000, 2_000], 0)).toEqual([0, 0]);
  });

  it('handles an empty file list', () => {
    expect(allocateContentBudget([], 9_000)).toEqual([]);
  });

  it('treats an empty file as needing nothing', () => {
    const [empty, rest] = allocateContentBudget([0, 40_000], 9_000);
    expect(empty).toBe(0);
    expect(rest).toBe(9_000);
  });
});

describe('getLocalContentBudgetChars', () => {
  it('reads the current Prompt API property names', () => {
    const budget = getLocalContentBudgetChars({ contextWindow: 6_144, contextUsage: 144 });
    // (6144 - 144 - 768 reserve) * 3 chars per token
    expect(budget).toBe((6_144 - 144 - 768) * 3);
  });

  it('still reads the pre-rename property names for older Chrome builds', () => {
    const renamed = getLocalContentBudgetChars({ contextWindow: 6_144, contextUsage: 144 });
    const legacy = getLocalContentBudgetChars({ inputQuota: 6_144, inputUsage: 144 });
    expect(legacy).toBe(renamed);
  });

  it('prefers the new names when a session reports both', () => {
    const budget = getLocalContentBudgetChars({
      contextWindow: 6_144,
      contextUsage: 144,
      inputQuota: 999_999,
      inputUsage: 0
    });
    expect(budget).toBe((6_144 - 144 - 768) * 3);
  });

  it('falls back conservatively when the session reports no window', () => {
    // Mocked sessions in tests and older Chrome builds report nothing. The fallback
    // must stay well under the smallest real window (~6k tokens) rather than guess high.
    const budget = getLocalContentBudgetChars({});
    expect(budget).toBeGreaterThan(0);
    expect(budget).toBeLessThan(6_144 * 3);
  });

  it('tolerates a null or undefined session', () => {
    expect(getLocalContentBudgetChars(undefined)).toBeGreaterThan(0);
    expect(getLocalContentBudgetChars(null)).toBeGreaterThan(0);
  });

  it('ignores non-numeric values rather than producing NaN', () => {
    const budget = getLocalContentBudgetChars({
      contextWindow: 'lots' as unknown as number,
      contextUsage: null
    });
    expect(Number.isFinite(budget)).toBe(true);
    expect(budget).toBeGreaterThan(0);
  });

  it('returns zero rather than a negative budget when the window is already full', () => {
    expect(getLocalContentBudgetChars({ contextWindow: 100, contextUsage: 100 })).toBe(0);
  });

  it('is far smaller than the cloud budget', () => {
    // The whole point of the change: the two providers get different budgets.
    expect(getLocalContentBudgetChars({ contextWindow: 6_144 })).toBeLessThan(CLOUD_CONTENT_BUDGET_CHARS);
  });
});

describe('renderContentSnippet', () => {
  it('returns content unchanged when it fits', () => {
    expect(renderContentSnippet('<w:p/>', 100)).toBe('<w:p/>');
  });

  it('marks truncation explicitly so the model cannot treat a partial file as complete', () => {
    const rendered = renderContentSnippet('x'.repeat(500), 100);
    expect(rendered).toContain('TRUNCATED');
    expect(rendered).toContain('do not describe it as complete');
  });

  it('keeps exactly the requested number of characters of real content', () => {
    const rendered = renderContentSnippet('x'.repeat(500), 100);
    expect(rendered.startsWith('x'.repeat(100))).toBe(true);
    expect(rendered.startsWith('x'.repeat(101))).toBe(false);
  });

  it('does not truncate content of exactly the limit', () => {
    const exact = 'x'.repeat(100);
    expect(renderContentSnippet(exact, 100)).toBe(exact);
  });

  it('handles a zero limit without throwing', () => {
    const rendered = renderContentSnippet('content', 0);
    expect(rendered).toContain('TRUNCATED');
  });
});
