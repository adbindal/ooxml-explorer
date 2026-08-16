import { describe, it, expect } from '../services/browserTestRunner';
import { explainElement } from '../services/aiService';
import { analyzeFile, analyzeDiff } from '../services/geminiService';
import { useAppStore } from '../store/appStore';

/**
 * Real Local AI Model Eval (Gemini Nano) - opt-in, no mocks.
 *
 * Every other AI test in this suite mocks window.LanguageModel, which proves the
 * app's plumbing is correct but can never catch a real model behaving differently
 * than the mock assumes - which is exactly how the "local AI echoes the JSON schema
 * back instead of filling it in" bug slipped past every automated test.
 *
 * This file intentionally does NOT mock window.LanguageModel. It feature-detects the
 * real Prompt API and:
 *   - In CI / `vitest run` / Playwright's isolated automation profile: the API is
 *     absent, so these tests no-op with a clear skip message. This is expected and
 *     is not a failure - it's the same "not available in this environment" pattern
 *     used throughout the app's own local-AI availability checks.
 *   - In a developer's real Chrome, with the on-device model flags enabled and the
 *     model downloaded (see services/README.md), opening the in-browser Validator
 *     and running the test suite there executes these against the REAL model. That
 *     run is the actual signal for "is local AI usable," not the mocked tests.
 */

const isRealLocalAiAvailable = async (): Promise<boolean> => {
  if (typeof window === 'undefined' || !window.LanguageModel) return false;
  try {
    return (await window.LanguageModel.availability()) === 'available';
  } catch {
    return false;
  }
};

const skip = (label: string) => {
  console.warn(`[Real Model Eval] Skipping "${label}": window.LanguageModel is not available in this environment.`);
  expect(true).toBe(true);
};

describe('Real Local AI Model Eval (Gemini Nano, opt-in - see file header)', () => {
  it('produces an accurate, grounded explanation for a known tag via the real local model', async () => {
    if (!(await isRealLocalAiAvailable())) return skip('grounded tag explanation');

    useAppStore.setState(state => ({ ui: { ...state.ui, aiProvider: 'chrome-local', dlpMode: true } }));

    const result = await explainElement(
      'cantSplit',
      '<w:trPr><w:cantSplit/></w:trPr>',
      'docx'
    );

    // Grounded in the real RAG data (table-row splitting), not a hallucinated
    // guess like "text boxes" - this is the behavior the RAG layer exists to
    // guarantee, and only a real model run can verify it.
    expect(result.grounded).toBe(true);
    const lower = result.explanation.toLowerCase();
    expect(lower).toContain('row');
    expect(lower.includes('page') || lower.includes('split')).toBe(true);
  });

  it('produces a schema-valid structured file analysis via the real local model', async () => {
    if (!(await isRealLocalAiAvailable())) return skip('structured file analysis');

    useAppStore.setState(state => ({ ui: { ...state.ui, aiProvider: 'chrome-local', dlpMode: true } }));

    const result = await analyzeFile(
      [{ fileName: 'word/document.xml', content: '<w:document><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>' }],
      'explain'
    );

    // The exact assertion the schema-echo bug would have failed: a real Zod-valid
    // instance, not the schema definition echoed back.
    expect(typeof result.summary).toBe('string');
    expect(result.summary.length > 0).toBe(true);
    expect(Array.isArray(result.criticalIssues)).toBe(true);
    expect(Array.isArray(result.keyElements)).toBe(true);
  });

  it('produces a schema-valid structured diff analysis via the real local model', async () => {
    if (!(await isRealLocalAiAvailable())) return skip('structured diff analysis');

    useAppStore.setState(state => ({ ui: { ...state.ui, aiProvider: 'chrome-local', dlpMode: true } }));

    const result = await analyzeDiff(
      [{
        fileName: 'word/document.xml',
        original: '<w:document><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>',
        modified: '<w:document><w:body><w:p><w:r><w:t>Hello World</w:t></w:r></w:p></w:body></w:document>'
      }],
      'summary'
    );

    expect(typeof result.summary).toBe('string');
    expect(result.summary.length > 0).toBe(true);
    expect(Array.isArray(result.changesList)).toBe(true);
  });
});
