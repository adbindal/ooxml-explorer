import { describe, it, expect } from 'vitest';
import { selectEvidenceTier, type ComputedEvidence } from '../services/aiService';

/**
 * The badge must be computed from evidence provenance, never asserted by the model.
 *
 * Benchmark work (ALCE, EMNLP 2023) found that on long-form open-ended questions even
 * the best models fail to fully support their own citations about half the time, so a
 * model-reported confidence would be wrong roughly that often in exactly the register
 * this app writes in. These tests pin the rule that replaces it.
 */

const computed = (lines: string[], unresolved: string[] = []): ComputedEvidence =>
  ({ lines, unresolved });

describe('with nothing backing the answer', () => {
  it('is unverified', () => {
    expect(selectEvidenceTier(false)).toBe('unverified');
    expect(selectEvidenceTier(false, null)).toBe('unverified');
  });

  it('treats empty computed evidence as no evidence', () => {
    expect(selectEvidenceTier(false, computed([]))).toBe('unverified');
  });
});

describe('with one source', () => {
  it('is grounded on a citation alone', () => {
    expect(selectEvidenceTier(true)).toBe('grounded');
  });

  it('is verified on a complete computation alone', () => {
    expect(selectEvidenceTier(false, computed(['sz = 32 (from style:Heading1)']))).toBe('verified');
  });

  it('caps a computation with gaps at grounded', () => {
    // A computation that could not establish everything still beats recall, but it
    // is not entitled to call itself verified.
    expect(selectEvidenceTier(false, computed(['sz = 32'], ['numbering part is missing'])))
      .toBe('grounded');
  });
});

describe('with several sources, the computation governs', () => {
  it('keeps a complete computation at verified when a citation also matched', () => {
    // This is the case the rule used to get wrong. A dictionary hit is a fact about
    // the FORMAT; the computation is a claim about THIS FILE. Corroboration from the
    // first says nothing about the certainty of the second.
    expect(selectEvidenceTier(true, computed(['ind = 720 (from numbering:1/0)']))).toBe('verified');
  });

  it('reports the same tier whether the dictionary hit or missed', () => {
    // The decisive property. Under the old minimum rule a HIT scored lower than a
    // MISS, so finding supporting evidence made the answer rank worse. Knowing more
    // must never reduce confidence.
    const evidence = computed(['b = applied (from style:Heading1)']);
    expect(selectEvidenceTier(true, evidence)).toBe(selectEvidenceTier(false, evidence));
  });

  it('still refuses to lift a gapped computation to verified', () => {
    // The cap that must survive the change: a citation about the format cannot fill a
    // hole in what was resolved from the file.
    expect(selectEvidenceTier(true, computed(['x'], ['numbering part is missing'])))
      .toBe('grounded');
  });
});

describe('absent evidence is not a weak source', () => {
  it('does not let a RAG miss drag down a verified computation', () => {
    // The common case: the tag is not in the curated corpus, but the cascade
    // resolved completely. That answer is verified, not unverified.
    expect(selectEvidenceTier(false, computed(['b = applied (from style:Heading1)']))).toBe('verified');
  });
});
