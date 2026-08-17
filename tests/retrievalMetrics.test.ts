import { describe, it, expect, beforeEach } from 'vitest';
import {
  readRetrievalMetrics,
  recordRetrieval,
  resetRetrievalMetrics,
  summariseRetrieval
} from '../services/retrievalMetrics';

/**
 * These counters exist to settle one question: how often does the natural-language
 * fallback fire? It is the only measurement that would justify semantic search here,
 * and it has been unanswerable since the retrieval design was first reviewed.
 */

describe('counting', () => {
  beforeEach(() => resetRetrievalMetrics());

  it('starts at zero for every path', () => {
    expect(readRetrievalMetrics()).toEqual({
      exactHit: 0, offlineHit: 0, naturalLanguageAttempts: 0,
      naturalLanguageHits: 0, overrideHit: 0, miss: 0
    });
  });

  it('accumulates', () => {
    recordRetrieval('exactHit');
    recordRetrieval('exactHit');
    recordRetrieval('miss');
    const m = readRetrievalMetrics();
    expect(m.exactHit).toBe(2);
    expect(m.miss).toBe(1);
  });

  it('survives corrupt stored data rather than throwing', () => {
    localStorage.setItem('ooxml_retrieval_metrics', 'not json');
    expect(readRetrievalMetrics().exactHit).toBe(0);
  });

  it('defaults a counter added later to zero rather than undefined', () => {
    localStorage.setItem('ooxml_retrieval_metrics', JSON.stringify({ exactHit: 5 }));
    expect(readRetrievalMetrics().miss).toBe(0);
    expect(readRetrievalMetrics().exactHit).toBe(5);
  });
});

describe('what the numbers are for', () => {
  beforeEach(() => resetRetrievalMetrics());

  it('says nothing when there is no data, rather than guessing', () => {
    expect(summariseRetrieval(readRetrievalMetrics()).lines[0]).toContain('No lookups recorded');
  });

  it('reports semantic search as unjustified when the NL path is rare', () => {
    for (let i = 0; i < 99; i++) recordRetrieval('exactHit');
    recordRetrieval('naturalLanguageAttempts');
    const s = summariseRetrieval(readRetrievalMetrics());
    expect(s.total).toBe(100);
    expect(s.naturalLanguageShare).toBeCloseTo(0.01);
    expect(s.lines.join(' ')).toContain('not justified by this data');
  });

  it('recommends BM25 before embeddings when the NL path is common', () => {
    for (let i = 0; i < 5; i++) recordRetrieval('exactHit');
    for (let i = 0; i < 5; i++) recordRetrieval('naturalLanguageAttempts');
    const s = summariseRetrieval(readRetrievalMetrics());
    expect(s.naturalLanguageShare).toBeCloseTo(0.5);
    expect(s.lines.join(' ')).toContain('BM25 first');
  });

  it('counts an NL attempt once, whether or not it found anything', () => {
    // Attempts drive the decision; hits describe how well the current scan works.
    recordRetrieval('naturalLanguageAttempts');
    recordRetrieval('naturalLanguageHits');
    expect(summariseRetrieval(readRetrievalMetrics()).total).toBe(1);
  });
});

describe('privacy', () => {
  it('stores only counts, never query text', () => {
    // The app has a DLP mode whose promise is that nothing leaves the device. A log
    // of what someone searched is exactly what that promise is about.
    resetRetrievalMetrics();
    recordRetrieval('naturalLanguageAttempts');
    const stored = localStorage.getItem('ooxml_retrieval_metrics') ?? '';
    expect(stored).toMatch(/^\{("[a-zA-Z]+":\d+,?)+\}$/);
  });
});
