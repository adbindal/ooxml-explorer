import { describe, it, expect } from 'vitest';
import { tokenise, buildIndex, search, bestMatch, K1, B } from '../services/bm25';
import type { ReferenceDoc } from '../services/staticKnowledgeBase';

const doc = (tag: string, definition?: string, attributes: string[] = []): ReferenceDoc => ({
  tag,
  namespace: 'w',
  domain: 'docx',
  definition,
  attributes,
  parents: []
});

const corpus: ReferenceDoc[] = [
  doc('cantSplit', 'Prevents a table row from breaking across pages.', ['val']),
  doc('tblHeader', 'Repeats a table row as a header at the top of each page.', ['val']),
  doc('bookmarkStart', 'Marks the beginning of a bookmark range.', ['id', 'name']),
  doc('bookmarkEnd', 'Marks the end of a bookmark range.', ['id']),
  doc('tbl', 'A table.', []),
  doc('p', 'A paragraph.', []),
  doc('kinsoku', 'Applies East Asian line breaking rules.', ['val'])
];

const index = buildIndex(corpus);

describe('tokenise', () => {
  it('splits camelCase so a partial word still reaches the tag', () => {
    // Someone typing "bookmark" must reach w:bookmarkStart. Without the split, the only
    // token is "bookmarkstart" and the query misses entirely.
    expect(tokenise('bookmarkStart')).toContain('bookmark');
    expect(tokenise('bookmarkStart')).toContain('start');
  });

  it('keeps the whole token as well as the parts, so an exact tag outranks a partial', () => {
    expect(tokenise('bookmarkStart')).toContain('bookmarkstart');
  });

  it('treats the namespace colon as a separator', () => {
    expect(tokenise('w:cantSplit')).toEqual(expect.arrayContaining(['w', 'cantsplit', 'cant', 'split']));
  });

  it('splits letter-digit boundaries, which OOXML type names use', () => {
    expect(tokenise('ST_OnOff1')).toEqual(expect.arrayContaining(['st', 'on', 'off', '1']));
  });

  it('returns nothing for punctuation alone', () => {
    expect(tokenise('   ?? -- ')).toEqual([]);
  });
});

describe('ranking', () => {
  it('ranks the right record first for a natural-language question', () => {
    // The case the fallback exists for: nobody types w:cantSplit.
    const [top] = search(index, 'stop a table row splitting across pages');

    expect(top.doc.tag).toBe('cantSplit');
  });

  it('beats first-substring-hit on an ambiguous term', () => {
    // "table" substring-matches four records here. The old router took whichever the
    // cursor yielded first; ranking has to prefer the one whose own NAME is the match.
    const [top] = search(index, 'table');

    expect(top.doc.tag).toBe('tbl');
  });

  it('prefers a term in the tag over the same term in prose', () => {
    // Isolates the tag doubling: both records mention "split" once in prose, but only
    // one has it in its NAME. An earlier version of this test had the prose repeat the
    // term three times and FAILED - correctly, because the doubling is a mild
    // preference, not a decisive one. See the note on documentText.
    // The competitor is named so it sorts BEFORE splitCell: without the doubling both
    // records score identically and the alphabetical tie-break would hand the win to
    // splitCell anyway, so the test would pass without testing anything.
    const tagged = doc('splitCell', 'Divides one cell into two.');
    const wordy = doc('mergeCell', 'Continues a split region downward.');
    const [top] = search(buildIndex([tagged, wordy]), 'split');

    expect(top.doc.tag).toBe('splitCell');
  });

  it('drops records that matched nothing rather than scoring them zero', () => {
    // A record that matched no term is not a weak answer, it is not an answer.
    const results = search(index, 'kinsoku');

    expect(results).toHaveLength(1);
    expect(results[0].doc.tag).toBe('kinsoku');
  });

  it('returns nothing for a query with no usable terms', () => {
    expect(search(index, '   ')).toEqual([]);
  });

  it('reports which terms matched, so a result can be explained', () => {
    const [top] = search(index, 'bookmark range');

    expect(top.matched).toEqual(expect.arrayContaining(['bookmark', 'range']));
  });

  it('ranks equal-scoring records the same whatever order the corpus arrived in', () => {
    // The real property. Re-running the same search twice passes even without a
    // tie-break, because Array.sort is stable - so it proves nothing. Two records with
    // identical text score identically, and IndexedDB makes no ordering promise.
    const a = doc('alpha', 'Marks the end of a bookmark range.');
    const b = doc('beta', 'Marks the end of a bookmark range.');

    const forwards = search(buildIndex([a, b]), 'bookmark range').map(r => r.doc.tag);
    const backwards = search(buildIndex([b, a]), 'bookmark range').map(r => r.doc.tag);

    expect(forwards).toEqual(backwards);
  });

  it('never lets a common term reduce a score', () => {
    // Textbook BM25 IDF goes negative for a term appearing in more than half the
    // corpus, so matching a common term would SUBTRACT. The record has to contain both
    // terms for that to be observable - an earlier version of this test used a record
    // that did not contain the common one, so the score never moved either way.
    // 'a' appears in six of these seven definitions; 'cantsplit' in one.
    const rareOnly = search(index, 'cantSplit').find(r => r.doc.tag === 'cantSplit')!.score;
    const rarePlusCommon = search(index, 'cantSplit a').find(r => r.doc.tag === 'cantSplit')!.score;

    expect(rarePlusCommon).toBeGreaterThanOrEqual(rareOnly);
  });

  it('honours the limit', () => {
    expect(search(index, 'table row page', 2)).toHaveLength(2);
  });
});

describe('bestMatch refuses to guess', () => {
  it('returns null when nothing clears the floor', () => {
    // The evidence tier depends on this. One incidental shared word is not grounding,
    // and citing it under a Grounded badge is the failure the tier exists to prevent.
    expect(bestMatch(index, 'zzzz nonexistent query', 1)).toBeNull();
  });

  it('returns the record when the match is convincing', () => {
    expect(bestMatch(index, 'cantSplit')?.tag).toBe('cantSplit');
  });

  it('applies the floor rather than always taking the top hit', () => {
    // With an impossible floor even a real match is refused, which proves the
    // threshold is consulted at all.
    expect(bestMatch(index, 'cantSplit', 1e6)).toBeNull();
  });
});

describe('the index', () => {
  it('handles an empty corpus without dividing by zero', () => {
    const empty = buildIndex([]);

    expect(empty.averageLength).toBe(0);
    expect(search(empty, 'anything')).toEqual([]);
  });

  it('survives a corpus whose documents have no tokens at all', () => {
    // Robustness only. This does NOT exercise a divide-by-zero: scoring is reached
    // only by a document that matched a term, which therefore has at least one token,
    // so averageLength is above zero whenever the division happens. An earlier version
    // of this test claimed to cover a guard that was provably unreachable, and the
    // guard has since been removed rather than left as false reassurance.
    const blank = buildIndex([doc(''), doc('')]);

    expect(blank.averageLength).toBe(0);
    expect(() => search(blank, 'anything')).not.toThrow();
    expect(search(blank, 'anything')).toEqual([]);
  });

  it('counts document frequency per document, not per occurrence', () => {
    // 'bookmark' appears twice inside bookmarkStart's text but in two documents.
    expect(buildIndex(corpus).documentFrequency.get('bookmark')).toBe(2);
  });

  it('keeps the standard BM25 parameters, documented rather than tuned', () => {
    // Tuning these against no relevance judgements would be fitting noise. Pinned so a
    // future change is a deliberate one.
    expect(K1).toBe(1.2);
    expect(B).toBe(0.75);
  });
});
