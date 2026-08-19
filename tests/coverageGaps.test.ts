import { describe, it, expect, beforeEach } from 'vitest';
import {
  readCoverageGaps,
  recordCoverageGap,
  resetCoverageGaps,
  summariseCoverageGaps,
  normalisePartPath
} from '../services/coverageGaps';

beforeEach(() => resetCoverageGaps());

describe('normalisePartPath', () => {
  it('collapses instance numbering so one gap does not become fifty', () => {
    // header1/header7 are the same gap. Keeping them apart fills the log with one
    // entry per document and hides the pattern the log exists to reveal.
    expect(normalisePartPath('word/header1.xml')).toBe('word/header#.xml');
    expect(normalisePartPath('word/header7.xml')).toBe('word/header#.xml');
    expect(normalisePartPath('ppt/slides/slide12.xml')).toBe('ppt/slides/slide#.xml');
  });

  it('leaves an unnumbered path alone', () => {
    expect(normalisePartPath('word/document.xml')).toBe('word/document.xml');
  });
});

describe('recording', () => {
  it('counts repeats of the same gap', () => {
    recordCoverageGap('word/glossary/document.xml');
    recordCoverageGap('word/glossary/document.xml');

    expect(readCoverageGaps().parts['word/glossary/document.xml']).toBe(2);
  });

  it('merges numbered instances into one count', () => {
    recordCoverageGap('ppt/notesSlides/notesSlide1.xml');
    recordCoverageGap('ppt/notesSlides/notesSlide2.xml');

    expect(readCoverageGaps().parts['ppt/notesSlides/notesSlide#.xml']).toBe(2);
  });

  it('records the selected element separately from the part', () => {
    // A part can be well covered in general while one element in it is not, so the
    // element is the more actionable signal.
    recordCoverageGap('word/document.xml', 'w:smartTag');

    const gaps = readCoverageGaps();
    expect(gaps.elements['w:smartTag']).toBe(1);
    expect(gaps.parts['word/document.xml']).toBe(1);
  });

  it('records the part even when no element was selected', () => {
    recordCoverageGap('customXml/item1.xml');

    expect(readCoverageGaps().elements).toEqual({});
    expect(readCoverageGaps().parts['customXml/item#.xml']).toBe(1);
  });

  it('stores nothing beyond counts of spec vocabulary', () => {
    // The DLP promise: no query text, no document content, no file names beyond the
    // structural part path. Anything else here would be a leak.
    recordCoverageGap('word/document.xml', 'w:smartTag');
    const gaps = readCoverageGaps();

    expect(Object.keys(gaps).sort()).toEqual(['elements', 'parts']);
    for (const bucket of [gaps.parts, gaps.elements]) {
      for (const value of Object.values(bucket)) expect(typeof value).toBe('number');
    }
  });
});

describe('the backlog', () => {
  it('says plainly when nothing is missing', () => {
    const summary = summariseCoverageGaps();

    expect(summary.total).toBe(0);
    expect(summary.lines[0]).toContain('No coverage gaps recorded');
  });

  it('ranks the most-requested gap first', () => {
    recordCoverageGap('word/rare.xml');
    for (let i = 0; i < 3; i++) recordCoverageGap('word/common.xml');

    const lines = summariseCoverageGaps().lines;
    const commonAt = lines.findIndex(l => l.includes('word/common.xml'));
    const rareAt = lines.findIndex(l => l.includes('word/rare.xml'));

    expect(commonAt).toBeLessThan(rareAt);
  });

  it('counts total requests, not distinct gaps', () => {
    recordCoverageGap('a/x.xml');
    recordCoverageGap('a/x.xml');
    recordCoverageGap('b/y.xml');

    expect(summariseCoverageGaps().total).toBe(3);
  });

  it('is stable for equal counts, so the list does not reshuffle between reads', () => {
    recordCoverageGap('b/second.xml');
    recordCoverageGap('a/first.xml');

    const first = summariseCoverageGaps().lines;
    const second = summariseCoverageGaps().lines;

    expect(first).toEqual(second);
    expect(first.findIndex(l => l.includes('a/first.xml'))).toBeLessThan(
      first.findIndex(l => l.includes('b/second.xml'))
    );
  });

  it('keeps counting known keys after the cap, rather than dropping the tail', () => {
    for (let i = 0; i < 250; i++) recordCoverageGap(`part${String.fromCharCode(97 + (i % 26))}${i}/x.xml`);
    recordCoverageGap('word/document.xml');
    recordCoverageGap('word/document.xml');

    const gaps = readCoverageGaps();
    expect(Object.keys(gaps.parts).length).toBeLessThanOrEqual(200);
    // Whatever is tracked keeps counting; nothing throws and nothing is lost silently.
    expect(summariseCoverageGaps().total).toBeGreaterThan(0);
  });
});
