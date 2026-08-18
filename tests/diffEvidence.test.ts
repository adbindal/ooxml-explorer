import { describe, it, expect } from 'vitest';
import { diffPackages, explainDiff } from '../services/ooxmlDiff';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const doc = (body: string) => `<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`;
const para = (text: string, props = '') =>
  `<w:p>${props ? `<w:pPr>${props}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;

/**
 * These pin the behaviour the diff panel now depends on. Before this wiring the panel
 * sent raw before/after text to the model, so none of it was exercised anywhere.
 */
describe('diff evidence handed to the panel', () => {
  it('calls two saves of the same document equivalent despite textual noise', () => {
    // The case people actually open a diff for: revision ids get rewritten on every
    // save, so a line-by-line diff is full of red for a document nobody edited.
    const before = {
      'word/document.xml': doc(`<w:p w:rsidR="00A1"><w:r><w:t>Hello</w:t></w:r></w:p>`)
    };
    const after = {
      'word/document.xml': doc(
        `<w:p w:rsidR="00FF"><w:proofErr w:type="spellStart"/><w:r><w:t>Hello</w:t></w:r></w:p>`
      )
    };

    const result = diffPackages(before, after);

    expect(result.equivalent).toBe(true);
    expect(explainDiff(result)[0]).toContain('SEMANTICALLY EQUIVALENT');
  });

  it('still reports a real content change', () => {
    const result = diffPackages(
      { 'word/document.xml': doc(para('Hello')) },
      { 'word/document.xml': doc(para('Goodbye')) }
    );

    expect(result.equivalent).toBe(false);
    expect(result.records.length).toBeGreaterThan(0);
  });

  it('every record carries a remediation, so "how do I make them the same" is answerable', () => {
    const result = diffPackages(
      { 'word/document.xml': doc(para('Hello')) },
      { 'word/document.xml': doc(para('Hello', '<w:jc w:val="center"/>')) }
    );

    expect(result.records.length).toBeGreaterThan(0);
    for (const record of result.records) {
      expect(record.remediation.length).toBeGreaterThan(0);
    }
  });

  it('reports a part added and a part removed', () => {
    const result = diffPackages(
      { 'word/document.xml': doc(para('x')) },
      { 'word/document.xml': doc(para('x')), 'word/footer1.xml': doc(para('f')) }
    );

    expect(result.records.some(r => r.kind === 'part-added')).toBe(true);
  });

  it('travels with what it could not establish, so the tier can be capped', () => {
    const result = diffPackages(
      { 'word/document.xml': doc(para('Hello')) },
      { 'word/document.xml': doc(para('Goodbye')) }
    );

    // unresolved is a contract, not an optional extra: a non-empty list is what stops
    // the panel from claiming Verified on a partial derivation.
    expect(Array.isArray(result.unresolved)).toBe(true);
  });

  it('does not throw on malformed XML, so the panel can degrade instead of dying', () => {
    expect(() => diffPackages({ 'word/document.xml': '<w:document><oops>' }, { 'word/document.xml': doc(para('x')) })).not.toThrow();
  });
});
