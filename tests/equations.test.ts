import { describe, it, expect } from 'vitest';
import { equationFindings, hasEquations } from '../services/equations';
import { analyzePackage } from '../services/analyzers';
import type { PackageParts } from '../services/packageIntegrity';

/**
 * Office Math faults that render cleanly and state the wrong mathematics.
 *
 * Every block asserts both ways. The negative case matters more here than usual: a
 * correct equation is a very ordinary thing to find in a document, and an analyzer that
 * fired on ordinary equations would bury the four faults that are worth reading.
 */

const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const PART = 'word/document.xml';

/** A document whose single paragraph holds `math`. */
const docWith = (math: string): PackageParts => ({
  [PART]:
    `<?xml version="1.0"?><w:document xmlns:w="${W}" xmlns:m="${M}"><w:body><w:p>` +
    `<m:oMath>${math}</m:oMath>` +
    `</w:p></w:body></w:document>`
});

const codes = (parts: PackageParts) => equationFindings(parts, PART).map(f => f.code);

/** A math run carrying text, which is the ordinary healthy case. */
const run = (text: string) => `<m:r><m:t>${text}</m:t></m:r>`;

// --- n-ary operators -------------------------------------------------------

const nary = (properties: string) =>
  `<m:nary>${properties}<m:sub>${run('1')}</m:sub><m:sup>${run('n')}</m:sup><m:e>${run('x')}</m:e></m:nary>`;

describe('n-ary operators that do not state their operator', () => {
  it('reports an n-ary with no m:naryPr at all', () => {
    expect(codes(docWith(nary('')))).toEqual(['equation/nary-operator-unstated']);
  });

  it('reports an n-ary whose properties omit m:chr', () => {
    // The realistic shape: properties survived a conversion, the character did not.
    const properties = '<m:naryPr><m:limLoc m:val="undOvr"/></m:naryPr>';
    expect(codes(docWith(nary(properties)))).toEqual(['equation/nary-operator-unstated']);
  });

  it('stays quiet when the operator is stated', () => {
    const properties = '<m:naryPr><m:chr m:val="∑"/></m:naryPr>';
    expect(codes(docWith(nary(properties)))).toEqual([]);
  });

  it('is silent and a warning, because the equation draws either way', () => {
    const [found] = equationFindings(docWith(nary('')), PART);
    expect(found.silent).toBe(true);
    expect(found.severity).toBe('warning');
  });

  it('does not name the character a consumer would substitute', () => {
    // ECMA-376 states a default and the SDK schema data does not carry it, so naming one
    // would be a citation this project cannot support. The finding says "unstated".
    const [found] = equationFindings(docWith(nary('')), PART);
    expect(found.message).toContain('does not state its operator');
    expect(found.message).not.toMatch(/defaults to (an? )?(integral|∫)/i);
  });
});

// --- delimiters ------------------------------------------------------------

const delimiter = (properties: string) =>
  `<m:d>${properties}<m:e>${run('x')}</m:e></m:d>`;

describe('delimiters with one bracket customised', () => {
  it('reports an opening bracket set with the closing one defaulted', () => {
    const properties = '<m:dPr><m:begChr m:val="["/></m:dPr>';
    expect(codes(docWith(delimiter(properties)))).toEqual(['equation/delimiter-asymmetric']);
  });

  it('reports a closing bracket set with the opening one defaulted', () => {
    const properties = '<m:dPr><m:endChr m:val="]"/></m:dPr>';
    expect(codes(docWith(delimiter(properties)))).toEqual(['equation/delimiter-asymmetric']);
  });

  it('stays quiet when both ends are stated', () => {
    const properties = '<m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr>';
    expect(codes(docWith(delimiter(properties)))).toEqual([]);
  });

  it('stays quiet when neither end is stated', () => {
    // Both default, so they match. This is the common healthy case and must not fire.
    expect(codes(docWith(delimiter('')))).toEqual([]);
    expect(codes(docWith(delimiter('<m:dPr><m:grow m:val="1"/></m:dPr>')))).toEqual([]);
  });

  it('names which end was stated, so the reader knows what to compare', () => {
    const properties = '<m:dPr><m:begChr m:val="["/></m:dPr>';
    const [found] = equationFindings(docWith(delimiter(properties)), PART);
    expect(found.message).toContain('opening "["');
    expect(found.subject).toMatchObject({ stated: '[' });
  });
});

// --- emptiness -------------------------------------------------------------

describe('equations and runs with nothing in them', () => {
  it('reports an equation containing no elements', () => {
    // Schema-valid: CT_OMath wraps its content group with minOccurs="0".
    expect(codes(docWith(''))).toEqual(['equation/empty-equation']);
  });

  it('reports an equation holding nothing but control properties', () => {
    expect(codes(docWith('<m:ctrlPr/>'))).toEqual(['equation/empty-equation']);
  });

  it('reports a math run with no text', () => {
    expect(codes(docWith('<m:r/>'))).toEqual(['equation/empty-math-run']);
  });

  it('stays quiet for an ordinary equation', () => {
    expect(codes(docWith(run('x') + run('+') + run('y')))).toEqual([]);
  });

  it('does not call a run empty when its text is empty but present', () => {
    // <m:t/> is a run that deliberately contributes nothing, which is different from a
    // run that lost its text element.
    expect(codes(docWith('<m:r><m:t/></m:r>'))).toEqual([]);
  });
});

// --- scope and wiring ------------------------------------------------------

describe('scope', () => {
  it('says nothing about a document with no equations', () => {
    const parts: PackageParts = {
      [PART]: `<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body><w:p/></w:body></w:document>`
    };
    expect(equationFindings(parts, PART)).toEqual([]);
    expect(hasEquations(parts)).toBe(false);
  });

  it('survives a part that is not well-formed', () => {
    expect(() => equationFindings({ [PART]: '<w:document>' }, PART)).not.toThrow();
    expect(equationFindings({ [PART]: '<w:document>' }, PART)).toEqual([]);
  });

  it('looks in headers, footers and notes, not only the main story', () => {
    const equation = `<m:oMath xmlns:m="${M}"><m:nary><m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary></m:oMath>`;
    for (const part of ['word/header1.xml', 'word/footnotes.xml', 'word/endnotes.xml']) {
      const parts: PackageParts = { [part]: `<?xml version="1.0"?><w:root xmlns:w="${W}">${equation}</w:root>` };
      expect(hasEquations(parts), part).toBe(true);
      expect(equationFindings(parts, part).map(f => f.code), part)
        .toContain('equation/nary-operator-unstated');
    }
  });

  it('reaches the registry, and a healthy document produces nothing', () => {
    const healthy = analyzePackage({
      '[Content_Types].xml': '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
      ...docWith(`<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:e>${run('x')}</m:e></m:nary>`)
    });
    expect(healthy.findings.filter(f => f.code.startsWith('equation/'))).toEqual([]);

    const broken = analyzePackage({
      '[Content_Types].xml': '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
      ...docWith(nary(''))
    });
    expect(broken.ran).toContain('equation');
    expect(broken.findings.map(f => f.code)).toContain('equation/nary-operator-unstated');
  });
});
