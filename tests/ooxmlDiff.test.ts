import { describe, it, expect } from 'vitest';
import { diffPackages, atomize, normalizePart, explainDiff } from '../services/ooxmlDiff';
import { W_NAMESPACE } from '../services/wordStyleResolver';
import type { PackageParts } from '../services/packageIntegrity';

/**
 * The output of this module is consumed by another program, so these tests are as much
 * about the shape of the payload as about correctness. A change record a caller cannot
 * act on is a bug even when the difference it describes is real.
 */

const W = W_NAMESPACE;
const STYLES = `<?xml version="1.0"?>
<w:styles xmlns:w="${W}">
  <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Plain"><w:pPr><w:ind w:left="0"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Indented"><w:pPr><w:ind w:left="720"/></w:pPr></w:style>
</w:styles>`;

const doc = (body: string) =>
  `<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`;

const pkg = (body: string): PackageParts => ({
  'word/styles.xml': STYLES,
  'word/document.xml': doc(body)
});

const para = (text: string, pPr = '') =>
  `<w:p>${pPr}<w:r><w:t>${text}</w:t></w:r></w:p>`;

const parse = (xml: string) => new DOMParser().parseFromString(xml, 'application/xml');

describe('normalization reports what it discards', () => {
  it('strips revision-save ids and counts them', () => {
    const d = parse(doc('<w:p w:rsidR="00AB12" w:rsidRDefault="00AB12"><w:r w:rsidR="00CD34"><w:t>x</w:t></w:r></w:p>'));
    const report = normalizePart(d);
    expect(report.find(r => r.rule.includes('revision-save'))?.count).toBe(3);
  });

  it('strips spell-check state and stale layout caches', () => {
    const d = parse(doc('<w:p><w:proofErr w:type="spellStart"/><w:r><w:t>x</w:t></w:r><w:lastRenderedPageBreak/></w:p>'));
    const rules = normalizePart(d).map(r => r.rule);
    expect(rules).toContain('proofErr removed');
    expect(rules).toContain('lastRenderedPageBreak removed');
  });

  it('reports nothing for a clean part', () => {
    expect(normalizePart(parse(doc(para('x'))))).toEqual([]);
  });
});

describe('atomization defeats run splitting', () => {
  it('produces the same atoms whether text is one run or many', () => {
    const one = atomize(parse(doc('<w:p><w:r><w:t>Hello</w:t></w:r></w:p>')));
    const many = atomize(parse(doc(
      '<w:p><w:r><w:t>He</w:t></w:r><w:r><w:t>ll</w:t></w:r><w:r><w:t>o</w:t></w:r></w:p>'
    )));
    expect(one.map(a => a.value).join('')).toBe(many.map(a => a.value).join(''));
    expect(one.length).toBe(many.length);
  });

  it('treats the paragraph mark as content, so a split is a real change', () => {
    const joined = atomize(parse(doc('<w:p><w:r><w:t>ab</w:t></w:r></w:p>')));
    const split = atomize(parse(doc('<w:p><w:r><w:t>a</w:t></w:r></w:p><w:p><w:r><w:t>b</w:t></w:r></w:p>')));
    expect(split.length).toBeGreaterThan(joined.length);
  });

  it('records embedded objects as single atoms', () => {
    const atoms = atomize(parse(doc('<w:p><w:r><w:drawing/></w:r></w:p>')));
    expect(atoms.some(a => a.kind === 'object' && a.value === 'drawing')).toBe(true);
  });
});

describe('the noise problem this module exists for', () => {
  it('reports NO changes for a document that was merely re-saved', () => {
    // Re-saving in Word rewrites revision ids across the file. A textual diff calls
    // this a heavily changed document; a reader sees nothing.
    const before = pkg('<w:p w:rsidR="00AA11"><w:r w:rsidR="00AA11"><w:t>Same text</w:t></w:r></w:p>');
    const after = pkg('<w:p w:rsidR="00BB22"><w:r w:rsidR="00CC33"><w:t>Same text</w:t></w:r></w:p>');
    const result = diffPackages(before, after);
    expect(result.records).toEqual([]);
  });

  it('says how much it filtered, so an empty diff is not ambiguous', () => {
    const before = pkg('<w:p w:rsidR="00AA11"><w:r><w:t>x</w:t></w:r></w:p>');
    const after = pkg('<w:p w:rsidR="00BB22"><w:r><w:t>x</w:t></w:r></w:p>');
    const result = diffPackages(before, after);
    expect(result.normalized.find(n => n.rule.includes('revision-save'))?.count).toBe(2);
    expect(explainDiff(result).join(' ')).toContain('ignored as rendering-irrelevant');
  });

  it('reports NO changes when identical text is split across different runs', () => {
    const before = pkg('<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>');
    const after = pkg('<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r></w:p>');
    expect(diffPackages(before, after).records).toEqual([]);
  });
});

describe('content changes', () => {
  it('reports an insertion with its text and location', () => {
    const result = diffPackages(pkg(para('Hello')), pkg(para('Hello there')));
    const record = result.records.find(r => r.kind === 'content-inserted')!;
    expect(record.after).toBe(' there');
    expect(record.part).toBe('word/document.xml');
    expect(record.location).toBe('paragraph[0]');
    expect(record.visible).toBe(true);
  });

  it('reports a deletion', () => {
    const result = diffPackages(pkg(para('Hello there')), pkg(para('Hello')));
    expect(result.records.find(r => r.kind === 'content-deleted')?.before).toBe(' there');
  });

  it('reports an added part', () => {
    const before = pkg(para('x'));
    const after = { ...pkg(para('x')), 'word/header1.xml': doc(para('h')) };
    const record = diffPackages(before, after).records.find(r => r.kind === 'part-added')!;
    expect(record.part).toBe('word/header1.xml');
  });

  it('reports a removed part', () => {
    const before = { ...pkg(para('x')), 'word/footer1.xml': doc(para('f')) };
    expect(diffPackages(before, pkg(para('x'))).records.some(r => r.kind === 'part-removed')).toBe(true);
  });
});

describe('formatting changes are diffed on RESOLVED values', () => {
  it('reports a style swap as the property that actually changed', () => {
    // The markup change is a pStyle value; the *effect* is an indent change, which is
    // what someone chasing a visual regression needs to see.
    const result = diffPackages(
      pkg(para('x', '<w:pPr><w:pStyle w:val="Plain"/></w:pPr>')),
      pkg(para('x', '<w:pPr><w:pStyle w:val="Indented"/></w:pPr>'))
    );
    const ind = result.records.find(r => r.kind === 'formatting-changed' && r.property === 'ind');
    expect(ind).toBeDefined();
    // Reported as the resolved attribute state, because w:ind carries its meaning in
    // w:left rather than a w:val - comparing val alone would report no change at all.
    expect(ind!.before).toBe('w:left=0');
    expect(ind!.after).toBe('w:left=720');
  });

  it('reports no formatting change when different markup resolves the same', () => {
    // Direct formatting versus a style that sets the identical value. Comparing markup
    // would report a change; comparing resolved values correctly reports none.
    const result = diffPackages(
      pkg(para('x', '<w:pPr><w:pStyle w:val="Indented"/></w:pPr>')),
      pkg(para('x', '<w:pPr><w:ind w:left="720"/></w:pPr>'))
    );
    expect(result.records.filter(r => r.kind === 'formatting-changed')).toEqual([]);
  });

  it('carries the property name so a caller can act on it', () => {
    const result = diffPackages(
      pkg(para('x', '<w:pPr><w:jc w:val="left"/></w:pPr>')),
      pkg(para('x', '<w:pPr><w:jc w:val="center"/></w:pPr>'))
    );
    const jc = result.records.find(r => r.property === 'jc')!;
    expect(jc.before).toBe('left');
    expect(jc.after).toBe('center');
  });
});

describe('the payload shape, for agent-to-agent handoff', () => {
  it('emits flat typed records rather than prose', () => {
    const result = diffPackages(pkg(para('a')), pkg(para('b')));
    for (const record of result.records) {
      expect(typeof record.kind).toBe('string');
      expect(typeof record.part).toBe('string');
      expect(typeof record.location).toBe('string');
      expect(typeof record.visible).toBe('boolean');
      expect(record).toHaveProperty('before');
      expect(record).toHaveProperty('after');
    }
  });

  it('is JSON-serialisable without loss', () => {
    const result = diffPackages(pkg(para('a')), pkg(para('b')));
    expect(JSON.parse(JSON.stringify(result))).toEqual(result);
  });

  it('derives the prose from the records, so the two cannot drift', () => {
    const result = diffPackages(pkg(para('Hello')), pkg(para('Hello there')));
    expect(explainDiff(result).join('\n')).toContain('inserted " there"');
  });

  it('carries unresolved items as a contract, not a footnote', () => {
    const before = pkg(para('x'));
    const after = { ...pkg(para('x')), 'word/document.xml': '<w:document><unclosed>' };
    const result = diffPackages(before, after);
    expect(result.unresolved.join(' ')).toContain('could not be parsed');
  });

  it('lists the parts it examined', () => {
    expect(diffPackages(pkg(para('a')), pkg(para('b'))).parts).toEqual(['word/document.xml']);
  });
});
