import { describe, it, expect } from 'vitest';
import {
  loadWordContext,
  analyzeParagraphAt,
  locateParagraphByMarkup,
  computeEvidenceForMarkup
} from '../services/wordFormattingAnalysis';
import type { PackageParts } from '../services/packageIntegrity';
import { selectEvidenceTier } from '../services/aiService';

/**
 * End-to-end tests for the composition layer.
 *
 * The individual engines are unit-tested elsewhere. What these cover is the join:
 * that a real package's parts are found, that the right layers are assembled in the
 * right order, and — most importantly — that what could NOT be resolved is reported
 * rather than silently omitted.
 */

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

const STYLES = `<?xml version="1.0"?>
<w:styles xmlns:w="${W}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:rPr><w:sz w:val="32"/><w:b/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>
  </w:style>
  <w:style w:type="table" w:styleId="GridTable">
    <w:tblPr><w:tblStyleRowBandSize w:val="1"/></w:tblPr>
    <w:tblStylePr w:type="wholeTable"><w:rPr><w:sz w:val="20"/></w:rPr></w:tblStylePr>
    <w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr>
  </w:style>
  <w:style w:type="table" w:styleId="NoBandSize">
    <w:tblStylePr w:type="band1Horz"><w:rPr><w:i/></w:rPr></w:tblStylePr>
  </w:style>
</w:styles>`;

const NUMBERING = `<?xml version="1.0"?>
<w:numbering xmlns:w="${W}">
  <w:abstractNum w:abstractNumId="5">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
      <w:pPr><w:ind w:left="720"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="5"/></w:num>
</w:numbering>`;

const document = (body: string) => `<?xml version="1.0"?>
<w:document xmlns:w="${W}" xmlns:mc="${MC}"><w:body>${body}</w:body></w:document>`;

const pkg = (body: string, over: Partial<PackageParts> = {}): PackageParts => ({
  'word/styles.xml': STYLES,
  'word/numbering.xml': NUMBERING,
  'word/document.xml': document(body),
  ...over
});

describe('loading the package', () => {
  it('finds styles, numbering and the document', () => {
    const ctx = loadWordContext(pkg('<w:p/>'));
    expect(ctx.styles.styles.size).toBeGreaterThan(0);
    expect(ctx.numbering).not.toBeNull();
    expect(ctx.document).not.toBeNull();
    expect(ctx.unresolved).toEqual([]);
  });

  it('reports a missing numbering part instead of failing', () => {
    const parts = pkg('<w:p/>');
    delete parts['word/numbering.xml'];
    expect(loadWordContext(parts).unresolved.join(' ')).toContain('numbering.xml');
  });

  it('reports a missing styles part and still resolves direct formatting', () => {
    const parts = pkg('<w:p><w:r><w:rPr><w:sz w:val="48"/></w:rPr></w:r></w:p>');
    delete parts['word/styles.xml'];
    const ctx = loadWordContext(parts);
    expect(ctx.unresolved.join(' ')).toContain('styles.xml');
    expect(analyzeParagraphAt(ctx, 0)!.run!.get('sz')).toBe('48');
  });

  it('reports malformed document.xml rather than throwing', () => {
    const ctx = loadWordContext(pkg('', { 'word/document.xml': '<w:document><unclosed>' }));
    expect(ctx.unresolved.join(' ')).toContain('not well-formed');
    expect(ctx.document).toBeNull();
  });

  it('resolves markup compatibility before anything walks the tree', () => {
    // A shape written twice must not be counted twice. Without the MCE pass the
    // fallback paragraph would appear as a second paragraph.
    const ctx = loadWordContext(pkg(`
      <mc:AlternateContent>
        <mc:Choice Requires="w"><w:p><w:r><w:t>modern</w:t></w:r></w:p></mc:Choice>
        <mc:Fallback><w:p><w:r><w:t>legacy</w:t></w:r></w:p></mc:Fallback>
      </mc:AlternateContent>`));
    const paragraphs = ctx.document!.getElementsByTagNameNS(W, 'p');
    expect(paragraphs.length).toBe(1);
    expect(paragraphs[0].textContent).toBe('modern');
  });
});

describe('resolving a plain paragraph', () => {
  it('applies the paragraph style over document defaults', () => {
    const ctx = loadWordContext(pkg('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>'));
    const a = analyzeParagraphAt(ctx, 0)!;
    expect(a.run!.get('sz')).toBe('32');
    expect(a.run!.isOn('b')).toBe(true);
  });

  it('falls back to document defaults when the paragraph has no style', () => {
    const ctx = loadWordContext(pkg('<w:p><w:r><w:t>Hi</w:t></w:r></w:p>'));
    expect(analyzeParagraphAt(ctx, 0)!.run!.get('sz')).toBe('22');
  });

  it('explains where each property came from', () => {
    const ctx = loadWordContext(pkg('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>'));
    expect(analyzeParagraphAt(ctx, 0)!.explanation.join('\n')).toContain('sz = 32 (from style:Heading1)');
  });
});

describe('numbering', () => {
  it('resolves numbering declared directly on the paragraph', () => {
    const ctx = loadWordContext(pkg('<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:p>'));
    const a = analyzeParagraphAt(ctx, 0)!;
    expect(a.numbering?.numFmt).toBe('decimal');
    expect(a.explanation.join('\n')).toContain('Numbered: numId 1, level 0');
  });

  it('inherits numbering from the paragraph style when the paragraph says nothing', () => {
    const ctx = loadWordContext(pkg('<w:p><w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr></w:p>'));
    expect(analyzeParagraphAt(ctx, 0)!.numbering?.numId).toBe('1');
  });

  it('treats numId 0 as removal, not as inherited numbering', () => {
    // The paragraph uses a numbered style but explicitly cancels numbering.
    const ctx = loadWordContext(pkg(
      '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:p>'
    ));
    const a = analyzeParagraphAt(ctx, 0)!;
    expect(a.numbering).toBeNull();
    expect(a.unresolved.join(' ')).not.toContain('could not be resolved');
  });

  it('reports a numbered paragraph whose numbering part is missing', () => {
    const parts = pkg('<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:p>');
    delete parts['word/numbering.xml'];
    const a = analyzeParagraphAt(loadWordContext(parts), 0)!;
    expect(a.unresolved.join(' ')).toContain('numbering part is missing');
  });
});

describe('tables', () => {
  const tableDoc = (styleId: string) => pkg(`
    <w:tbl>
      <w:tblPr><w:tblStyle w:val="${styleId}"/><w:tblLook w:firstRow="1"/></w:tblPr>
      <w:tr><w:tc><w:p><w:r><w:t>head</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>body</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>`);

  it('detects the cell position and reports it', () => {
    const a = analyzeParagraphAt(loadWordContext(tableDoc('GridTable')), 0)!;
    expect(a.explanation.join('\n')).toContain('row 1 of 2, column 1 of 1');
  });

  it('applies whole-table then first-row conditional formatting', () => {
    const a = analyzeParagraphAt(loadWordContext(tableDoc('GridTable')), 0)!;
    expect(a.tableFormats).toContain('wholeTable');
    expect(a.tableFormats).toContain('firstRow');
    expect(a.run!.isOn('b')).toBe(true);
  });

  it('does not apply firstRow to the second row', () => {
    const a = analyzeParagraphAt(loadWordContext(tableDoc('GridTable')), 1)!;
    expect(a.tableFormats).not.toContain('firstRow');
  });

  it('surfaces the band-size deviation that silently disables banding', () => {
    // NoBandSize defines band1Horz but omits tblStyleRowBandSize. Word treats the
    // absent value as 0, so the banding never renders — the single most common
    // "my table style doesn't work" cause.
    const a = analyzeParagraphAt(loadWordContext(tableDoc('NoBandSize')), 1)!;
    expect(a.tableFormats.some(t => t.endsWith('Horz'))).toBe(false);
    expect(a.unresolved.join(' ')).toContain('disables row banding');
  });

  it('reports a table referencing a style that does not exist', () => {
    const a = analyzeParagraphAt(loadWordContext(tableDoc('Ghost')), 0)!;
    expect(a.unresolved.join(' ')).toContain('not defined in styles.xml');
  });

  it('says nothing about tables for a paragraph outside one', () => {
    const a = analyzeParagraphAt(loadWordContext(pkg('<w:p><w:r><w:t>x</w:t></w:r></w:p>')), 0)!;
    expect(a.tableFormats).toEqual([]);
  });
});

describe('honesty of the output', () => {
  it('labels the unresolved section so a model does not assert it', () => {
    const parts = pkg('<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:p>');
    delete parts['word/numbering.xml'];
    const a = analyzeParagraphAt(loadWordContext(parts), 0)!;
    expect(a.explanation.join('\n')).toContain('do not assert these');
  });

  it('omits the unresolved section entirely when nothing is unresolved', () => {
    const a = analyzeParagraphAt(loadWordContext(pkg('<w:p><w:r><w:t>x</w:t></w:r></w:p>')), 0)!;
    expect(a.unresolved).toEqual([]);
    expect(a.explanation.join('\n')).not.toContain('do not assert');
  });

  it('returns null for a paragraph index that does not exist', () => {
    expect(analyzeParagraphAt(loadWordContext(pkg('<w:p/>')), 99)).toBeNull();
  });
});

describe('locating a paragraph from a snippet', () => {
  const doc = (body: string) => loadWordContext(pkg(body)).document!;

  it('matches a paragraph given its own markup', () => {
    const d = doc('<w:p><w:r><w:t>alpha</w:t></w:r></w:p>');
    const found = locateParagraphByMarkup(d, '<w:p><w:r><w:t>alpha</w:t></w:r></w:p>');
    expect(found?.paragraph.textContent).toBe('alpha');
  });

  it('tolerates the pretty-printing the editor applies', () => {
    // The panel holds a formatted snippet; the document is parsed from raw text.
    const d = doc('<w:p><w:r><w:t>alpha</w:t></w:r></w:p>');
    const found = locateParagraphByMarkup(d, `<w:p>
      <w:r>
        <w:t>alpha</w:t>
      </w:r>
    </w:p>`);
    expect(found?.paragraph.textContent).toBe('alpha');
  });

  it('finds the containing paragraph when given a run', () => {
    const d = doc('<w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:p><w:r><w:t>beta</w:t></w:r></w:p>');
    const found = locateParagraphByMarkup(d, '<w:r><w:t>beta</w:t></w:r>');
    expect(found?.paragraph.textContent).toBe('beta');
    expect(found?.run).toBeDefined();
  });

  it('REFUSES to guess between identical paragraphs', () => {
    // Empty paragraphs are everywhere in real documents. Picking one at random and
    // reporting its formatting as Verified would be a confidently wrong answer.
    const d = doc('<w:p/><w:p/><w:p/>');
    expect(locateParagraphByMarkup(d, '<w:p/>')).toBeNull();
  });

  it('refuses when a snippet appears in several paragraphs', () => {
    const d = doc('<w:p><w:r><w:t>same</w:t></w:r></w:p><w:p><w:r><w:t>same</w:t></w:r></w:p>');
    expect(locateParagraphByMarkup(d, '<w:r><w:t>same</w:t></w:r>')).toBeNull();
  });

  it('returns the paragraph without a run when the run is ambiguous inside it', () => {
    const d = doc('<w:p><w:r><w:t>x</w:t></w:r><w:r><w:t>x</w:t></w:r></w:p>');
    const found = locateParagraphByMarkup(d, '<w:r><w:t>x</w:t></w:r>');
    expect(found?.paragraph).toBeDefined();
    expect(found?.run).toBeUndefined();
  });

  it('returns null for markup that is not in the document', () => {
    expect(locateParagraphByMarkup(doc('<w:p/>'), '<w:tbl/>')).toBeNull();
  });

  it('returns null for an empty snippet', () => {
    expect(locateParagraphByMarkup(doc('<w:p/>'), '   ')).toBeNull();
  });
});

describe('computeEvidenceForMarkup', () => {
  it('produces evidence lines and an unresolved list', () => {
    const evidence = computeEvidenceForMarkup(
      pkg('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>'),
      '<w:r><w:t>Title</w:t></w:r>'
    )!;
    expect(evidence.lines.join('\n')).toContain('sz = 32 (from style:Heading1)');
    expect(evidence.unresolved).toEqual([]);
  });

  it('returns null rather than empty evidence when the snippet cannot be located', () => {
    // The caller must fall back to its ordinary path, not show a Verified answer
    // built on nothing.
    expect(computeEvidenceForMarkup(pkg('<w:p/>'), '<w:tbl/>')).toBeNull();
  });

  it('returns null for a package with no document part', () => {
    const parts = pkg('<w:p/>');
    delete parts['word/document.xml'];
    expect(computeEvidenceForMarkup(parts, '<w:p/>')).toBeNull();
  });
})

describe('the full chain: computed evidence earns the Verified tier', () => {
  it('a resolvable element yields evidence that selects verified', () => {
    // This is the assertion that proves the wiring closes. Anything weaker leaves
    // "Verified tier shipped" true in the code and false for the user.
    const evidence = computeEvidenceForMarkup(
      pkg('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>'),
      '<w:r><w:t>Title</w:t></w:r>'
    )!;
    expect(selectEvidenceTier(false, evidence)).toBe('verified');
  });

  it('an element whose numbering part is missing is capped at grounded', () => {
    const parts = pkg('<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>');
    delete parts['word/numbering.xml'];
    const evidence = computeEvidenceForMarkup(parts, '<w:r><w:t>Item</w:t></w:r>')!;
    expect(evidence.unresolved.length).toBeGreaterThan(0);
    expect(selectEvidenceTier(false, evidence)).toBe('grounded');
  });

  it('an unlocatable element falls back to unverified, not a hollow verified', () => {
    const evidence = computeEvidenceForMarkup(pkg('<w:p/><w:p/>'), '<w:p/>');
    expect(evidence).toBeNull();
    expect(selectEvidenceTier(false, evidence)).toBe('unverified');
  });
});
