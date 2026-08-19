import { describe, it, expect } from 'vitest';
import {
  styleReferenceFindings,
  checkExcelStylesheetIntegrity,
  computeStyleReferenceEvidenceForMarkup,
  hasStyleReferences,
  FIRST_CUSTOM_NUMBER_FORMAT_ID
} from '../services/styleReferences';
import { parseExcelStyles } from '../services/excelStyleResolver';
import type { PackageParts } from '../services/packageIntegrity';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const S = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

const styles = (body: string) => `<?xml version="1.0"?><w:styles ${W}>${body}</w:styles>`;
const style = (id: string, extra = '') =>
  `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${id}"/>${extra}</w:style>`;
const document = (body: string) => `<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`;
const para = (props: string) => `<w:p><w:pPr>${props}</w:pPr><w:r><w:t>text</w:t></w:r></w:p>`;

const numbering = (ids: string[]) =>
  `<?xml version="1.0"?><w:numbering ${W}>${ids.map(i => `<w:num w:numId="${i}"><w:abstractNumId w:val="0"/></w:num>`).join('')}</w:numbering>`;

describe('Word style references', () => {
  it('reports a paragraph style that is referenced but never defined', () => {
    // The most common defect in generated documents: the generator writes the
    // reference and forgets the definition. Word applies Normal and the file opens.
    const parts: PackageParts = {
      'word/styles.xml': styles(style('Normal')),
      'word/document.xml': document(para('<w:pStyle w:val="Heading1"/>'))
    };

    const problem = styleReferenceFindings(parts).find(p => p.code === 'styleRef/missing-paragraph-style');
    expect(problem?.subject?.styleId).toBe('Heading1');
    expect(problem?.silent).toBe(true);
    expect(problem?.message).toContain('Normal');
  });

  it('says nothing when the style exists', () => {
    const parts: PackageParts = {
      'word/styles.xml': styles(style('Normal') + style('Heading1')),
      'word/document.xml': document(para('<w:pStyle w:val="Heading1"/>'))
    };

    expect(styleReferenceFindings(parts)).toEqual([]);
  });

  it('reports one finding per missing style, not one per use', () => {
    // Forty paragraphs referencing one broken style is one broken style. Forty
    // identical findings would bury everything else in the report.
    const parts: PackageParts = {
      'word/styles.xml': styles(style('Normal')),
      'word/document.xml': document(
        Array.from({ length: 5 }, () => para('<w:pStyle w:val="Heading1"/>')).join('')
      )
    };

    expect(styleReferenceFindings(parts).filter(p => p.code === 'styleRef/missing-paragraph-style')).toHaveLength(1);
  });

  it('distinguishes paragraph, character and table style references', () => {
    const parts: PackageParts = {
      'word/styles.xml': styles(style('Normal')),
      'word/document.xml': document(
        `<w:p><w:pPr><w:pStyle w:val="MissingP"/></w:pPr><w:r><w:rPr><w:rStyle w:val="MissingR"/></w:rPr><w:t>x</w:t></w:r></w:p>` +
          `<w:tbl><w:tblPr><w:tblStyle w:val="MissingT"/></w:tblPr></w:tbl>`
      )
    };

    expect(styleReferenceFindings(parts).map(p => p.code).sort()).toEqual([
      'styleRef/missing-character-style',
      'styleRef/missing-paragraph-style',
      'styleRef/missing-table-style'
    ]);
  });

  it('checks every Word story, not just document.xml', () => {
    const parts: PackageParts = {
      'word/styles.xml': styles(style('Normal')),
      'word/header1.xml': document(para('<w:pStyle w:val="HeaderStyle"/>'))
    };

    expect(styleReferenceFindings(parts).map(p => p.subject?.styleId)).toContain('HeaderStyle');
  });
});

describe('numbering references', () => {
  it('reports a numId that numbering.xml does not declare', () => {
    const parts: PackageParts = {
      'word/styles.xml': styles(style('Normal')),
      'word/numbering.xml': numbering(['1']),
      'word/document.xml': document(para('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>'))
    };

    const problem = styleReferenceFindings(parts).find(p => p.code === 'styleRef/missing-numbering');
    expect(problem?.subject?.numId).toBe('7');
  });

  it('does NOT report numId="0", which means remove numbering', () => {
    // The specification is explicit that 0 never points at a definition. Treating it
    // as a lookup fires on every document that has ever had a list removed from it -
    // which is most of them.
    const parts: PackageParts = {
      'word/styles.xml': styles(style('Normal')),
      'word/numbering.xml': numbering(['1']),
      'word/document.xml': document(para('<w:numPr><w:numId w:val="0"/></w:numPr>'))
    };

    expect(styleReferenceFindings(parts).map(p => p.code)).not.toContain('styleRef/missing-numbering');
  });

  it('does not judge numbering at all when numbering.xml is absent', () => {
    // A missing part is packageIntegrity's finding. Reporting it here too would
    // double-count one fault and blame the paragraphs for it.
    const parts: PackageParts = {
      'word/styles.xml': styles(style('Normal')),
      'word/document.xml': document(para('<w:numPr><w:numId w:val="7"/></w:numPr>'))
    };

    expect(styleReferenceFindings(parts).map(p => p.code)).not.toContain('styleRef/missing-numbering');
  });
});

describe('stylesheet internal references', () => {
  it('reports a basedOn naming a style that does not exist', () => {
    const parts: PackageParts = {
      'word/styles.xml': styles(style('Normal') + style('Body', '<w:basedOn w:val="Ghost"/>')),
      'word/document.xml': document(para(''))
    };

    const problem = styleReferenceFindings(parts).find(p => p.code === 'styleRef/missing-based-on');
    expect(problem?.subject?.target).toBe('Ghost');
    expect(problem?.message).toContain('root style');
  });

  it('reports a broken link between paired styles', () => {
    const parts: PackageParts = {
      'word/styles.xml': styles(style('Normal') + style('Quote', '<w:link w:val="QuoteChar"/>')),
      'word/document.xml': document(para(''))
    };

    expect(styleReferenceFindings(parts).map(p => p.code)).toContain('styleRef/missing-linked-style');
  });

  it('says nothing when basedOn resolves', () => {
    const parts: PackageParts = {
      'word/styles.xml': styles(style('Normal') + style('Body', '<w:basedOn w:val="Normal"/>')),
      'word/document.xml': document(para(''))
    };

    expect(styleReferenceFindings(parts)).toEqual([]);
  });
});

describe('Excel format references', () => {
  const stylesXml = (body: string) => `<?xml version="1.0"?><styleSheet ${S}>${body}</styleSheet>`;
  const minimal =
    '<fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders>' +
    '<cellStyleXfs count="1"><xf/></cellStyleXfs>';

  const sheet = (cells: string) =>
    `<?xml version="1.0"?><worksheet ${S}><sheetData><row r="1">${cells}</row></sheetData></worksheet>`;

  it('reports a cell whose format index is past the end of cellXfs', () => {
    const parts: PackageParts = {
      'xl/styles.xml': stylesXml(`${minimal}<cellXfs count="2"><xf/><xf/></cellXfs>`),
      'xl/worksheets/sheet1.xml': sheet('<c r="A1" s="5"><v>1</v></c>')
    };

    const problem = styleReferenceFindings(parts).find(p => p.code === 'styleRef/cell-format-out-of-range');
    expect(problem?.subject?.index).toBe('5');
    expect(problem?.message).toContain('0–1');
  });

  it('accepts the last valid index, which is where an off-by-one would show', () => {
    const parts: PackageParts = {
      'xl/styles.xml': stylesXml(`${minimal}<cellXfs count="2"><xf/><xf/></cellXfs>`),
      'xl/worksheets/sheet1.xml': sheet('<c r="A1" s="1"><v>1</v></c>')
    };

    expect(styleReferenceFindings(parts)).toEqual([]);
  });

  it('reports the index exactly one past the end, which is where an off-by-one lives', () => {
    // cellXfs has 2 entries, so 2 is the first invalid index. The earlier test used 5,
    // which is caught under either a < or a <= bound - so it proved nothing about the
    // boundary itself.
    const parts: PackageParts = {
      'xl/styles.xml': stylesXml(`${minimal}<cellXfs count="2"><xf/><xf/></cellXfs>`),
      'xl/worksheets/sheet1.xml': sheet('<c r="A1" s="2"><v>1</v></c>')
    };

    expect(styleReferenceFindings(parts).map(p => p.code)).toContain('styleRef/cell-format-out-of-range');
  });

  it('reports a custom number format that is not declared', () => {
    const sheetObj = parseExcelStyles(
      stylesXml(`${minimal}<cellXfs count="1"><xf numFmtId="165"/></cellXfs>`)
    );
    const problem = checkExcelStylesheetIntegrity(sheetObj, 'xl/styles.xml').find(
      p => p.code === 'styleRef/missing-number-format'
    );

    expect(problem?.subject?.numFmtId).toBe('165');
    expect(problem?.message).toContain('General');
  });

  it('does NOT report a built-in number format id', () => {
    // Ids below 164 are built into Excel and declared nowhere. Checking them against
    // numFmts reports every ordinary workbook as broken.
    const sheetObj = parseExcelStyles(
      stylesXml(`${minimal}<cellXfs count="1"><xf numFmtId="14"/></cellXfs>`)
    );

    expect(checkExcelStylesheetIntegrity(sheetObj, 'xl/styles.xml')).toEqual([]);
  });

  it('treats 164 itself as custom, since that is the boundary', () => {
    const sheetObj = parseExcelStyles(
      stylesXml(`${minimal}<cellXfs count="1"><xf numFmtId="${FIRST_CUSTOM_NUMBER_FORMAT_ID}"/></cellXfs>`)
    );

    expect(checkExcelStylesheetIntegrity(sheetObj, 'xl/styles.xml').map(p => p.code)).toContain(
      'styleRef/missing-number-format'
    );
  });

  it('reports a format referencing a font that is not declared', () => {
    const sheetObj = parseExcelStyles(
      stylesXml(`${minimal}<cellXfs count="1"><xf fontId="9"/></cellXfs>`)
    );

    const problem = checkExcelStylesheetIntegrity(sheetObj, 'xl/styles.xml').find(
      p => p.code === 'styleRef/component-out-of-range'
    );
    expect(problem?.message).toContain('font 9');
  });
});

describe('evidence for the panel', () => {
  it('returns null for a package with no stylesheet', () => {
    expect(hasStyleReferences({ 'word/document.xml': document(para('')) })).toBe(false);
    expect(computeStyleReferenceEvidenceForMarkup({ 'word/document.xml': document(para('')) })).toBeNull();
  });

  it('says plainly when everything resolves', () => {
    const evidence = computeStyleReferenceEvidenceForMarkup({
      'word/styles.xml': styles(style('Normal')),
      'word/document.xml': document(para('<w:pStyle w:val="Normal"/>'))
    });

    expect(evidence!.lines.some(l => l.includes('resolves to a definition that exists'))).toBe(true);
  });

  it('frames the failure as plainer-than-intended rather than broken', () => {
    const evidence = computeStyleReferenceEvidenceForMarkup({
      'word/styles.xml': styles(style('Normal')),
      'word/document.xml': document(para('<w:pStyle w:val="Heading1"/>'))
    });

    expect(evidence!.lines.some(l => l.includes('plainer than intended'))).toBe(true);
  });

  it('caps the claim to the stylesheets it was actually given', () => {
    const evidence = computeStyleReferenceEvidenceForMarkup({
      'word/styles.xml': styles(style('Normal')),
      'word/document.xml': document(para('<w:pStyle w:val="Heading1"/>'))
    });

    expect(evidence!.unresolved.some(u => u.includes('not loaded here'))).toBe(true);
  });
});
