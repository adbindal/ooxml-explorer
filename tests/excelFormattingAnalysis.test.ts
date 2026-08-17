import { describe, it, expect } from 'vitest';
import {
  loadExcelContext,
  parseCellReference,
  locateCellByMarkup,
  analyzeCell,
  computeExcelEvidenceForMarkup
} from '../services/excelFormattingAnalysis';
import { selectEvidenceTier } from '../services/aiService';
import { S_NAMESPACE } from '../services/excelStyleResolver';
import type { PackageParts } from '../services/packageIntegrity';

const STYLES = `<?xml version="1.0"?>
<styleSheet xmlns="${S_NAMESPACE}">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>
  <fonts count="2"><font><sz val="11"/></font><font><b/></font></fonts>
  <fills count="1"><fill/></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

const SHARED = `<?xml version="1.0"?>
<sst xmlns="${S_NAMESPACE}" count="2" uniqueCount="2">
  <si><t>Hello</t></si>
  <si><r><t>Wor</t></r><r><t>ld</t></r></si>
</sst>`;

const WORKBOOK = `<?xml version="1.0"?>
<workbook xmlns="${S_NAMESPACE}"><workbookPr date1904="0"/></workbook>`;

const sheet = (body: string, cols = '') => `<?xml version="1.0"?>
<worksheet xmlns="${S_NAMESPACE}">${cols}<sheetData>${body}</sheetData></worksheet>`;

const pkg = (sheetXml: string, over: Partial<PackageParts> = {}): PackageParts => ({
  'xl/styles.xml': STYLES,
  'xl/workbook.xml': WORKBOOK,
  'xl/sharedStrings.xml': SHARED,
  'xl/worksheets/sheet1.xml': sheetXml,
  ...over
});

describe('cell references', () => {
  it('parses single and multi-letter columns', () => {
    expect(parseCellReference('A1')).toEqual({ column: 1, row: 1 });
    expect(parseCellReference('Z9')).toEqual({ column: 26, row: 9 });
    expect(parseCellReference('AA1')).toEqual({ column: 27, row: 1 });
    expect(parseCellReference('BC12')).toEqual({ column: 55, row: 12 });
  });

  it('returns null for junk', () => {
    expect(parseCellReference('1A')).toBeNull();
    expect(parseCellReference('')).toBeNull();
  });
});

describe('loading the package', () => {
  it('finds styles, workbook, shared strings and worksheets', () => {
    const ctx = loadExcelContext(pkg(sheet('<row r="1"><c r="A1"/></row>')));
    expect(ctx.styles.cellXfs.length).toBe(4);
    expect(ctx.sharedStrings).toEqual(['Hello', 'World']);
    expect(ctx.sheets.size).toBe(1);
    expect(ctx.unresolved).toEqual([]);
  });

  it('concatenates rich-text runs into one shared string', () => {
    // The same visible string can be one run or twenty, exactly as with Word's w:r.
    expect(loadExcelContext(pkg(sheet(''))).sharedStrings[1]).toBe('World');
  });

  it('reports a missing styles part', () => {
    const parts = pkg(sheet(''));
    delete parts['xl/styles.xml'];
    expect(loadExcelContext(parts).unresolved.join(' ')).toContain('styles.xml');
  });

  it('reports a missing workbook and assumes the default date system', () => {
    const parts = pkg(sheet(''));
    delete parts['xl/workbook.xml'];
    const ctx = loadExcelContext(parts);
    expect(ctx.unresolved.join(' ')).toContain('date system cannot be determined');
    expect(ctx.dateSystem).toBe('1900-compat');
  });

  it('reports a package with no worksheets', () => {
    const parts = pkg(sheet(''));
    delete parts['xl/worksheets/sheet1.xml'];
    expect(loadExcelContext(parts).unresolved.join(' ')).toContain('no worksheet parts');
  });
});

describe('resolving a cell', () => {
  const analyze = (parts: PackageParts, ref: string) => {
    const ctx = loadExcelContext(parts);
    const sheetDoc = [...ctx.sheets.values()][0];
    const cell = Array.from(sheetDoc.getElementsByTagNameNS(S_NAMESPACE, 'c'))
      .find(c => c.getAttribute('r') === ref)!;
    return analyzeCell(ctx, sheetDoc, cell);
  };

  it('uses the cell style index when present', () => {
    const a = analyze(pkg(sheet('<row r="1"><c r="A1" s="2"><v>1.5</v></c></row>')), 'A1');
    expect(a.format.xfIndex).toBe(2);
    expect(a.format.source).toBe('cell');
    expect(a.format.formatCode).toBe('0.000');
  });

  it('ignores row/@s unless customFormat is set', () => {
    const a = analyze(pkg(sheet('<row r="1" s="3"><c r="A1"><v>1</v></c></row>')), 'A1');
    expect(a.format.source).toBe('default');
    expect(a.format.trace.join(' ')).toContain('ignored because customFormat');
  });

  it('uses row/@s when customFormat is set', () => {
    const a = analyze(pkg(sheet('<row r="1" s="3" customFormat="1"><c r="A1"><v>1</v></c></row>')), 'A1');
    expect(a.format.source).toBe('row');
    expect(a.format.xfIndex).toBe(3);
  });

  it('finds the column definition by span, not by position', () => {
    // col entries carry min/max ranges, so a cell in column 3 is covered by a
    // definition spanning 1-5. Indexing positionally would miss it.
    const ctx = loadExcelContext(pkg(
      sheet('<row r="1"><c r="C1"><v>1</v></c></row>', '<cols><col min="1" max="5" style="2"/></cols>')
    ));
    const sheetDoc = [...ctx.sheets.values()][0];
    const cell = sheetDoc.getElementsByTagNameNS(S_NAMESPACE, 'c').item(0)!;
    // The cell exists, so the column default does not apply - but it was considered.
    expect(analyzeCell(ctx, sheetDoc, cell).format.source).toBe('default');
  });
});

describe('reading values - the t="s" vs t="str" trap', () => {
  const read = (cellXml: string) => {
    const parts = pkg(sheet(`<row r="1">${cellXml}</row>`));
    const ctx = loadExcelContext(parts);
    const sheetDoc = [...ctx.sheets.values()][0];
    const cell = sheetDoc.getElementsByTagNameNS(S_NAMESPACE, 'c').item(0)!;
    return analyzeCell(ctx, sheetDoc, cell).value;
  };

  it('resolves t="s" through the shared string table', () => {
    // <v>0</v> under t="s" is an INDEX, not the text "0".
    expect(read('<c r="A1" t="s"><v>0</v></c>').display).toBe('Hello');
  });

  it('treats t="str" as a literal string, not an index', () => {
    // One character different from t="s", and the opposite meaning.
    expect(read('<c r="A1" t="str"><v>0</v></c>').display).toBe('0');
  });

  it('flags a shared string index that is out of range', () => {
    expect(read('<c r="A1" t="s"><v>99</v></c>').notes.join(' ')).toContain('not present in the shared string table');
  });

  it('reads booleans as TRUE/FALSE', () => {
    expect(read('<c r="A1" t="b"><v>1</v></c>').display).toBe('TRUE');
    expect(read('<c r="A1" t="b"><v>0</v></c>').display).toBe('FALSE');
  });

  it('reads inline strings from is, not v', () => {
    expect(read('<c r="A1" t="inlineStr"><is><t>Inline</t></is></c>').display).toBe('Inline');
  });

  it('defaults to numeric when @t is absent', () => {
    expect(read('<c r="A1"><v>42</v></c>')).toMatchObject({ type: 'n', display: '42' });
  });

  it('flags a formula with no cached result', () => {
    // Libraries routinely write f with no v, and the cell renders blank everywhere
    // except Excel, which recalculates.
    expect(read('<c r="A1"><f>SUM(B1:B2)</f></c>').notes.join(' ')).toContain('no cached result');
  });
});

describe('dates are a number plus a date-shaped format', () => {
  it('interprets a numeric value as a date when the format says so', () => {
    // Nothing in the cell marks it as a date; only numFmtId 14 does.
    const parts = pkg(sheet('<row r="1"><c r="A1" s="1"><v>45292</v></c></row>'));
    const ctx = loadExcelContext(parts);
    const sheetDoc = [...ctx.sheets.values()][0];
    const cell = sheetDoc.getElementsByTagNameNS(S_NAMESPACE, 'c').item(0)!;
    const a = analyzeCell(ctx, sheetDoc, cell);
    expect(a.value.display).toBe('2024-01-01');
    expect(a.value.notes.join(' ')).toContain('1900-compat');
  });

  it('leaves the same number alone under a non-date format', () => {
    const parts = pkg(sheet('<row r="1"><c r="A1" s="2"><v>45292</v></c></row>'));
    const ctx = loadExcelContext(parts);
    const sheetDoc = [...ctx.sheets.values()][0];
    const cell = sheetDoc.getElementsByTagNameNS(S_NAMESPACE, 'c').item(0)!;
    expect(analyzeCell(ctx, sheetDoc, cell).value.display).toBe('45292');
  });
});

describe('locating a cell, and refusing to guess', () => {
  it('matches a cell by its markup', () => {
    const ctx = loadExcelContext(pkg(sheet('<row r="1"><c r="A1" t="s"><v>0</v></c></row>')));
    expect(locateCellByMarkup(ctx, '<c r="A1" t="s"><v>0</v></c>')?.cell.getAttribute('r')).toBe('A1');
  });

  it('tolerates pretty-printing', () => {
    const ctx = loadExcelContext(pkg(sheet('<row r="1"><c r="A1"><v>7</v></c></row>')));
    expect(locateCellByMarkup(ctx, `<c r="A1">
      <v>7</v>
    </c>`)).not.toBeNull();
  });

  it('REFUSES when two cells share identical markup', () => {
    // Spreadsheets are full of identical cells. Picking one and reporting its format
    // as Verified would be a confidently wrong answer.
    const ctx = loadExcelContext(pkg(sheet('<row r="1"><c/><c/></row>')));
    expect(locateCellByMarkup(ctx, '<c/>')).toBeNull();
  });

  it('searches across every worksheet in the package', () => {
    const parts = pkg(sheet('<row r="1"><c r="A1"><v>1</v></c></row>'), {
      'xl/worksheets/sheet2.xml': sheet('<row r="1"><c r="B2"><v>2</v></c></row>')
    });
    const ctx = loadExcelContext(parts);
    expect(locateCellByMarkup(ctx, '<c r="B2"><v>2</v></c>')?.cell.getAttribute('r')).toBe('B2');
  });
});

describe('the full chain', () => {
  it('produces evidence that selects the verified tier', () => {
    const evidence = computeExcelEvidenceForMarkup(
      pkg(sheet('<row r="1"><c r="A1" s="2"><v>1.5</v></c></row>')),
      '<c r="A1" s="2"><v>1.5</v></c>'
    )!;
    expect(evidence.lines.join('\n')).toContain('cellXfs[2]');
    expect(selectEvidenceTier(false, evidence)).toBe('verified');
  });

  it('caps at grounded when a part is missing', () => {
    const parts = pkg(sheet('<row r="1"><c r="A1" s="2"><v>1.5</v></c></row>'));
    delete parts['xl/workbook.xml'];
    const evidence = computeExcelEvidenceForMarkup(parts, '<c r="A1" s="2"><v>1.5</v></c>')!;
    expect(evidence.unresolved.length).toBeGreaterThan(0);
    expect(selectEvidenceTier(false, evidence)).toBe('grounded');
  });

  it('returns null when the cell cannot be located', () => {
    expect(computeExcelEvidenceForMarkup(pkg(sheet('<row r="1"><c r="A1"/></row>')), '<c r="Z9"/>')).toBeNull();
  });

  it('states that the named style is provenance only', () => {
    const evidence = computeExcelEvidenceForMarkup(
      pkg(sheet('<row r="1"><c r="A1" s="0"><v>1</v></c></row>')),
      '<c r="A1" s="0"><v>1</v></c>'
    )!;
    expect(evidence.lines.join('\n')).toContain('NOT merged in');
  });
});
