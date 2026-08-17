import { describe, it, expect } from 'vitest';
import {
  parseExcelStyles,
  selectXfIndex,
  resolveCellFormat,
  serialToDate,
  readDateSystem,
  BUILT_IN_NUMBER_FORMATS,
  DATE_SYSTEM_OFFSET_DAYS,
  S_NAMESPACE
} from '../services/excelStyleResolver';

const styles = (body: string) =>
  parseExcelStyles(`<?xml version="1.0"?><styleSheet xmlns="${S_NAMESPACE}">${body}</styleSheet>`);

const SHEET = styles(`
  <numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>
  <fonts count="2"><font><sz val="11"/></font><font><b/><sz val="14"/></font></fonts>
  <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"/></fill></fills>
  <borders count="2"><border/><border><left style="thin"/></border></borders>
  <cellStyleXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" applyFont="1"/>
  </cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="1" applyFont="0"/>
    <xf numFmtId="999" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="2">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
    <cellStyle name="Heading 1" xfId="1" builtinId="16"/>
  </cellStyles>`);

describe('parsing styles.xml', () => {
  it('reads the positional component tables', () => {
    expect(SHEET.fonts.length).toBe(2);
    expect(SHEET.fills.length).toBe(3);
    expect(SHEET.borders.length).toBe(2);
    expect(SHEET.cellXfs.length).toBe(4);
  });

  it('keys numFmts by the explicit numFmtId, not by position', () => {
    // Every other component table is positional; this one is not. Indexing it by
    // position is a real and silent bug.
    expect(SHEET.numFmts.get(164)).toBe('0.000');
    expect(SHEET.numFmts.get(0)).toBeUndefined();
  });

  it('reads named styles with their xfId join key', () => {
    expect(SHEET.cellStyles).toContainEqual({ name: 'Heading 1', xfId: 1, builtinId: 16 });
  });
});

describe('choosing the xf - fallback, never merge', () => {
  it('uses the cell index when the cell carries one', () => {
    expect(selectXfIndex({ cellStyleIndex: 2, rowStyleIndex: 1, rowCustomFormat: true }).source).toBe('cell');
  });

  it('ignores row/@s unless customFormat is set', () => {
    // The specification is explicit that row/@s applies "only if customFormat is 1".
    // Applying it unconditionally formats every cell in every row that merely
    // records an index.
    const without = selectXfIndex({ rowStyleIndex: 1, rowCustomFormat: false, cellExists: false });
    expect(without.source).toBe('default');
    expect(without.trace.join(' ')).toContain('ignored because customFormat');
  });

  it('uses row/@s when customFormat is set', () => {
    const r = selectXfIndex({ rowStyleIndex: 1, rowCustomFormat: true, cellExists: false });
    expect(r).toMatchObject({ index: 1, source: 'row' });
  });

  it('applies col/@style only to cells that do not exist yet', () => {
    expect(selectXfIndex({ columnStyleIndex: 2, cellExists: false }).source).toBe('column');
    expect(selectXfIndex({ columnStyleIndex: 2, cellExists: true }).source).toBe('default');
  });

  it('falls back to cellXfs[0]', () => {
    expect(selectXfIndex({ cellExists: false })).toMatchObject({ index: 0, source: 'default' });
  });
});

describe('xfId is provenance, not inheritance', () => {
  // The headline deviation. ECMA says both records must be read; MS-OI29500 says only
  // the cellXfs record defines the formatting. A resolver that merges cellStyleXfs
  // underneath would disagree with Excel on nearly every workbook.
  it('names the style the record came from', () => {
    expect(resolveCellFormat(SHEET, { cellStyleIndex: 2 }).namedStyle).toBe('Heading 1');
  });

  it('says explicitly that the named style is not merged in', () => {
    const r = resolveCellFormat(SHEET, { cellStyleIndex: 2 });
    expect(r.trace.join(' ')).toContain('NOT merged in');
  });

  it('takes every component from cellXfs, never from cellStyleXfs', () => {
    // cellXfs[2] and cellStyleXfs[1] happen to agree here; what matters is that the
    // resolved font is the one cellXfs names.
    const r = resolveCellFormat(SHEET, { cellStyleIndex: 2 });
    expect(r.font).toBe(SHEET.fonts[1]);
    expect(r.fill).toBe(SHEET.fills[2]);
    expect(r.border).toBe(SHEET.borders[1]);
  });
});

describe('apply* flags are not render gates', () => {
  it('does not let applyFont="0" suppress the record\'s own font', () => {
    // cellXfs[2] carries applyFont="0". Read as a render gate that would mean "ignore
    // fontId", which is wrong - it means "propagate later changes to the named style".
    const r = resolveCellFormat(SHEET, { cellStyleIndex: 2 });
    expect(r.font).toBe(SHEET.fonts[1]);
  });

  it('explains what the flags actually do', () => {
    const r = resolveCellFormat(SHEET, { cellStyleIndex: 2 });
    expect(r.notes.join(' ')).toContain('do not gate rendering');
  });
});

describe('number formats', () => {
  it('resolves a custom format by its numFmtId value', () => {
    const r = resolveCellFormat(SHEET, { cellStyleIndex: 2 });
    expect(r.formatCode).toBe('0.000');
    expect(r.isBuiltIn).toBe(false);
  });

  it('reports Excel\'s built-in code rather than the standard\'s by default', () => {
    // numFmtId 14 is the most common date format in existence, and the two readings
    // differ on two-digit versus four-digit year.
    const r = resolveCellFormat(SHEET, { cellStyleIndex: 1 });
    expect(r.formatCode).toBe('m/d/yyyy');
    expect(r.isBuiltIn).toBe(true);
  });

  it('can report the standard\'s reading when asked', () => {
    const r = resolveCellFormat(SHEET, { cellStyleIndex: 1 }, { preferExcelBehaviour: false });
    expect(r.formatCode).toBe('mm-dd-yy');
  });

  it('flags that the two readings disagree', () => {
    expect(resolveCellFormat(SHEET, { cellStyleIndex: 1 }).notes.join(' '))
      .toContain('Excel differs from the standard');
  });

  it('records the known gaps in the built-in table', () => {
    // 5-8 and 23-36 are absent from the All Languages table; treating them as
    // built-in would invent a format.
    for (const id of [5, 6, 7, 8, 23, 36]) {
      expect(BUILT_IN_NUMBER_FORMATS.has(id)).toBe(false);
    }
  });

  it('flags a numFmtId that is neither custom nor built-in', () => {
    expect(resolveCellFormat(SHEET, { cellStyleIndex: 3 }).notes.join(' '))
      .toContain('implementation-defined');
  });
});

describe('out-of-range references', () => {
  it('reports a missing cellXfs entry rather than throwing', () => {
    expect(resolveCellFormat(SHEET, { cellStyleIndex: 99 }).notes.join(' '))
      .toContain('no entry at index 99');
  });

  it('reports an out-of-range component index', () => {
    const broken = styles(`
      <fonts count="1"><font/></fonts>
      <cellXfs count="1"><xf numFmtId="0" fontId="7" fillId="0" borderId="0"/></cellXfs>`);
    expect(resolveCellFormat(broken, { cellStyleIndex: 0 }).notes.join(' '))
      .toContain('font index 7 is out of range');
  });
});

describe('date systems', () => {
  it('has the offset the standard\'s own figures imply', () => {
    expect(DATE_SYSTEM_OFFSET_DAYS).toBe(1462);
    expect(695055 - 693593).toBe(DATE_SYSTEM_OFFSET_DAYS);
    expect(2958465 - 2957003).toBe(DATE_SYSTEM_OFFSET_DAYS);
  });

  it('maps serial 1 to 1900-01-01 in the compatibility system', () => {
    expect(serialToDate(1, '1900-compat')!.toISOString().slice(0, 10)).toBe('1900-01-01');
  });

  it('agrees with the true 1900 base above the phantom day', () => {
    // Serial 61 onwards, the compat system and a 1899-12-30 base coincide. That is
    // why the true system's epoch looks "wrong by two days".
    expect(serialToDate(61, '1900-compat')!.toISOString().slice(0, 10))
      .toBe(serialToDate(61, '1900')!.toISOString().slice(0, 10));
  });

  it('places a familiar date correctly', () => {
    // 2024-01-01 is serial 45292 in the compatibility system Excel actually uses.
    expect(serialToDate(45292, '1900-compat')!.toISOString().slice(0, 10)).toBe('2024-01-01');
  });

  it('maps serial 0 to the 1904 epoch', () => {
    expect(serialToDate(0, '1904')!.toISOString().slice(0, 10)).toBe('1904-01-01');
  });

  it('differs between the two systems by exactly the offset', () => {
    const a = serialToDate(1000, '1900')!.getTime();
    const b = serialToDate(1000, '1904')!.getTime();
    expect((b - a) / 86400000).toBe(DATE_SYSTEM_OFFSET_DAYS);
  });

  it('refuses negative serials, which Excel does not support', () => {
    expect(serialToDate(-1)).toBeNull();
  });
});

describe('reading the workbook date system', () => {
  const workbook = (attrs: string) =>
    `<?xml version="1.0"?><workbook xmlns="${S_NAMESPACE}" ${attrs}><workbookPr date1904="0"/></workbook>`;

  it('defaults Transitional files to the compatibility system', () => {
    // Excel writes Transitional by default, so real files carry the phantom day.
    expect(readDateSystem(workbook(''))).toBe('1900-compat');
  });

  it('gives Strict files the true 1900 system', () => {
    expect(readDateSystem(workbook('conformance="strict"'))).toBe('1900');
  });

  it('honours date1904', () => {
    expect(readDateSystem(
      `<?xml version="1.0"?><workbook xmlns="${S_NAMESPACE}"><workbookPr date1904="1"/></workbook>`
    )).toBe('1904');
  });
});
