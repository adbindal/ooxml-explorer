import { describe, it, expect } from 'vitest';
import {
  readTableLook,
  readBandSizes,
  applicableConditionalFormats,
  readConditionalBlocks,
  type TableLook,
  type BandSizes
} from '../services/wordTableStyles';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const frag = (xml: string): Element => {
  const doc = new DOMParser().parseFromString(`<?xml version="1.0"?><root ${W}>${xml}</root>`, 'application/xml');
  return doc.documentElement.firstElementChild!;
};

const look = (over: Partial<TableLook> = {}): TableLook => ({
  firstRow: false, lastRow: false, firstColumn: false, lastColumn: false,
  noHBand: false, noVBand: false, source: 'named-attributes', ...over
});
const bands = (rowBandSize = 1, colBandSize = 1): BandSizes =>
  ({ rowBandSize, colBandSize, notes: [] });

describe('w:tblLook — two incompatible encodings', () => {
  it('reads the named attributes', () => {
    const l = readTableLook(frag('<w:tblPr><w:tblLook w:firstRow="1" w:lastRow="0" w:noVBand="1"/></w:tblPr>'));
    expect(l.firstRow).toBe(true);
    expect(l.lastRow).toBe(false);
    expect(l.noVBand).toBe(true);
    expect(l.source).toBe('named-attributes');
  });

  it('reads the legacy bitmask when no named attribute is present', () => {
    // 0x04A0 = firstRow (0x20) | lastRow? no | firstColumn (0x80) | noVBand (0x400)
    const l = readTableLook(frag('<w:tblPr><w:tblLook w:val="04A0"/></w:tblPr>'));
    expect(l.firstRow).toBe(true);
    expect(l.firstColumn).toBe(true);
    expect(l.noVBand).toBe(true);
    expect(l.source).toBe('val-bitmask');
  });

  it('IGNORES the bitmask when named attributes are present', () => {
    // This is what Word writes, and the reason editing w:val alone does nothing.
    // The mask says firstRow; the named attribute says otherwise and wins.
    const l = readTableLook(frag('<w:tblPr><w:tblLook w:val="04A0" w:firstRow="0" w:noVBand="0"/></w:tblPr>'));
    expect(l.firstRow).toBe(false);
    expect(l.noVBand).toBe(false);
    expect(l.source).toBe('named-attributes');
  });

  it('treats a single named attribute as enough to disable the bitmask', () => {
    const l = readTableLook(frag('<w:tblPr><w:tblLook w:val="04A0" w:lastRow="1"/></w:tblPr>'));
    expect(l.source).toBe('named-attributes');
    expect(l.firstRow).toBe(false);
    expect(l.lastRow).toBe(true);
  });

  it('inverts the banding bits correctly — a set bit SUPPRESSES banding', () => {
    expect(readTableLook(frag('<w:tblPr><w:tblLook w:val="0200"/></w:tblPr>')).noHBand).toBe(true);
    expect(readTableLook(frag('<w:tblPr><w:tblLook w:val="0000"/></w:tblPr>')).noHBand).toBe(false);
  });

  it('reports absence rather than guessing', () => {
    expect(readTableLook(frag('<w:tblPr/>')).source).toBe('absent');
    expect(readTableLook(null).source).toBe('absent');
  });

  it('does not throw on a malformed mask', () => {
    expect(readTableLook(frag('<w:tblPr><w:tblLook w:val="zzz"/></w:tblPr>')).source).toBe('absent');
  });
});

describe('band sizes — Word defaults to 0, not 1', () => {
  it('defaults an absent row band size to 0, disabling banding', () => {
    // The headline deviation: a table style with a correct band1Horz renders with
    // no banding at all until tblStyleRowBandSize is written explicitly.
    const b = readBandSizes(frag('<w:tblPr/>'));
    expect(b.rowBandSize).toBe(0);
    expect(b.notes.join(' ')).toContain('disables row banding');
  });

  it('explains that the specification would have defaulted it to 1', () => {
    expect(readBandSizes(frag('<w:tblPr/>')).notes.join(' ')).toContain('specification would default it to 1');
  });

  it('reads an explicit band size', () => {
    const b = readBandSizes(frag('<w:tblPr><w:tblStyleRowBandSize w:val="2"/><w:tblStyleColBandSize w:val="3"/></w:tblPr>'));
    expect(b.rowBandSize).toBe(2);
    expect(b.colBandSize).toBe(3);
    expect(b.notes).toEqual([]);
  });

  it('produces no banding at all when the size is 0', () => {
    const applied = applicableConditionalFormats(look(), bands(0, 0), { row: 1, col: 1, rowCount: 4, colCount: 4 });
    expect(applied).toEqual(['wholeTable']);
  });
});

describe('Word application order, which differs from the specification', () => {
  it('applies row banding BEFORE column banding, so column banding wins', () => {
    // The specification orders column banding first. Where both apply, the later
    // entry overrides, so this ordering decides which shading a user actually sees.
    const applied = applicableConditionalFormats(
      look(), bands(1, 1), { row: 1, col: 1, rowCount: 5, colCount: 5 }
    );
    expect(applied.indexOf('band2Horz')).toBeLessThan(applied.indexOf('band2Vert'));
  });

  it('applies first/last column BEFORE first/last row, so row formatting wins', () => {
    const applied = applicableConditionalFormats(
      look({ firstRow: true, firstColumn: true }), bands(0, 0),
      { row: 0, col: 0, rowCount: 3, colCount: 3 }
    );
    expect(applied.indexOf('firstCol')).toBeLessThan(applied.indexOf('firstRow'));
  });

  it('puts corner cells last, so they win outright', () => {
    const applied = applicableConditionalFormats(
      look({ firstRow: true, firstColumn: true }), bands(0, 0),
      { row: 0, col: 0, rowCount: 3, colCount: 3 }
    );
    expect(applied[applied.length - 1]).toBe('nwCell');
  });

  it('always starts from wholeTable', () => {
    expect(applicableConditionalFormats(look(), bands(0, 0), { row: 1, col: 1, rowCount: 3, colCount: 3 })[0])
      .toBe('wholeTable');
  });
});

describe('which regions apply', () => {
  it('applies firstRow only when the look enables it', () => {
    const pos = { row: 0, col: 1, rowCount: 3, colCount: 3 };
    expect(applicableConditionalFormats(look({ firstRow: true }), bands(0, 0), pos)).toContain('firstRow');
    expect(applicableConditionalFormats(look({ firstRow: false }), bands(0, 0), pos)).not.toContain('firstRow');
  });

  it('applies lastRow to the final row only', () => {
    const l = look({ lastRow: true });
    expect(applicableConditionalFormats(l, bands(0, 0), { row: 2, col: 0, rowCount: 3, colCount: 3 })).toContain('lastRow');
    expect(applicableConditionalFormats(l, bands(0, 0), { row: 1, col: 0, rowCount: 3, colCount: 3 })).not.toContain('lastRow');
  });

  it('suppresses banding when noHBand is set', () => {
    const applied = applicableConditionalFormats(look({ noHBand: true }), bands(1, 0), { row: 1, col: 0, rowCount: 4, colCount: 4 });
    expect(applied.some(t => t.endsWith('Horz'))).toBe(false);
  });

  it('excludes the header row from banding when firstRow is enabled', () => {
    // Banding counts from after the header, so row 0 is not a band member.
    const applied = applicableConditionalFormats(look({ firstRow: true }), bands(1, 0), { row: 0, col: 0, rowCount: 4, colCount: 4 });
    expect(applied.some(t => t.endsWith('Horz'))).toBe(false);
  });

  it('alternates bands with a band size greater than one', () => {
    const l = look(), b = bands(2, 0);
    const bandOf = (row: number) =>
      applicableConditionalFormats(l, b, { row, col: 0, rowCount: 8, colCount: 3 }).find(t => t.endsWith('Horz'));
    expect(bandOf(0)).toBe('band1Horz');
    expect(bandOf(1)).toBe('band1Horz');
    expect(bandOf(2)).toBe('band2Horz');
    expect(bandOf(3)).toBe('band2Horz');
    expect(bandOf(4)).toBe('band1Horz');
  });

  it('identifies all four corners', () => {
    const l = look({ firstRow: true, lastRow: true, firstColumn: true, lastColumn: true });
    const at = (row: number, col: number) =>
      applicableConditionalFormats(l, bands(0, 0), { row, col, rowCount: 3, colCount: 3 });
    expect(at(0, 0)).toContain('nwCell');
    expect(at(0, 2)).toContain('neCell');
    expect(at(2, 0)).toContain('swCell');
    expect(at(2, 2)).toContain('seCell');
  });
});

describe('reading conditional blocks from a style', () => {
  it('keys tblStylePr blocks by their type', () => {
    const style = frag(`<w:style w:type="table" w:styleId="Grid">
      <w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr>
      <w:tblStylePr w:type="band1Horz"><w:tcPr/></w:tblStylePr>
    </w:style>`);
    const blocks = readConditionalBlocks(style);
    expect([...blocks.keys()].sort()).toEqual(['band1Horz', 'firstRow']);
  });

  it('returns an empty map for a style with no conditional formatting', () => {
    expect(readConditionalBlocks(frag('<w:style w:styleId="Plain"/>')).size).toBe(0);
  });
});
