/**
 * WordprocessingML table style conditional formatting — cascade layer 2.
 *
 * A table style can define formatting for thirteen conditional regions (first row,
 * banded columns, corner cells, …) but which of them actually apply is decided by
 * `w:tblLook` **on the individual table**, not by the style. Getting that gate wrong
 * produces a table that looks nothing like the style says it should, with no error.
 *
 * This layer carries an unusual density of documented Word-versus-specification
 * divergence, and all three of the ones encoded here are silent:
 *
 *   1. `w:tblStyleRowBandSize` — the specification says an omitted value defaults to 1.
 *      Word defaults it to **0**, and 0 means "apply no banded-row formatting at all".
 *      A table style with a perfectly correct `band1Horz` therefore renders with no
 *      banding until the element is written explicitly. This is reportedly the single
 *      most frequent "my table style doesn't work" cause.
 *
 *   2. Application order. The specification orders conditional formats
 *      whole-table → column banding → row banding → first/last row → first/last column
 *      → corners. Word applies them
 *      whole-table → row banding → **column banding** → first/last **column** →
 *      first/last **row** → corners. So where both apply, column banding wins over row
 *      banding in Word and loses in a specification-faithful renderer, and first-row
 *      formatting beats first-column.
 *
 *   3. `w:tblLook` has two incompatible encodings. ECMA-376 1st edition used a single
 *      hex bitmask in `@w:val`; ISO 29500 added named attributes. Word writes both, and
 *      reads `@w:val` **only if none of the named attributes are present**. Editing the
 *      bitmask on a table that also carries named attributes is a silent no-op.
 */

import { W_NAMESPACE } from './wordStyleResolver';

/** The thirteen conditional regions a table style may define. */
export type ConditionalFormatType =
  | 'wholeTable'
  | 'band1Horz' | 'band2Horz'
  | 'band1Vert' | 'band2Vert'
  | 'firstCol' | 'lastCol'
  | 'firstRow' | 'lastRow'
  | 'nwCell' | 'neCell' | 'swCell' | 'seCell';

/** Bit positions in the legacy `w:tblLook/@w:val` mask. */
const LOOK_BITS = {
  firstRow: 0x0020,
  lastRow: 0x0040,
  firstColumn: 0x0080,
  lastColumn: 0x0100,
  /** Note the inversion: a set bit *disables* banding. */
  noHBand: 0x0200,
  noVBand: 0x0400
} as const;

export interface TableLook {
  firstRow: boolean;
  lastRow: boolean;
  firstColumn: boolean;
  lastColumn: boolean;
  /** True means row banding is *suppressed*. */
  noHBand: boolean;
  /** True means column banding is *suppressed*. */
  noVBand: boolean;
  /** Which encoding actually decided the result. */
  source: 'named-attributes' | 'val-bitmask' | 'absent';
}

const NAMED_LOOK_ATTRIBUTES = ['firstRow', 'lastRow', 'firstColumn', 'lastColumn', 'noHBand', 'noVBand'] as const;

const isTruthy = (raw: string | null): boolean =>
  raw !== null && !['0', 'false', 'off'].includes(raw.trim().toLowerCase());

/**
 * Reads `w:tblLook`, honouring Word's rule that the legacy bitmask is consulted only
 * when no named attribute is present.
 *
 * A table carrying both — which is what Word itself writes — is decided entirely by
 * the named attributes, so changing `@w:val` alone does nothing.
 */
export const readTableLook = (tblPr: Element | null | undefined): TableLook => {
  const base: TableLook = {
    firstRow: false, lastRow: false, firstColumn: false, lastColumn: false,
    noHBand: false, noVBand: false, source: 'absent'
  };
  if (!tblPr) return base;

  const look = Array.from(tblPr.children).find(
    el => el.namespaceURI === W_NAMESPACE && el.localName === 'tblLook'
  );
  if (!look) return base;

  const present = NAMED_LOOK_ATTRIBUTES.filter(
    name => look.getAttributeNS(W_NAMESPACE, name) !== null
  );

  if (present.length > 0) {
    const read = (name: typeof NAMED_LOOK_ATTRIBUTES[number]) =>
      isTruthy(look.getAttributeNS(W_NAMESPACE, name));
    return {
      firstRow: read('firstRow'),
      lastRow: read('lastRow'),
      firstColumn: read('firstColumn'),
      lastColumn: read('lastColumn'),
      noHBand: read('noHBand'),
      noVBand: read('noVBand'),
      source: 'named-attributes'
    };
  }

  const raw = look.getAttributeNS(W_NAMESPACE, 'val');
  if (raw === null) return base;

  const mask = Number.parseInt(raw, 16);
  if (Number.isNaN(mask)) return base;

  return {
    firstRow: (mask & LOOK_BITS.firstRow) !== 0,
    lastRow: (mask & LOOK_BITS.lastRow) !== 0,
    firstColumn: (mask & LOOK_BITS.firstColumn) !== 0,
    lastColumn: (mask & LOOK_BITS.lastColumn) !== 0,
    noHBand: (mask & LOOK_BITS.noHBand) !== 0,
    noVBand: (mask & LOOK_BITS.noVBand) !== 0,
    source: 'val-bitmask'
  };
};

export interface BandSizes {
  /** Rows per horizontal band. 0 disables row banding entirely. */
  rowBandSize: number;
  /** Columns per vertical band. 0 disables column banding entirely. */
  colBandSize: number;
  /** Set when a value was absent and Word's default (0) was applied. */
  notes: string[];
}

/**
 * Reads band sizes from a table style, applying **Word's** defaults rather than the
 * specification's.
 *
 * The specification says an omitted `w:tblStyleRowBandSize` defaults to 1. Word treats
 * it as 0, which switches banding off. Encoding the specification's default here would
 * make the resolver claim banding that Word does not draw.
 */
export const readBandSizes = (styleTblPr: Element | null | undefined): BandSizes => {
  const notes: string[] = [];
  const read = (localName: string, label: string): number => {
    const el = styleTblPr
      ? Array.from(styleTblPr.children).find(
          c => c.namespaceURI === W_NAMESPACE && c.localName === localName
        )
      : undefined;
    const raw = el?.getAttributeNS(W_NAMESPACE, 'val');
    if (raw === null || raw === undefined) {
      notes.push(`${localName} is absent; Word treats it as 0, which disables ${label} banding (the specification would default it to 1)`);
      return 0;
    }
    const parsed = Number.parseInt(raw, 10);
    return Number.isNaN(parsed) ? 0 : parsed;
  };

  return {
    rowBandSize: read('tblStyleRowBandSize', 'row'),
    colBandSize: read('tblStyleColBandSize', 'column'),
    notes
  };
};

export interface CellPosition {
  /** 0-based row index within the table. */
  row: number;
  /** 0-based column index within the row. */
  col: number;
  rowCount: number;
  colCount: number;
}

/**
 * Returns the conditional formats that apply to a cell, **in Word's application
 * order** — earlier entries are overridden by later ones.
 *
 * Word's order is not the specification's; see the module comment. The difference is
 * observable whenever a cell sits in both a banded row and a banded column, or in both
 * a first row and a first column.
 */
export const applicableConditionalFormats = (
  look: TableLook,
  bands: BandSizes,
  pos: CellPosition
): ConditionalFormatType[] => {
  const applied: ConditionalFormatType[] = ['wholeTable'];

  // Banding indices are counted from after the header row / first column when those
  // are enabled, which is why the offsets below depend on the look.
  const bandRowIndex = pos.row - (look.firstRow ? 1 : 0);
  const bandColIndex = pos.col - (look.firstColumn ? 1 : 0);

  const isFirstRow = look.firstRow && pos.row === 0;
  const isLastRow = look.lastRow && pos.row === pos.rowCount - 1;
  const isFirstCol = look.firstColumn && pos.col === 0;
  const isLastCol = look.lastColumn && pos.col === pos.colCount - 1;

  // 1. Row banding — before column banding, unlike the specification.
  if (!look.noHBand && bands.rowBandSize > 0 && !isFirstRow && !isLastRow && bandRowIndex >= 0) {
    const band = Math.floor(bandRowIndex / bands.rowBandSize);
    applied.push(band % 2 === 0 ? 'band1Horz' : 'band2Horz');
  }

  // 2. Column banding.
  if (!look.noVBand && bands.colBandSize > 0 && !isFirstCol && !isLastCol && bandColIndex >= 0) {
    const band = Math.floor(bandColIndex / bands.colBandSize);
    applied.push(band % 2 === 0 ? 'band1Vert' : 'band2Vert');
  }

  // 3. First/last column — before first/last row, unlike the specification.
  if (isFirstCol) applied.push('firstCol');
  if (isLastCol) applied.push('lastCol');

  // 4. First/last row.
  if (isFirstRow) applied.push('firstRow');
  if (isLastRow) applied.push('lastRow');

  // 5. Corner cells. Word styles these individually where a row and a column
  //    condition meet; the specification does not describe the case.
  if (isFirstRow && isFirstCol) applied.push('nwCell');
  if (isFirstRow && isLastCol) applied.push('neCell');
  if (isLastRow && isFirstCol) applied.push('swCell');
  if (isLastRow && isLastCol) applied.push('seCell');

  return applied;
};

/**
 * Collects a table style's `w:tblStylePr` blocks, keyed by conditional type.
 */
export const readConditionalBlocks = (
  styleElement: Element | null | undefined
): Map<ConditionalFormatType, Element> => {
  const blocks = new Map<ConditionalFormatType, Element>();
  if (!styleElement) return blocks;
  for (const el of Array.from(styleElement.children)) {
    if (el.namespaceURI !== W_NAMESPACE || el.localName !== 'tblStylePr') continue;
    const type = el.getAttributeNS(W_NAMESPACE, 'type') as ConditionalFormatType | null;
    if (type) blocks.set(type, el);
  }
  return blocks;
};
