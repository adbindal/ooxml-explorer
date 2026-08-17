/**
 * Composes the SpreadsheetML format resolution over a whole package.
 *
 * The counterpart to `wordFormattingAnalysis`, and the same job: take the parts of an
 * `.xlsx`, take a cell, and answer what Excel will show and why — by computation, so
 * the result can be presented as verified rather than merely grounded.
 *
 * The shape of the problem is different from Word's, though. There is no cascade to
 * assemble; the work is finding which single `xf` applies (which needs the cell's row
 * and column context, not just the cell), and then interpreting the stored value, which
 * needs the shared string table and the workbook's date system. Getting the value wrong
 * is easier than getting the format wrong: `t="s"` means `<v>` is an *index*, and
 * reading it as text is silent, plausible corruption.
 */

import {
  parseExcelStyles,
  resolveCellFormat,
  readDateSystem,
  serialToDate,
  S_NAMESPACE,
  type ExcelStyleSheet,
  type DateSystem,
  type ResolvedCellFormat
} from './excelStyleResolver';
import type { PackageParts } from './packageIntegrity';

const STYLES_PART = 'xl/styles.xml';
const WORKBOOK_PART = 'xl/workbook.xml';
const SHARED_STRINGS_PART = 'xl/sharedStrings.xml';

export interface ExcelDocumentContext {
  styles: ExcelStyleSheet;
  dateSystem: DateSystem;
  /** Shared strings by index. Empty when the part is absent, which is legal. */
  sharedStrings: string[];
  /** Parsed worksheet parts, keyed by part path. */
  sheets: Map<string, Document>;
  unresolved: string[];
}

const parseXml = (xml: string): Document | null => {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
};

/**
 * Reads the shared string table.
 *
 * An `si` may hold either a plain `<t>` or a sequence of `<r>` runs with their own
 * formatting. The runs are concatenated: the same visible string can legitimately be
 * one run or twenty, exactly as with Word's `w:r`, so treating run boundaries as
 * meaningful would make identical text compare unequal.
 */
const readSharedStrings = (doc: Document): string[] =>
  Array.from(doc.getElementsByTagNameNS(S_NAMESPACE, 'si')).map(si =>
    Array.from(si.getElementsByTagNameNS(S_NAMESPACE, 't'))
      .map(t => t.textContent ?? '')
      .join('')
  );

export const loadExcelContext = (parts: PackageParts): ExcelDocumentContext => {
  const unresolved: string[] = [];

  const stylesXml = parts[STYLES_PART];
  if (stylesXml === undefined) {
    unresolved.push(`${STYLES_PART} is not in the package; only the default format can be reported`);
  }
  const styles = stylesXml
    ? parseExcelStyles(stylesXml)
    : { numFmts: new Map(), fonts: [], fills: [], borders: [], cellStyleXfs: [], cellXfs: [], cellStyles: [] };

  const workbookXml = parts[WORKBOOK_PART];
  if (workbookXml === undefined) {
    unresolved.push(`${WORKBOOK_PART} is not in the package; the date system cannot be determined and the default is assumed`);
  }
  const dateSystem = workbookXml ? readDateSystem(workbookXml) : '1900-compat';

  const sharedXml = parts[SHARED_STRINGS_PART];
  let sharedStrings: string[] = [];
  if (sharedXml !== undefined) {
    const doc = parseXml(sharedXml);
    if (doc) sharedStrings = readSharedStrings(doc);
    else unresolved.push(`${SHARED_STRINGS_PART} is not well-formed XML; shared string values cannot be read`);
  }

  const sheets = new Map<string, Document>();
  for (const [path, content] of Object.entries(parts)) {
    if (!/^xl\/worksheets\/[^/]+\.xml$/.test(path)) continue;
    const doc = parseXml(content);
    if (doc) sheets.set(path, doc);
    else unresolved.push(`${path} is not well-formed XML`);
  }
  if (sheets.size === 0) {
    unresolved.push('no worksheet parts were found in the package');
  }

  return { styles, dateSystem, sharedStrings, sheets, unresolved };
};

/** `"BC12"` -> `{ column: 55, row: 12 }`, both 1-based. Null when unparseable. */
export const parseCellReference = (ref: string): { column: number; row: number } | null => {
  const match = /^([A-Za-z]+)(\d+)$/.exec(ref.trim());
  if (!match) return null;
  const letters = match[1].toUpperCase();
  let column = 0;
  for (const ch of letters) column = column * 26 + (ch.charCodeAt(0) - 64);
  return { column, row: Number.parseInt(match[2], 10) };
};

const intAttr = (el: Element | null, name: string): number | null => {
  const raw = el?.getAttribute(name);
  if (raw === null || raw === undefined) return null;
  const parsed = Number.parseInt(raw, 10);
  return Number.isNaN(parsed) ? null : parsed;
};

const isTruthy = (raw: string | null | undefined): boolean =>
  raw !== null && raw !== undefined && !['0', 'false'].includes(raw.trim().toLowerCase());

/**
 * Finds the `col` definition covering a column index.
 *
 * `col` entries carry `min`/`max` spans rather than one entry per column, so a linear
 * scan for a containing range is required — indexing by position would be wrong.
 */
const columnDefinitionFor = (sheet: Document, column: number): Element | null =>
  Array.from(sheet.getElementsByTagNameNS(S_NAMESPACE, 'col')).find(col => {
    const min = intAttr(col, 'min');
    const max = intAttr(col, 'max');
    return min !== null && max !== null && column >= min && column <= max;
  }) ?? null;

export interface CellValueReading {
  /** Raw `<v>` text, or the inline string when `t="inlineStr"`. */
  raw: string | null;
  /** What the value means once `@t` is honoured. */
  display: string | null;
  /** `@t`, defaulted to `n`. */
  type: string;
  notes: string[];
}

/**
 * Interprets a cell's stored value.
 *
 * The single most common misread in the format is `t="s"`: `<v>5</v>` under it is a
 * zero-based index into the shared string table, not the text "5". `t="str"` — one
 * character different — means `<v>` *is* the string, because the cell holds a formula
 * whose result was text. Confusing them is silent and plausible.
 */
export const readCellValue = (
  cell: Element,
  context: ExcelDocumentContext,
  format: ResolvedCellFormat
): CellValueReading => {
  const notes: string[] = [];
  const type = cell.getAttribute('t') ?? 'n';

  const isEl = Array.from(cell.children).find(
    el => el.namespaceURI === S_NAMESPACE && el.localName === 'is'
  );
  const vEl = Array.from(cell.children).find(
    el => el.namespaceURI === S_NAMESPACE && el.localName === 'v'
  );

  if (type === 'inlineStr') {
    const text = isEl
      ? Array.from(isEl.getElementsByTagNameNS(S_NAMESPACE, 't')).map(t => t.textContent ?? '').join('')
      : null;
    return { raw: text, display: text, type, notes };
  }

  const raw = vEl?.textContent ?? null;
  if (raw === null) {
    const hasFormula = Array.from(cell.children).some(
      el => el.namespaceURI === S_NAMESPACE && el.localName === 'f'
    );
    if (hasFormula) {
      notes.push('the cell holds a formula with no cached result; Excel will show nothing until it recalculates');
    }
    return { raw: null, display: null, type, notes };
  }

  switch (type) {
    case 's': {
      const index = Number.parseInt(raw, 10);
      const text = Number.isNaN(index) ? null : context.sharedStrings[index] ?? null;
      if (text === null) {
        notes.push(`shared string index ${raw} is not present in the shared string table`);
      }
      return { raw, display: text, type, notes };
    }
    case 'b':
      return { raw, display: raw === '1' ? 'TRUE' : 'FALSE', type, notes };
    case 'e':
      return { raw, display: raw, type, notes };
    case 'str':
      // Deliberately distinct from 's': here the value IS the string.
      return { raw, display: raw, type, notes };
    default: {
      const numeric = Number.parseFloat(raw);
      if (Number.isNaN(numeric)) return { raw, display: raw, type, notes };

      // A date is a number plus a date-shaped format; nothing in the cell says so.
      const looksLikeDate = /[dmyhs]/i.test(format.formatCode ?? '') && !/^(General|@)$/.test(format.formatCode ?? '');
      if (looksLikeDate) {
        const date = serialToDate(numeric, context.dateSystem);
        if (date) {
          notes.push(`interpreted as a date using the ${context.dateSystem} system, because the number format is date-shaped`);
          return { raw, display: date.toISOString().slice(0, 10), type, notes };
        }
      }
      return { raw, display: raw, type, notes };
    }
  }
};

export interface CellAnalysis {
  reference: string | null;
  format: ResolvedCellFormat;
  value: CellValueReading;
  unresolved: string[];
  explanation: string[];
}

/** Resolves the format and value of one `c` element. */
export const analyzeCell = (
  context: ExcelDocumentContext,
  sheet: Document,
  cell: Element
): CellAnalysis => {
  const unresolved = [...context.unresolved];
  const reference = cell.getAttribute('r');
  const address = reference ? parseCellReference(reference) : null;

  const row = cell.parentElement;
  const isRow = row?.namespaceURI === S_NAMESPACE && row.localName === 'row';
  if (!isRow) {
    unresolved.push('the cell is not inside a row element, so row-level formatting could not be considered');
  }

  const columnDefinition = address ? columnDefinitionFor(sheet, address.column) : null;
  if (!address && reference) {
    unresolved.push(`cell reference "${reference}" could not be parsed, so column-level formatting was not considered`);
  }

  const format = resolveCellFormat(context.styles, {
    cellStyleIndex: intAttr(cell, 's'),
    rowStyleIndex: isRow ? intAttr(row, 's') : null,
    rowCustomFormat: isRow ? isTruthy(row.getAttribute('customFormat')) : false,
    columnStyleIndex: intAttr(columnDefinition, 'style'),
    cellExists: true
  });

  const value = readCellValue(cell, context, format);
  unresolved.push(...value.notes.filter(n => n.includes('not present') || n.includes('no cached result')));

  const explanation: string[] = [
    `Cell ${reference ?? '(no reference)'}:`,
    `  format record: cellXfs[${format.xfIndex}], chosen from the ${format.source}`,
    ...format.trace.map(t => `  ${t}`),
    `  number format: ${format.formatCode ?? '(unresolved)'}${format.isBuiltIn ? ' (built-in, implied by numFmtId)' : ''}`,
    `  stored type: ${value.type}, raw value: ${value.raw ?? '(none)'}`,
    `  displays as: ${value.display ?? '(nothing)'}`
  ];
  if (format.namedStyle) {
    explanation.push(`  built from named style "${format.namedStyle}" (recorded as provenance only)`);
  }
  for (const note of [...format.notes, ...value.notes]) explanation.push(`  note: ${note}`);
  if (unresolved.length > 0) {
    explanation.push('Not established by this analysis (do not assert these):', ...unresolved.map(u => `  ${u}`));
  }

  return { reference, format, value, unresolved, explanation };
};

const normalizeMarkup = (xml: string): string =>
  xml
    .replace(/\s+xmlns(:[A-Za-z0-9_-]+)?="[^"]*"/g, '')
    .replace(/>\s+</g, '><')
    .replace(/\s+/g, ' ')
    .trim();

/**
 * Locates a cell from a snippet of markup, across every worksheet in the package.
 *
 * Refuses to guess, on the same principle as the Word locator: a spreadsheet is full of
 * cells with identical markup, and resolving the wrong one would be a confidently wrong
 * answer under a "Verified" badge.
 */
export const locateCellByMarkup = (
  context: ExcelDocumentContext,
  rawXml: string
): { sheet: Document; cell: Element } | null => {
  const needle = normalizeMarkup(rawXml);
  if (needle === '') return null;

  const matches: { sheet: Document; cell: Element }[] = [];
  for (const sheet of context.sheets.values()) {
    for (const cell of Array.from(sheet.getElementsByTagNameNS(S_NAMESPACE, 'c'))) {
      if (normalizeMarkup(new XMLSerializer().serializeToString(cell)) === needle) {
        matches.push({ sheet, cell });
      }
    }
  }
  return matches.length === 1 ? matches[0] : null;
};

/**
 * One-call entry point mirroring the Word side: package parts plus a snippet, out comes
 * evidence ready for the AI layer, or null when nothing could be computed.
 */
export const computeExcelEvidenceForMarkup = (
  parts: PackageParts,
  rawXml: string
): { lines: string[]; unresolved: string[] } | null => {
  const context = loadExcelContext(parts);
  const located = locateCellByMarkup(context, rawXml);
  if (!located) return null;

  const analysis = analyzeCell(context, located.sheet, located.cell);
  return { lines: analysis.explanation, unresolved: analysis.unresolved };
};
