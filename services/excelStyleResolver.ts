/**
 * SpreadsheetML cell format resolution.
 *
 * **This is deliberately not a cascade, and that is the whole point of the module.**
 *
 * ECMA-376 §18.8.9 and §18.8.10 both say that "both the cell style xf records and cell
 * xf records shall be read to understand the full set of formatting applied to a cell",
 * which reads exactly like Word's style hierarchy and invites the same implementation.
 * Microsoft's own [MS-OI29500] contradicts it, in identical wording on both clause
 * pages:
 *
 *   > "In Office, only the cell xf record defines the formatting applied to a cell."
 *
 * So `cellXfs[c/@s]` is **complete and self-contained**. `xfId` records which named
 * style the cell was built from — provenance, not an inheritance pointer — and nothing
 * merges into the record at render time. A resolver written faithfully from the
 * standard would disagree with Excel on nearly every workbook, and would look correct
 * in review.
 *
 * The second trap in the same area: the `apply*` flags are not render gates.
 * `applyFont="0"` does **not** mean "ignore this record's fontId". In `cellXfs` they
 * are edit-time propagation bits meaning "if the named style's font changes later, push
 * that change into this record". Their defaults are asymmetric — `true` in
 * `cellStyleXfs`, `false` in `cellXfs` — and the schema declares no default for any of
 * them, so those defaults exist only in the Microsoft document.
 *
 * What Excel *does* have is a fallback chain for choosing **which single `xf` index**
 * applies, and exactly one genuine overlay layer (`dxf`, for conditional formatting).
 */

export const S_NAMESPACE = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/**
 * Built-in number formats whose `formatCode` is implied rather than written.
 *
 * Where Excel disagrees with the standard, both are recorded: the deviation is
 * invisible in the file — the same `numFmtId` means different things — and a two-digit
 * versus four-digit year is not a subtlety anyone spots by reading markup.
 */
export const BUILT_IN_NUMBER_FORMATS: ReadonlyMap<number, { spec: string; excel?: string }> = new Map([
  [0, { spec: 'General' }],
  [1, { spec: '0' }],
  [2, { spec: '0.00' }],
  [3, { spec: '#,##0' }],
  [4, { spec: '#,##0.00' }],
  [9, { spec: '0%' }],
  [10, { spec: '0.00%' }],
  [11, { spec: '0.00E+00' }],
  [12, { spec: '# ?/?' }],
  [13, { spec: '# ??/??' }],
  [14, { spec: 'mm-dd-yy', excel: 'm/d/yyyy' }],
  [15, { spec: 'd-mmm-yy' }],
  [16, { spec: 'd-mmm' }],
  [17, { spec: 'mmm-yy' }],
  [18, { spec: 'h:mm AM/PM' }],
  [19, { spec: 'h:mm:ss AM/PM' }],
  [20, { spec: 'h:mm' }],
  [21, { spec: 'h:mm:ss' }],
  [22, { spec: 'm/d/yy h:mm', excel: 'm/d/yyyy h:mm' }],
  [37, { spec: '#,##0 ;(#,##0)', excel: '#,##0_);(#,##0)' }],
  [38, { spec: '#,##0 ;[Red](#,##0)', excel: '#,##0_);[Red](#,##0)' }],
  [39, { spec: '#,##0.00;(#,##0.00)', excel: '#,##0.00_);(#,##0.00)' }],
  [40, { spec: '#,##0.00;[Red](#,##0.00)', excel: '#,##0.00_);[Red](#,##0.00)' }],
  [45, { spec: 'mm:ss' }],
  [46, { spec: '[h]:mm:ss' }],
  [47, { spec: 'mmss.0', excel: 'mm:ss.0' }],
  [48, { spec: '##0.0E+0' }],
  [49, { spec: '@' }]
]);

export interface ExcelStyleSheet {
  /** Custom formats only, keyed by the **explicit `numFmtId` value**, not by position. */
  numFmts: Map<number, string>;
  /** Positionally indexed component tables. */
  fonts: Element[];
  fills: Element[];
  borders: Element[];
  cellStyleXfs: Element[];
  cellXfs: Element[];
  /** Named styles, joined to `cellStyleXfs` on `xfId`. */
  cellStyles: { name: string; xfId: number; builtinId: number | null }[];
}

const childrenNamed = (parent: Element | null, localName: string): Element[] =>
  parent
    ? Array.from(parent.children).filter(
        el => el.namespaceURI === S_NAMESPACE && el.localName === localName
      )
    : [];

const firstNamed = (root: Document | Element, localName: string): Element | null =>
  ('getElementsByTagNameNS' in root
    ? root.getElementsByTagNameNS(S_NAMESPACE, localName).item(0)
    : null) ?? null;

const intAttr = (el: Element | null, name: string): number | null => {
  const raw = el?.getAttribute(name);
  if (raw === null || raw === undefined) return null;
  const parsed = Number.parseInt(raw, 10);
  return Number.isNaN(parsed) ? null : parsed;
};

/** Parses `xl/styles.xml`. */
export const parseExcelStyles = (stylesXml: string): ExcelStyleSheet => {
  const doc = new DOMParser().parseFromString(stylesXml, 'application/xml');

  const numFmts = new Map<number, string>();
  for (const el of childrenNamed(firstNamed(doc, 'numFmts'), 'numFmt')) {
    const id = intAttr(el, 'numFmtId');
    const code = el.getAttribute('formatCode');
    if (id !== null && code !== null) numFmts.set(id, code);
  }

  const cellStyles = childrenNamed(firstNamed(doc, 'cellStyles'), 'cellStyle').map(el => ({
    name: el.getAttribute('name') ?? '',
    xfId: intAttr(el, 'xfId') ?? 0,
    builtinId: intAttr(el, 'builtinId')
  }));

  return {
    numFmts,
    fonts: childrenNamed(firstNamed(doc, 'fonts'), 'font'),
    fills: childrenNamed(firstNamed(doc, 'fills'), 'fill'),
    borders: childrenNamed(firstNamed(doc, 'borders'), 'border'),
    cellStyleXfs: childrenNamed(firstNamed(doc, 'cellStyleXfs'), 'xf'),
    cellXfs: childrenNamed(firstNamed(doc, 'cellXfs'), 'xf'),
    cellStyles
  };
};

/** Where the winning `xf` index came from. */
export type XfSource = 'cell' | 'row' | 'column' | 'default';

export interface CellFormatInput {
  /** `c/@s` — present only when the cell element exists and carries it. */
  cellStyleIndex?: number | null;
  /** `row/@s`. */
  rowStyleIndex?: number | null;
  /** `row/@customFormat` — `row/@s` applies **only** when this is true. */
  rowCustomFormat?: boolean;
  /** `col/@style` — applies only to cells not yet allocated in the column. */
  columnStyleIndex?: number | null;
  /** Whether a `c` element actually exists for this address. */
  cellExists?: boolean;
}

export interface ResolvedCellFormat {
  /** The single winning index into `cellXfs`. */
  xfIndex: number;
  source: XfSource;
  numFmtId: number;
  /** The format code, resolved through custom formats then the built-in table. */
  formatCode: string | null;
  /** True when the code came from the implied built-in table rather than the file. */
  isBuiltIn: boolean;
  font: Element | null;
  fill: Element | null;
  border: Element | null;
  alignment: Element | null;
  /**
   * The named style this record was built from.
   *
   * Provenance only. Nothing from it is merged in — see the module comment.
   */
  namedStyle: string | null;
  trace: string[];
  notes: string[];
}

/**
 * Chooses the single `xf` that applies to a cell.
 *
 * Fallback, not merge: whichever source wins supplies the whole format. There is no
 * per-property inheritance anywhere in this chain.
 */
export const selectXfIndex = (input: CellFormatInput): { index: number; source: XfSource; trace: string[] } => {
  const trace: string[] = [];

  if (input.cellExists !== false && typeof input.cellStyleIndex === 'number') {
    trace.push(`cell carries s="${input.cellStyleIndex}"`);
    return { index: input.cellStyleIndex, source: 'cell', trace };
  }

  // row/@s is inert unless customFormat says otherwise. Applying it unconditionally
  // formats every cell in every row that merely records a style index.
  if (typeof input.rowStyleIndex === 'number') {
    if (input.rowCustomFormat) {
      trace.push(`row carries s="${input.rowStyleIndex}" with customFormat="1"`);
      return { index: input.rowStyleIndex, source: 'row', trace };
    }
    trace.push(`row s="${input.rowStyleIndex}" ignored because customFormat is not set`);
  }

  // col/@style governs cells "not yet allocated in the column" - it is a default for
  // new cells, not a format applied over existing ones.
  if (typeof input.columnStyleIndex === 'number' && input.cellExists === false) {
    trace.push(`column default style="${input.columnStyleIndex}" applies to this unallocated cell`);
    return { index: input.columnStyleIndex, source: 'column', trace };
  }

  trace.push('falling back to cellXfs[0], the default format');
  return { index: 0, source: 'default', trace };
};

/**
 * Resolves the effective format of a cell.
 *
 * `preferExcelBehaviour` controls which reading of the built-in number formats is
 * reported. It defaults to Excel's, because the question this answers is what the user
 * will see, not what a conformant consumer would do.
 */
export const resolveCellFormat = (
  sheet: ExcelStyleSheet,
  input: CellFormatInput,
  options: { preferExcelBehaviour?: boolean } = {}
): ResolvedCellFormat => {
  const preferExcel = options.preferExcelBehaviour ?? true;
  const { index, source, trace } = selectXfIndex(input);
  const notes: string[] = [];

  const xf = sheet.cellXfs[index] ?? null;
  if (!xf) {
    notes.push(`cellXfs has no entry at index ${index}; the file references a format that does not exist`);
  }

  const numFmtId = intAttr(xf, 'numFmtId') ?? 0;
  const fontId = intAttr(xf, 'fontId');
  const fillId = intAttr(xf, 'fillId');
  const borderId = intAttr(xf, 'borderId');
  const xfId = intAttr(xf, 'xfId');

  // numFmts is keyed by the explicit numFmtId attribute, unlike every other component
  // table, which is addressed positionally. Indexing it by position is a real bug.
  const custom = sheet.numFmts.get(numFmtId) ?? null;
  const builtIn = BUILT_IN_NUMBER_FORMATS.get(numFmtId) ?? null;

  let formatCode: string | null = custom;
  let isBuiltIn = false;
  if (custom === null && builtIn) {
    formatCode = preferExcel ? (builtIn.excel ?? builtIn.spec) : builtIn.spec;
    isBuiltIn = true;
    if (builtIn.excel && builtIn.excel !== builtIn.spec) {
      notes.push(
        `numFmtId ${numFmtId} is implied, and Excel differs from the standard here: ` +
        `the standard says "${builtIn.spec}", Excel uses "${builtIn.excel}"`
      );
    }
  }
  if (custom === null && !builtIn) {
    notes.push(`numFmtId ${numFmtId} is neither defined in numFmts nor a documented built-in; its meaning is implementation-defined`);
  }

  const named = xfId !== null
    ? sheet.cellStyles.find(style => style.xfId === xfId)?.name ?? null
    : null;

  if (xfId !== null) {
    trace.push(
      `xfId ${xfId}${named ? ` ("${named}")` : ''} records which named style this was built from; ` +
      `it is NOT merged in - Excel reads only the cellXfs record`
    );
  }

  // The apply* flags are edit-time propagation bits. Say so if any are present, since
  // reading them as render gates is the most common way to get this wrong.
  const applyFlags = ['applyNumberFormat', 'applyFont', 'applyFill', 'applyBorder', 'applyAlignment', 'applyProtection']
    .filter(name => xf?.getAttribute(name) !== null && xf?.getAttribute(name) !== undefined);
  if (applyFlags.length > 0) {
    notes.push(
      `${applyFlags.join(', ')} present on this record. In cellXfs these control whether a later ` +
      `change to the named style propagates into this record - they do not gate rendering, and are ignored here`
    );
  }

  const atIndex = (table: Element[], id: number | null, label: string): Element | null => {
    if (id === null) return null;
    const element = table[id] ?? null;
    if (!element) notes.push(`${label} index ${id} is out of range for the ${label} table`);
    return element;
  };

  return {
    xfIndex: index,
    source,
    numFmtId,
    formatCode,
    isBuiltIn,
    font: atIndex(sheet.fonts, fontId, 'font'),
    fill: atIndex(sheet.fills, fillId, 'fill'),
    border: atIndex(sheet.borders, borderId, 'border'),
    alignment: childrenNamed(xf, 'alignment')[0] ?? null,
    namedStyle: named,
    trace,
    notes
  };
};

// --- Dates -----------------------------------------------------------------

/**
 * Days between the 1900 and 1904 epochs.
 *
 * Cross-checked two ways against ECMA-376 §18.17.4.1's own figures: the lower limits
 * differ by 695055 − 693593, and the upper limits by 2958465 − 2957003. Both give 1462.
 */
export const DATE_SYSTEM_OFFSET_DAYS = 1462;

export type DateSystem = '1900' | '1900-compat' | '1904';

/**
 * Converts an Excel serial number to a date.
 *
 * Three epochs are in play, which is more than most implementations account for:
 *
 * - `1900`        — base 1899-12-30, the true system, used by Strict files.
 * - `1900-compat` — base 1899-12-31, and it contains one extra day that does not
 *                   exist. Serial 60 is that phantom. **This is what Transitional
 *                   files use, so it is what real-world workbooks use.**
 * - `1904`        — base 1904-01-01.
 *
 * The phantom day's *existence* is provable from the standard's own arithmetic: the
 * compat system's epoch is a day later than the true system's, yet both place
 * 9999-12-31 at serial 2958465. Its *identity* is not stated anywhere in ECMA-376,
 * which never names a specific date, so this code does not claim one.
 */
export const serialToDate = (serial: number, system: DateSystem = '1900-compat'): Date | null => {
  if (!Number.isFinite(serial)) return null;
  // Excel supports no negative serials at all, so the standard's entire pre-epoch
  // range is unreachable in practice.
  if (serial < 0) return null;

  let epochUtcMs: number;
  let adjusted = serial;

  if (system === '1904') {
    epochUtcMs = Date.UTC(1904, 0, 1);
  } else if (system === '1900') {
    epochUtcMs = Date.UTC(1899, 11, 30);
  } else {
    epochUtcMs = Date.UTC(1899, 11, 31);
    // Above the phantom, the compat system agrees with a 1899-12-30 base. Below it,
    // it agrees with 1899-12-31. Shifting by one past the phantom reconciles them.
    if (serial >= 61) adjusted = serial - 1;
  }

  const MS_PER_DAY = 86400000;
  return new Date(epochUtcMs + adjusted * MS_PER_DAY);
};

/**
 * Reads the workbook's date system.
 *
 * Note that Office ignores `dateCompatibility` entirely and selects on conformance
 * instead: Strict files get the true 1900 system, everything else gets the compat
 * system. Since Excel writes Transitional by default, real files carry the phantom.
 */
export const readDateSystem = (workbookXml: string): DateSystem => {
  const doc = new DOMParser().parseFromString(workbookXml, 'application/xml');
  const pr = firstNamed(doc, 'workbookPr');
  if (pr?.getAttribute('date1904') && !['0', 'false'].includes((pr.getAttribute('date1904') ?? '').toLowerCase())) {
    return '1904';
  }
  const conformance = doc.documentElement?.getAttribute('conformance');
  return conformance === 'strict' ? '1900' : '1900-compat';
};
