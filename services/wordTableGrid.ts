/**
 * WordprocessingML table geometry — the column grid, and the merges laid over it.
 *
 * This is the *shape* of a table, not its appearance. Conditional formatting — banding,
 * first/last row and column, the `w:tblLook` bitmask — lives in `wordTableStyles.ts` and
 * is deliberately not repeated here.
 *
 * THE MODEL.
 *
 * A `w:tbl` declares its columns once, up front:
 *
 *   <w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/></w:tblGrid>
 *
 * Every `w:tr` then lays cells across that grid. A cell occupies one grid column unless
 * `w:tcPr/w:gridSpan` says otherwise, and a row may also skip columns at either end with
 * `w:trPr/w:gridBefore` and `w:trPr/w:gridAfter`. So the arithmetic that must hold for
 * every row is:
 *
 *   gridBefore + Σ gridSpan + gridAfter  ==  number of w:gridCol
 *
 * WHY THIS DESERVES AN ANALYZER: WORD REPAIRS IT AND SAYS NOTHING.
 *
 * When a row does not add up, Word does not complain and does not refuse the file. It
 * guesses — widening the last cell, inventing a column, dropping the overflow — and
 * renders something plausible. The author sees a table that looks right. A converter, a
 * headless renderer, or a different word processor guesses differently, and the layout
 * arrives mangled somewhere the author will never look. [MS-OI29500] records 71 normative
 * variations against Part 1 §17.4, so "Word renders it" is unusually weak evidence that
 * a table is correct in this clause specifically.
 *
 * THE TRAP IN VERTICAL MERGING: THE DEFAULT IS INVERTED.
 *
 *   <w:vMerge w:val="restart"/>   begins a vertical merge
 *   <w:vMerge w:val="continue"/>  continues the merge above
 *   <w:vMerge/>                   ALSO continues it — absent @w:val means *continue*
 *
 * Almost every other on/off-ish element in WordprocessingML defaults the useful way
 * round; this one defaults to the passive half of the pair. Code that treats a bare
 * `<w:vMerge/>` as "starts a merge" — the obvious reading — produces a cell that merges
 * into nothing, and a `continue` in the very first row, or in a column with no `restart`
 * anywhere above it, is orphaned. Word absorbs the orphan into whatever is adjacent.
 * Nothing warns anyone.
 *
 * THE TRAP IN WALKING THE MARKUP: NESTED TABLES AND REVISION SHADOWS.
 *
 * Three constructs make `querySelectorAll` the wrong tool here, and each one produces a
 * confident wrong answer rather than a crash:
 *
 *   1. `w:tc` may contain a whole `w:tbl`. A descendant walk for `w:tr` from the outer
 *      table collects the inner table's rows too, and then measures them against the
 *      *outer* grid — inventing mismatches in a document that is perfectly correct.
 *   2. `w:tblGrid` may contain `w:tblGridChange`, which contains **another `w:tblGrid`**
 *      (`CT_TblGridBase`) holding the pre-revision `w:gridCol` list. A descendant walk
 *      for `w:gridCol` therefore double-counts the columns of any table whose grid was
 *      ever edited with track-changes on.
 *   3. `w:tcPr` may contain `w:tcPrChange`, which contains **another `w:tcPr`**
 *      (`CT_TcPrInner`) holding the pre-revision `w:gridSpan` and `w:vMerge`. Same
 *      failure: the old span is read as if it were current.
 *
 * Everything below therefore walks *direct children* only, descending solely through the
 * wrappers that are transparent to the table structure (`w:sdt`/`w:sdtContent` and
 * `w:customXml`, which the schema permits around both rows and cells).
 *
 * SCHEMA FACTS VERIFIED against the Open XML SDK schema data
 * (`schemas_openxmlformats_org_wordprocessingml_2006_main.json`):
 *
 *   - `CT_TblGridCol/w:gridCol` carries exactly one attribute, `@w:w`, typed
 *     ST_TwipsMeasure_O12 (2007) or a union of an unsigned decimal with a
 *     `[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)` pattern (2010+). Bare integers are twips.
 *   - `CT_TblGrid/w:tblGrid` has children `w:gridCol` and `w:tblGridChange`; and
 *     `CT_TblGridChange` has a child `CT_TblGridBase/w:tblGrid` — the nested-grid trap.
 *   - `w:gridSpan`, `w:gridBefore`, `w:gridAfter` are all `CT_DecimalNumber`, whose
 *     `@w:val` is Int32 and carries a **RequiredValidator**. A `w:gridSpan` with no
 *     `@w:val` is schema-invalid, not a defaulted 1.
 *   - `CT_VMerge/w:vMerge` and `CT_HMerge/w:hMerge` each carry one optional `@w:val` of
 *     type `ST_Merge`, whose facets are exactly `continue` and `restart`.
 *   - `CT_TblWidth` (`w:tblW`, `w:tcW`) carries `@w:w` and `@w:type`, the latter an
 *     `ST_TblWidth` with facets `nil`, `pct` (fiftieths of a percent), `dxa` (twips) and
 *     `auto`.
 *   - `CT_TcPr` children include `w:tcPrChange`; `CT_Row` children include `w:tc`,
 *     `w:customXml` (CT_CustomXmlCell) and `w:sdt` (CT_SdtCell); `CT_Tbl` children
 *     include `w:tr`, `w:customXml` (CT_CustomXmlRow) and `w:sdt` (CT_SdtRow).
 *
 * ⚠️ NOT VERIFIED from the SDK schema data, and flagged rather than asserted:
 *
 *   - **That an omitted `w:vMerge/@w:val` means `continue`.** The schema shows only that
 *     the attribute is optional and that the enumeration has two members; it carries no
 *     default. The continue-by-default rule is ECMA-376 Part 1 §17.4.85 prose, which is
 *     not in this dataset. It is encoded here because every implementation this project
 *     could check agrees on it, but the *citation* is to prose we have not read directly.
 *     The same applies to `w:hMerge/@w:val` (§17.4.17).
 *   - **What Word actually does with a row that does not add up.** That it repairs
 *     rather than refuses is observed behaviour, not a specified rule, and the specific
 *     repair is not documented anywhere we can cite. The findings below therefore say
 *     "renderers disagree", not "Word does X".
 *   - **The width tolerance.** `WIDTH_TOLERANCE_*` below are this module's judgement
 *     about what counts as a contradiction, not a threshold from any specification.
 */

import { W_NAMESPACE } from './wordStyleResolver';
import { finding, renderFindings, type Finding, type Severity } from './findings';

/**
 * Severity and silence for each kind, decided once here rather than at each call site.
 *
 * Every one of these is SILENT, and that is the whole point of the module: a table whose
 * geometry contradicts itself still opens, still renders, and still looks like a table.
 * The damage appears only in some *other* renderer — a PDF pipeline, a browser preview,
 * a competing word processor — which resolves the contradiction differently.
 */
const TABLE_RULES = {
  /** A row's columns do not add up to the declared grid. The core defect. */
  'row-span-mismatch':      { severity: 'error',   silent: true },
  /** `w:vMerge` continuing a merge that was never started. */
  'vmerge-orphan':          { severity: 'error',   silent: true },
  /** A continuation whose gridSpan disagrees with the restart it continues. */
  'vmerge-span-mismatch':   { severity: 'warning', silent: true },
  /** A restart nothing continues — renders identically to no merge at all. */
  'vmerge-restart-alone':   { severity: 'note',    silent: true },
  /** `w:hMerge` continuing a merge with no restart to its left. */
  'hmerge-orphan':          { severity: 'error',   silent: true },
  /** `w:hMerge` used at all, where `w:gridSpan` is the modern spelling. */
  'hmerge-legacy':          { severity: 'note',    silent: true },
  /** `w:gridSpan`/`w:gridBefore`/`w:gridAfter` with a missing or unusable `@w:val`. */
  'span-value-invalid':     { severity: 'error',   silent: true },
  /** `@w:val` on a merge element outside the two-member `ST_Merge` enumeration. */
  'merge-value-invalid':    { severity: 'warning', silent: true },
  /** A table with rows but no `w:tblGrid`, so column widths are anybody's guess. */
  'grid-missing':           { severity: 'warning', silent: true },
  /** Rows disagreeing with one another, checked when there is no grid to check against. */
  'ragged-rows':            { severity: 'warning', silent: true },
  /** `w:tblW` in twips contradicting the sum of the `w:gridCol` widths. */
  'width-contradicts-grid': { severity: 'warning', silent: true },
  /** A row's `w:tcW` widths in twips contradicting the grid they sit on. */
  'cell-width-mismatch':    { severity: 'warning', silent: true }
} as const satisfies Record<string, { severity: Severity; silent: boolean }>;

export type TableProblemKind = keyof typeof TABLE_RULES;

/** Builds a table finding, applying the severity table above. */
const tableFinding = (
  kind: TableProblemKind,
  part: string,
  message: string,
  remediation: string,
  subject?: Record<string, string>
): Finding =>
  finding(`table/${kind}`, part, message, remediation, { ...TABLE_RULES[kind], subject });

/**
 * How far a width total may drift from the grid before it counts as a contradiction.
 *
 * ⚠️ A judgement call, not a specified threshold. Real documents accumulate a twip or
 * two of rounding through unit conversions; flagging those would bury the cases where a
 * cell was resized and the grid was not.
 */
const WIDTH_TOLERANCE_TWIPS = 20;
const WIDTH_TOLERANCE_FRACTION = 0.01;

/**
 * Wrappers the schema allows around rows and around cells, which carry no geometry of
 * their own. Structured document tags and custom XML are transparent here: a `w:tr`
 * inside a `w:sdt` is still a row of the table, so the walk descends through them.
 */
const TRANSPARENT_WRAPPERS = new Set(['sdt', 'sdtContent', 'customXml']);

/** `ST_Merge` has exactly two facets. */
export type MergeState = 'restart' | 'continue';

export interface TableCellGeometry {
  cell: Element;
  /** 0-based index of the first grid column this cell occupies. */
  startColumn: number;
  /** `w:gridSpan/@w:val`, defaulting to 1 when the element is absent. */
  gridSpan: number;
  /** null when the cell carries no `w:vMerge`. A bare `<w:vMerge/>` reads as `continue`. */
  vMerge: MergeState | null;
  hMerge: MergeState | null;
  /** `w:tcW/@w:w` in twips, or null when absent, non-`dxa`, or unit-suffixed. */
  widthTwips: number | null;
  /** `w:tcW/@w:type` verbatim — `nil`, `pct`, `dxa`, `auto`, or null. */
  widthType: string | null;
}

export interface TableRowGeometry {
  row: Element;
  /** 0-based index among this table's own rows, nested tables excluded. */
  index: number;
  cells: TableCellGeometry[];
  /** `w:trPr/w:gridBefore` — grid columns skipped before the first cell. */
  gridBefore: number;
  /** `w:trPr/w:gridAfter` — grid columns skipped after the last cell. */
  gridAfter: number;
  /** gridBefore + Σ gridSpan + gridAfter. Should equal the table's `columnCount`. */
  columnsCovered: number;
}

/** A vertical merge as it actually resolved, restart through final continuation. */
export interface VerticalMergeRegion {
  /** Grid column the merge starts at. */
  column: number;
  /** gridSpan of the cell that restarted it. */
  gridSpan: number;
  startRow: number;
  /** Equal to `startRow` when nothing continued the merge. */
  endRow: number;
  /** How many rows the merged cell spans. 1 means the merge did nothing. */
  rowSpan: number;
}

export interface TableGeometry {
  table: Element;
  /** Document order across the whole part, nested tables included. */
  index: number;
  /** 0 for a top-level table, 1 for one inside a cell of a top-level table, and so on. */
  depth: number;
  /** Number of `w:gridCol` in this table's own `w:tblGrid`. 0 when there is no grid. */
  columnCount: number;
  /** `w:gridCol/@w:w` per column, in twips; null where unit-suffixed or unparseable. */
  columnWidths: (number | null)[];
  /** Sum of `columnWidths`, or null when any column width was missing or unparseable. */
  gridWidthTwips: number | null;
  /** `w:tblPr/w:tblW/@w:w` in twips, or null when absent or not typed `dxa`. */
  declaredWidthTwips: number | null;
  /** `w:tblPr/w:tblW/@w:type` verbatim. */
  declaredWidthType: string | null;
  rows: TableRowGeometry[];
  verticalMerges: VerticalMergeRegion[];
  problems: Finding[];
}

const attr = (el: Element, name: string) => el.getAttributeNS(W_NAMESPACE, name);

/** Direct children in the w namespace. Namespaces arrive normalised — see conformance.ts. */
const wChildren = (parent: Element): Element[] =>
  Array.from(parent.children).filter(el => el.namespaceURI === W_NAMESPACE);

const wChild = (parent: Element | null, localName: string): Element | null =>
  parent ? (wChildren(parent).find(el => el.localName === localName) ?? null) : null;

/**
 * Direct structural children of one kind, seen through transparent wrappers.
 *
 * Never descends into a `w:tr` or a `w:tc`, which is precisely what keeps a nested
 * table's rows from being measured against the outer table's grid.
 */
const structuralChildren = (parent: Element, localName: string): Element[] => {
  const found: Element[] = [];
  const visit = (node: Element) => {
    for (const child of wChildren(node)) {
      if (child.localName === localName) found.push(child);
      else if (TRANSPARENT_WRAPPERS.has(child.localName)) visit(child);
    }
  };
  visit(parent);
  return found;
};

/**
 * `@w:val` as a positive integer, distinguishing "no element" from "element that does
 * not parse". `CT_DecimalNumber/@w:val` is required, so an absent value is a defect and
 * not a default.
 */
const readSpanValue = (
  el: Element | null
): { value: number; problem: 'missing' | 'unparseable' | 'non-positive' | null } => {
  if (!el) return { value: 0, problem: null };
  const raw = attr(el, 'val');
  if (raw === null) return { value: 0, problem: 'missing' };
  if (!/^-?\d+$/.test(raw.trim())) return { value: 0, problem: 'unparseable' };
  const parsed = Number(raw.trim());
  if (parsed <= 0) return { value: 0, problem: 'non-positive' };
  return { value: parsed, problem: null };
};

/**
 * `ST_Merge`, with the inverted default applied.
 *
 * An absent `@w:val` reads as `continue` — see the "NOT VERIFIED" note at the top: the
 * schema proves the attribute is optional and the enumeration two-membered, but the
 * default itself comes from specification prose this project has not read directly.
 */
const readMergeState = (el: Element | null): { state: MergeState | null; invalidValue: string | null } => {
  if (!el) return { state: null, invalidValue: null };
  const raw = attr(el, 'val');
  if (raw === null) return { state: 'continue', invalidValue: null };
  const trimmed = raw.trim();
  if (trimmed === 'restart') return { state: 'restart', invalidValue: null };
  if (trimmed === 'continue') return { state: 'continue', invalidValue: null };
  return { state: 'continue', invalidValue: trimmed };
};

/**
 * A width in twips, or null when it is not expressed in twips.
 *
 * Bare integers are twips. The 2010 union also admits `12.5mm` and friends; converting
 * those would need a unit table this module does not carry, so they read as null rather
 * than as a wrong number.
 */
const twipsOf = (raw: string | null): number | null => {
  if (raw === null) return null;
  const trimmed = raw.trim();
  return /^\d+$/.test(trimmed) ? Number(trimmed) : null;
};

const isTable = (el: Element) => el.namespaceURI === W_NAMESPACE && el.localName === 'tbl';

/**
 * Every `w:tbl` in the part, in document order, each tagged with its nesting depth.
 *
 * `querySelectorAll` excludes the root itself, so a root that *is* a table — which is
 * what the editor hands over when a `w:tbl` is the selected element — is added back.
 */
const collectTables = (root: ParentNode): Array<{ table: Element; depth: number }> => {
  const self = 'localName' in root && isTable(root as Element) ? [root as Element] : [];
  const tables = [
    ...self,
    ...Array.from(root.querySelectorAll('*')).filter(isTable)
  ];
  return tables.map(table => {
    let depth = 0;
    for (let p = table.parentElement; p; p = p.parentElement) {
      if (p.namespaceURI === W_NAMESPACE && p.localName === 'tbl') depth += 1;
    }
    return { table, depth };
  });
};

const describeRow = (index: number) => `row ${index + 1}`;

/**
 * Read one table's geometry and everything wrong with it.
 *
 * Analyses this table alone: a nested table is a separate `TableGeometry` with its own
 * grid, and its rows never count towards its parent's.
 */
const readOneTable = (table: Element, index: number, depth: number, part: string): TableGeometry => {
  const problems: Finding[] = [];
  const where = depth > 0 ? `nested table ${index + 1}` : `table ${index + 1}`;

  // Direct child only: w:tblGridChange holds a second w:tblGrid with the pre-revision
  // columns, and counting those would double the column count of any track-changed grid.
  const gridEl = wChild(table, 'tblGrid');
  const gridCols = gridEl ? wChildren(gridEl).filter(el => el.localName === 'gridCol') : [];
  const columnWidths = gridCols.map(col => twipsOf(attr(col, 'w')));
  const gridWidthTwips = columnWidths.some(w => w === null)
    ? null
    : columnWidths.reduce<number>((sum, w) => sum + (w ?? 0), 0);

  const tblPr = wChild(table, 'tblPr');
  const tblW = wChild(tblPr, 'tblW');
  const declaredWidthType = tblW ? attr(tblW, 'type') : null;
  const declaredWidthTwips =
    tblW && declaredWidthType === 'dxa' ? twipsOf(attr(tblW, 'w')) : null;

  const rows: TableRowGeometry[] = [];

  structuralChildren(table, 'tr').forEach((row, rowIndex) => {
    const trPr = wChild(row, 'trPr');

    const readEdge = (name: 'gridBefore' | 'gridAfter'): number => {
      const el = wChild(trPr, name);
      const { value, problem } = readSpanValue(el);
      if (problem) {
        problems.push(tableFinding(
          'span-value-invalid', part,
          `${where}, ${describeRow(rowIndex)}: <w:${name}> has ${problem === 'missing' ? 'no w:val' : `an unusable w:val ("${attr(el!, 'val')}")`}. @w:val is required on this element, and it decides how many grid columns the row skips, so the row's width is undefined.`,
          `Give <w:${name}> a positive integer w:val, or remove the element if the row skips no columns.`,
          { row: String(rowIndex + 1), element: name }
        ));
        return 0;
      }
      return value;
    };

    const gridBefore = readEdge('gridBefore');
    const gridAfter = readEdge('gridAfter');

    let column = gridBefore;
    const cells: TableCellGeometry[] = [];

    // Direct cells only: a w:tc containing a nested w:tbl contributes itself, never the
    // inner table's cells.
    for (const cell of structuralChildren(row, 'tc')) {
      // Direct child again: w:tcPrChange carries a second w:tcPr with the pre-revision
      // gridSpan and vMerge.
      const tcPr = wChild(cell, 'tcPr');
      const gridSpanEl = wChild(tcPr, 'gridSpan');
      const { value: spanValue, problem: spanProblem } = readSpanValue(gridSpanEl);
      if (spanProblem) {
        problems.push(tableFinding(
          'span-value-invalid', part,
          `${where}, ${describeRow(rowIndex)}, cell ${cells.length + 1}: <w:gridSpan> has ${spanProblem === 'missing' ? 'no w:val' : `an unusable w:val ("${attr(gridSpanEl!, 'val')}")`}. @w:val is required, so this cell's width across the grid is undefined and each renderer picks its own.`,
          'Give <w:gridSpan> a positive integer w:val, or remove the element for a single-column cell.',
          { row: String(rowIndex + 1), cell: String(cells.length + 1) }
        ));
      }
      const gridSpan = spanProblem ? 1 : spanValue || 1;

      const vMergeRead = readMergeState(wChild(tcPr, 'vMerge'));
      const hMergeRead = readMergeState(wChild(tcPr, 'hMerge'));
      for (const [element, read] of [['vMerge', vMergeRead], ['hMerge', hMergeRead]] as const) {
        if (read.invalidValue === null) continue;
        problems.push(tableFinding(
          'merge-value-invalid', part,
          `${where}, ${describeRow(rowIndex)}, cell ${cells.length + 1}: <w:${element} w:val="${read.invalidValue}"/> is not one of the two permitted values. ST_Merge admits only "restart" and "continue"; this reads as a continuation, which is what an omitted value means.`,
          `Set w:val to "restart" to begin the merge or "continue" to continue it — or omit the attribute, which also means continue.`,
          { row: String(rowIndex + 1), cell: String(cells.length + 1), value: read.invalidValue }
        ));
      }

      const tcW = wChild(tcPr, 'tcW');
      const widthType = tcW ? attr(tcW, 'type') : null;

      cells.push({
        cell,
        startColumn: column,
        gridSpan,
        vMerge: vMergeRead.state,
        hMerge: hMergeRead.state,
        widthTwips: tcW && widthType === 'dxa' ? twipsOf(attr(tcW, 'w')) : null,
        widthType
      });
      column += gridSpan;
    }

    rows.push({
      row,
      index: rowIndex,
      cells,
      gridBefore,
      gridAfter,
      columnsCovered: column + gridAfter
    });
  });

  const columnCount = gridCols.length;

  if (columnCount === 0 && rows.length > 0) {
    problems.push(tableFinding(
      'grid-missing', part,
      `${where} has ${rows.length} row(s) but declares no <w:gridCol>, so nothing states how many columns it has or how wide they are. Each consumer infers the grid from the cells, and they do not infer the same one.`,
      'Add a <w:tblGrid> with one <w:gridCol w:w="…"/> per column, widths in twentieths of a point.',
      { rows: String(rows.length) }
    ));
  }

  // The core check. gridBefore and gridAfter are part of the sum: leaving them out
  // reports a perfectly correct indented row as broken, which is the false positive this
  // check is most likely to produce.
  if (columnCount > 0) {
    for (const row of rows) {
      if (row.columnsCovered === columnCount) continue;
      const edges = row.gridBefore > 0 || row.gridAfter > 0
        ? ` (including w:gridBefore ${row.gridBefore} and w:gridAfter ${row.gridAfter})`
        : '';
      problems.push(tableFinding(
        'row-span-mismatch', part,
        `${where}, ${describeRow(row.index)} covers ${row.columnsCovered} grid column(s)${edges} but the table declares ${columnCount}. Word repairs the row and draws something plausible, so the document looks correct; a converter or another renderer resolves the same contradiction differently and the layout arrives wrong.`,
        row.columnsCovered < columnCount
          ? `Add ${columnCount - row.columnsCovered} more grid column(s) to this row — another cell, a larger w:gridSpan, or a w:gridAfter — or remove a <w:gridCol> so the grid matches.`
          : `Remove ${row.columnsCovered - columnCount} grid column(s) from this row, or add <w:gridCol> entries so the grid declares ${row.columnsCovered} columns.`,
        { row: String(row.index + 1), covered: String(row.columnsCovered), declared: String(columnCount) }
      ));
    }
  } else if (rows.length > 1) {
    // With no grid to check against, rows can still be checked against each other.
    const counts = rows.map(r => r.columnsCovered);
    const tally = new Map<number, number>();
    for (const c of counts) tally.set(c, (tally.get(c) ?? 0) + 1);
    const modal = [...tally.entries()].sort((a, b) => b[1] - a[1] || a[0] - b[0])[0][0];
    for (const row of rows) {
      if (row.columnsCovered === modal) continue;
      problems.push(tableFinding(
        'ragged-rows', part,
        `${where}, ${describeRow(row.index)} covers ${row.columnsCovered} column(s) where most rows cover ${modal}. With no <w:tblGrid> to settle it, each consumer picks a different column count for the table.`,
        'Declare a <w:tblGrid> and make every row add up to it.',
        { row: String(row.index + 1), covered: String(row.columnsCovered), typical: String(modal) }
      ));
    }
  }

  const verticalMerges = resolveVerticalMerges(rows, where, part, problems);
  checkHorizontalMerges(rows, where, part, problems);
  checkWidths(
    { where, part, columnCount, gridWidthTwips, declaredWidthTwips, declaredWidthType, rows },
    problems
  );

  return {
    table,
    index,
    depth,
    columnCount,
    columnWidths,
    gridWidthTwips,
    declaredWidthTwips,
    declaredWidthType,
    rows,
    verticalMerges,
    problems
  };
};

/**
 * Resolve `w:vMerge` down the grid, reporting continuations that merge into nothing.
 *
 * A merge is keyed by the grid column its restart *starts at*. A continuation in a later
 * row belongs to that merge only if it starts at the same column — a row whose earlier
 * cells changed width shifts every later cell along the grid, which silently detaches
 * the continuation from the restart it was drawn to continue.
 */
const resolveVerticalMerges = (
  rows: readonly TableRowGeometry[],
  where: string,
  part: string,
  problems: Finding[]
): VerticalMergeRegion[] => {
  const regions: VerticalMergeRegion[] = [];
  /** Grid column → index into `regions` of the merge currently open at that column. */
  const open = new Map<number, number>();

  for (const row of rows) {
    const startingHere = new Map<number, TableCellGeometry>();
    for (const cell of row.cells) startingHere.set(cell.startColumn, cell);

    for (const cell of row.cells) {
      if (cell.vMerge === 'restart') {
        regions.push({
          column: cell.startColumn,
          gridSpan: cell.gridSpan,
          startRow: row.index,
          endRow: row.index,
          rowSpan: 1
        });
        open.set(cell.startColumn, regions.length - 1);
        continue;
      }
      if (cell.vMerge !== 'continue') continue;

      const regionIndex = open.get(cell.startColumn);
      if (regionIndex === undefined) {
        problems.push(tableFinding(
          'vmerge-orphan', part,
          row.index === 0
            ? `${where}, ${describeRow(row.index)}, grid column ${cell.startColumn + 1}: a <w:vMerge> continuation sits in the first row, so there is no cell above for it to merge into. Note that an omitted w:val means "continue", not "restart" — a bare <w:vMerge/> here is the usual cause.`
            : `${where}, ${describeRow(row.index)}, grid column ${cell.startColumn + 1}: a <w:vMerge> continuation has no <w:vMerge w:val="restart"/> above it in that column, so it merges into nothing. Word absorbs the orphan into an adjacent cell and the table looks intact; other renderers leave a stray cell or drop its content.`,
          row.index === 0
            ? 'Set w:val="restart" on this cell, or remove the <w:vMerge> if the cell was never meant to merge.'
            : 'Add w:val="restart" to the topmost cell of the intended merge, or remove this <w:vMerge>. Check too that the cells above start at the same grid column — a changed w:gridSpan earlier in a row shifts them.',
          { row: String(row.index + 1), column: String(cell.startColumn + 1) }
        ));
        continue;
      }

      const region = regions[regionIndex];
      region.endRow = row.index;
      region.rowSpan += 1;
      if (region.gridSpan !== cell.gridSpan) {
        problems.push(tableFinding(
          'vmerge-span-mismatch', part,
          `${where}, ${describeRow(row.index)}, grid column ${cell.startColumn + 1}: this continuation spans ${cell.gridSpan} column(s) but the merge it continues, started in ${describeRow(region.startRow)}, spans ${region.gridSpan}. A merged region with ragged edges has no single correct rendering.`,
          `Make the continuation's w:gridSpan ${region.gridSpan}, matching the cell that restarted the merge.`,
          { row: String(row.index + 1), column: String(cell.startColumn + 1) }
        ));
      }
    }

    // A merge survives into the next row only if this row carries a cell at the same
    // column that is itself part of a merge. Anything else — a plain cell, or no cell
    // starting at that column at all — ends it.
    for (const [column] of [...open]) {
      const cell = startingHere.get(column);
      if (!cell || cell.vMerge === null) open.delete(column);
    }
  }

  for (const region of regions) {
    if (region.rowSpan > 1) continue;
    problems.push(tableFinding(
      'vmerge-restart-alone', part,
      `${where}, ${describeRow(region.startRow)}, grid column ${region.column + 1}: <w:vMerge w:val="restart"/> begins a vertical merge that no later row continues, so the cell spans one row and the merge has no effect.`,
      'Add <w:vMerge/> to the cells below that should join the merge, or remove the restart.',
      { row: String(region.startRow + 1), column: String(region.column + 1) }
    ));
  }

  return regions;
};

/**
 * `w:hMerge` — the horizontal equivalent, and largely a legacy of the pre-`gridSpan`
 * spelling. It is still legal, still read by Word, and still breaks the same way: a
 * continuation with no restart to its left in the same row merges into nothing.
 */
const checkHorizontalMerges = (
  rows: readonly TableRowGeometry[],
  where: string,
  part: string,
  problems: Finding[]
): void => {
  let usedAnywhere = false;

  for (const row of rows) {
    let openToTheLeft = false;
    for (const [position, cell] of row.cells.entries()) {
      if (cell.hMerge === null) {
        openToTheLeft = false;
        continue;
      }
      usedAnywhere = true;
      if (cell.hMerge === 'restart') {
        openToTheLeft = true;
        continue;
      }
      if (!openToTheLeft) {
        problems.push(tableFinding(
          'hmerge-orphan', part,
          `${where}, ${describeRow(row.index)}, cell ${position + 1}: a <w:hMerge> continuation has no <w:hMerge w:val="restart"/> to its left in this row, so it merges into nothing. As with w:vMerge, an omitted w:val means "continue".`,
          'Add w:val="restart" to the leftmost cell of the intended merge, or express the merge with <w:gridSpan> instead.',
          { row: String(row.index + 1), cell: String(position + 1) }
        ));
      }
    }
  }

  if (usedAnywhere) {
    problems.push(tableFinding(
      'hmerge-legacy', part,
      `${where} expresses horizontal merging with <w:hMerge>. It is legal and Word reads it, but <w:gridSpan> is the spelling every current producer writes, and tooling that only handles gridSpan will render these cells unmerged.`,
      'Replace each hMerge run with a single cell carrying <w:gridSpan w:val="n"/>.',
      {}
    ));
  }
};

/**
 * Widths that contradict the grid.
 *
 * Only `dxa` (twips) values are compared. `pct`, `auto` and `nil` are measured against
 * something this module cannot see — the page, the containing cell — so a mismatch
 * against them would be a guess, and a guess stated as a finding is worse than silence.
 */
const checkWidths = (
  ctx: {
    where: string;
    part: string;
    columnCount: number;
    gridWidthTwips: number | null;
    declaredWidthTwips: number | null;
    declaredWidthType: string | null;
    rows: readonly TableRowGeometry[];
  },
  problems: Finding[]
): void => {
  const { where, part, columnCount, gridWidthTwips, declaredWidthTwips, rows } = ctx;
  if (gridWidthTwips === null || gridWidthTwips === 0) return;

  const contradicts = (total: number) => {
    const delta = Math.abs(total - gridWidthTwips);
    return delta > Math.max(WIDTH_TOLERANCE_TWIPS, gridWidthTwips * WIDTH_TOLERANCE_FRACTION);
  };

  if (declaredWidthTwips !== null && contradicts(declaredWidthTwips)) {
    problems.push(tableFinding(
      'width-contradicts-grid', part,
      `${where} declares <w:tblW w:w="${declaredWidthTwips}" w:type="dxa"/> but its ${columnCount} <w:gridCol> widths sum to ${gridWidthTwips} twips. The two disagree by ${Math.abs(declaredWidthTwips - gridWidthTwips)}, and which one wins is a renderer's choice.`,
      `Set w:tblW to ${gridWidthTwips}, or adjust the <w:gridCol> widths to sum to ${declaredWidthTwips}.`,
      { declared: String(declaredWidthTwips), grid: String(gridWidthTwips) }
    ));
  }

  for (const row of rows) {
    // Only rows that span the whole grid, with every cell measured in twips, can be
    // compared: a row that skips columns is measuring a different span of the grid.
    if (columnCount > 0 && row.columnsCovered !== columnCount) continue;
    if (row.gridBefore > 0 || row.gridAfter > 0) continue;
    if (row.cells.length === 0) continue;
    if (row.cells.some(c => c.widthTwips === null)) continue;

    const total = row.cells.reduce((sum, c) => sum + (c.widthTwips ?? 0), 0);
    if (!contradicts(total)) continue;
    problems.push(tableFinding(
      'cell-width-mismatch', part,
      `${where}, ${describeRow(row.index)}: its <w:tcW> widths sum to ${total} twips while the grid sums to ${gridWidthTwips}. The row occupies the full grid, so the two are measuring the same thing and disagree; the cell widths and the column widths will be reconciled differently by different renderers.`,
      'Recalculate the cell widths so they sum to the grid width, or update the <w:gridCol> widths to match the cells.',
      { row: String(row.index + 1), cells: String(total), grid: String(gridWidthTwips) }
    ));
  }
};

/**
 * Every table in a document part, each measured against its own grid.
 *
 * Nested tables appear as their own entries with `depth > 0`; a nested table's rows never
 * count towards its parent's grid, and a nested table's own defects are still reported.
 *
 * Pass the parsed `word/document.xml`, or any body part — headers, footers, footnotes and
 * endnotes all carry tables.
 */
export function readTableGrids(doc: Document | Element, partPath = ''): TableGeometry[] {
  const root: ParentNode =
    'documentElement' in doc && doc.documentElement ? doc.documentElement : (doc as Element);
  return collectTables(root).map(({ table, depth }, index) =>
    readOneTable(table, index, depth, partPath)
  );
}

/** Every finding across every table in the part, for the analyzer registry. */
export const tableGridFindings = (doc: Document | Element, partPath = ''): Finding[] =>
  readTableGrids(doc, partPath).flatMap(t => t.problems);

/** Word body parts, all of which can contain tables. */
const WORD_BODY_PART =
  /^word\/(?:document\d*|header[^/]*|footer[^/]*|footnotes\d*|endnotes\d*|comments\d*)\.xml$/;

/**
 * Fingerprint of a table, used to match the element open in the editor to the table it
 * came from. The selected markup is a detached fragment with no identity of its own, so
 * shape is the only handle available: column count, row count, and each row's spans.
 */
const tableSignature = (t: TableGeometry): string =>
  `${t.columnCount}|${t.rows.map(r => `${r.gridBefore}:${r.cells.map(c => c.gridSpan).join(',')}:${r.gridAfter}`).join(';')}`;

const signatureOfMarkup = (rawXml: string): string | null => {
  if (!/<w:tbl[\s>]/.test(rawXml)) return null;
  const doc = new DOMParser().parseFromString(rawXml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0 || !doc.documentElement) return null;
  const el = doc.documentElement;
  if (el.namespaceURI !== W_NAMESPACE || el.localName !== 'tbl') return null;
  const [only] = readTableGrids(el, '');
  return only ? tableSignature(only) : null;
};

/**
 * Evidence lines for the AI panel: what shape these tables actually are, and where the
 * geometry contradicts itself.
 *
 * Everything here is computed from the markup, never asserted by the model. What the
 * markup does not settle — the selected table's identity when two tables have the same
 * shape, the effect of a table style on the grid — goes to `unresolved` so the tier is
 * capped rather than the gap papered over.
 */
export function computeTableEvidenceForMarkup(
  parts: Record<string, string>,
  rawXml: string
): { lines: string[]; unresolved: string[] } | null {
  const entry = Object.entries(parts).find(([path]) => WORD_BODY_PART.test(path));
  if (!entry) return null;

  const [path, xml] = entry;
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) return null;

  const tables = readTableGrids(doc, path);
  if (tables.length === 0) return null;

  const lines: string[] = [];
  const unresolved: string[] = [];

  const nested = tables.filter(t => t.depth > 0).length;
  lines.push(
    `${path} contains ${tables.length} table(s)${nested > 0 ? `, ${nested} of them nested inside a cell of another table` : ''}. ` +
      `Each is measured against its own <w:tblGrid>; a nested table's rows do not count towards the grid of the table containing it.`
  );

  const wanted = signatureOfMarkup(rawXml);
  const matches = wanted === null ? [] : tables.filter(t => tableSignature(t) === wanted);
  const selected = matches.length === 1 ? matches[0] : null;

  if (wanted !== null && matches.length > 1) {
    unresolved.push(
      `The open selection is a table, but ${matches.length} tables in ${path} have the same shape, so which one is selected cannot be determined from the markup alone. The lines below describe every table in the part.`
    );
  }

  const described = selected ? [selected] : tables;
  for (const t of described) {
    const label = t.depth > 0 ? `Nested table ${t.index + 1}` : `Table ${t.index + 1}`;
    const widths = t.gridWidthTwips === null
      ? 'not all column widths are stated in twips'
      : `${t.gridWidthTwips} twips wide in total`;
    lines.push(
      t.columnCount === 0
        ? `${label} has ${t.rows.length} row(s) and no <w:tblGrid> at all, so its column count is inferred rather than declared.`
        : `${label} declares ${t.columnCount} grid column(s) (${widths}) and has ${t.rows.length} row(s).`
    );

    for (const region of t.verticalMerges) {
      if (region.rowSpan < 2) continue;
      lines.push(
        `${label}: a vertical merge starting at grid column ${region.column + 1} spans rows ${region.startRow + 1}–${region.endRow + 1}. ` +
          `Only the first of those cells carries w:val="restart"; the rest carry <w:vMerge/> or w:val="continue", and an omitted w:val means continue.`
      );
    }

    if (t.declaredWidthType !== null && t.declaredWidthType !== 'dxa') {
      unresolved.push(
        `${label} states its width as w:type="${t.declaredWidthType}", which is measured against the page or the containing cell rather than in twips, so it was not compared against the grid.`
      );
    }
  }

  lines.push(...renderFindings(described.flatMap(t => t.problems)));

  unresolved.push(
    `Whether a table style adjusts these widths at render time was not checked; only the geometry written on the table itself was read.`
  );

  return { lines, unresolved };
}
