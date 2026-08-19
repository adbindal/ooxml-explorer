/**
 * SpreadsheetML formulas — the number on screen is a cache, and it can be wrong.
 *
 * **This is the largest cluster of documented Office-versus-specification divergence in
 * the whole format: 218 normative variations against Part 1 §18.17** — three times the
 * next-largest SpreadsheetML clause (counted from the published [MS-OI29500] table of
 * contents; see `docs/ooxml-expert-agent/RESEARCH-STATE.md` §8p). Formulas are where
 * Excel and ECMA-376 agree least.
 *
 * A cell carries two things that can disagree:
 *
 *   <c r="B2" t="n"><f>SUM(A1:A10)</f><v>55</v></c>
 *           the program ──┘              └── the answer, from the last recalculation
 *
 * **Every reader that does not implement a calculation engine shows `<v>`.** That is
 * almost every reader other than Excel: converters, extractors, dashboards, `openpyxl`
 * in its default mode, and this tool. So a workbook can display numbers that its own
 * formulas would no longer produce, and nothing about it looks wrong — same failure
 * class as a Word field's cached result, at spreadsheet scale.
 *
 * Excel signals its own doubt, and those signals are worth reading:
 *
 *   - `calcPr/@fullCalcOnLoad="1"` — Excel will recalculate everything on open, so the
 *     stored values are **known stale by declaration**. Anything reading the file
 *     without calculating is reading numbers Excel has already disowned.
 *   - `calcPr/@calcMode="manual"` — the user turned automatic calculation off. Values
 *     drift from their formulas by design and nothing warns anyone.
 *   - `calcPr/@calcCompleted="0"` — the last calculation did not finish.
 *
 * SHARED FORMULAS, AND THE ONE THAT SILENTLY LOSES A FORMULA ENTIRELY.
 *
 * Excel compresses a filled-down column by writing the formula once:
 *
 *   <c r="B2"><f t="shared" ref="B2:B9" si="0">A2*2</f><v>4</v></c>   ← master, has text
 *   <c r="B3"><f t="shared" si="0"/><v>6</v></c>                      ← follower, EMPTY
 *
 * A follower carries **no formula text at all**. Its formula exists only as an offset
 * from the master's. Delete or fail to write the master — which happens whenever a tool
 * rewrites rows without understanding `si` — and every follower becomes a cell with a
 * cached number and no way to recompute it. Excel repairs this on open, quietly; other
 * readers see a formula-shaped element with nothing in it. This module reports a
 * follower whose master is missing as an error, because the formula is genuinely gone.
 *
 * Verified against the Open XML SDK schema: `x:f/@t` is an enum of exactly
 * `Normal | Array | DataTable | Shared` and is **optional** (absent means normal);
 * `@si` is a `UInt32`; `@ref` carries the master's range; `x:c/@t` is
 * `Boolean | Number | Error | SharedString | String | InlineString | Date`; and
 * `x:calcPr` declares `@calcId`, `@calcMode`, `@fullCalcOnLoad`, `@calcCompleted`,
 * `@calcOnSave` and `@forceFullCalc`.
 *
 * ⚠️ **Nothing here evaluates a formula.** This module reads what is stored and reports
 * what disagrees or is missing; it never claims a cached value is *numerically* wrong,
 * because deciding that would need a calculation engine. Saying "this value may be
 * stale" is honest; saying "this value is 55 but should be 60" would not be.
 */

import { S_NAMESPACE } from './excelStyleResolver';
import { finding, renderFindings, type Finding, type Severity } from './findings';
import type { PackageParts } from './packageIntegrity';

/**
 * Severity and silence per kind.
 *
 * Everything here is SILENT by construction: a spreadsheet with stale values opens,
 * renders, prints and exports without a murmur. The only reason anyone finds out is
 * that a number was wrong in a meeting.
 */
const FORMULA_RULES = {
  'shared-master-missing': { severity: 'error', silent: true },
  'cached-error-value': { severity: 'error', silent: false },
  'stale-by-declaration': { severity: 'warning', silent: true },
  'manual-calculation': { severity: 'warning', silent: true },
  'calculation-incomplete': { severity: 'warning', silent: true },
  'formula-without-value': { severity: 'warning', silent: false },
  'value-without-formula': { severity: 'note', silent: true },
  'volatile-formula': { severity: 'note', silent: true },
  'external-reference': { severity: 'warning', silent: true },
  'array-master-missing-ref': { severity: 'warning', silent: true }
} as const satisfies Record<string, { severity: Severity; silent: boolean }>;

export type FormulaProblemKind = keyof typeof FORMULA_RULES;

const formulaFinding = (
  kind: FormulaProblemKind,
  part: string,
  message: string,
  remediation: string,
  subject?: Record<string, string>
): Finding => finding(`formula/${kind}`, part, message, remediation, { ...FORMULA_RULES[kind], subject });

/**
 * Functions whose result depends on something other than their inputs, so a cached
 * value is stale the moment the file is saved.
 *
 * Not exhaustive and deliberately conservative — these are the ones whose volatility is
 * unambiguous. `INDIRECT` and `OFFSET` are volatile too but appear constantly in
 * legitimate static use, so flagging them would drown the report.
 */
const VOLATILE_FUNCTIONS = ['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'RANDARRAY'];

/** Error literals Excel stores in `<v>` when `@t="e"`. */
const ERROR_VALUES = new Set(['#REF!', '#VALUE!', '#DIV/0!', '#NAME?', '#N/A', '#NULL!', '#NUM!', '#GETTING_DATA', '#SPILL!', '#CALC!']);

export type FormulaKind = 'normal' | 'array' | 'dataTable' | 'shared';

export interface CellFormula {
  /** A1-style reference, e.g. `B2`. */
  cell: string;
  kind: FormulaKind;
  /** Formula text. **Empty for a shared follower** — that is the point, not a defect. */
  text: string;
  /** `@si` for shared formulas. */
  sharedIndex: string | null;
  /** `@ref` — present on a shared or array master, absent on a follower. */
  range: string | null;
  /** The cached result, or null when the cell stores no value. */
  cachedValue: string | null;
  /** `x:c/@t` — `e` means the cached value is an error literal. */
  valueType: string | null;
  /** `@ca`/`@aca`: Excel is told to recalculate this cell always. */
  alwaysCalculate: boolean;
}

export interface CalcSettings {
  fullCalcOnLoad: boolean;
  calcMode: string | null;
  calcCompleted: boolean | null;
  calcId: string | null;
}

const isS = (el: Element, local: string) => el.namespaceURI === S_NAMESPACE && el.localName === local;
const isOn = (v: string | null) => v === '1' || v === 'true';

const parseXml = (xml: string | undefined): Document | null => {
  if (xml === undefined) return null;
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
};

const childOf = (parent: Element, local: string): Element | null =>
  Array.from(parent.children).find(c => isS(c, local)) ?? null;

/**
 * Reads the workbook's calculation settings.
 *
 * These are what turn "the values might be stale" from a general caveat into a specific
 * claim about this file, so they are read before any cell is examined.
 */
export function readCalcSettings(parts: PackageParts): CalcSettings | null {
  const doc = parseXml(parts['xl/workbook.xml']);
  const root = doc?.documentElement;
  if (!root) return null;
  const calcPr = Array.from(root.children).find(c => isS(c, 'calcPr'));
  if (!calcPr) return { fullCalcOnLoad: false, calcMode: null, calcCompleted: null, calcId: null };

  const completed = calcPr.getAttribute('calcCompleted');
  return {
    fullCalcOnLoad: isOn(calcPr.getAttribute('fullCalcOnLoad')),
    calcMode: calcPr.getAttribute('calcMode'),
    calcCompleted: completed === null ? null : isOn(completed),
    calcId: calcPr.getAttribute('calcId')
  };
}

/** Every formula in one worksheet, in document order. */
export function readFormulas(sheetXml: string): CellFormula[] {
  const doc = parseXml(sheetXml);
  const root = doc?.documentElement;
  if (!root) return [];

  const formulas: CellFormula[] = [];
  for (const cell of Array.from(root.querySelectorAll('*'))) {
    if (!isS(cell, 'c')) continue;
    const f = childOf(cell, 'f');
    if (!f) continue;

    const rawKind = (f.getAttribute('t') ?? 'normal').toLowerCase();
    const kind: FormulaKind =
      rawKind === 'array' ? 'array' : rawKind === 'datatable' ? 'dataTable' : rawKind === 'shared' ? 'shared' : 'normal';

    formulas.push({
      cell: cell.getAttribute('r') ?? '',
      kind,
      text: f.textContent ?? '',
      sharedIndex: f.getAttribute('si'),
      range: f.getAttribute('ref'),
      cachedValue: childOf(cell, 'v')?.textContent ?? null,
      valueType: cell.getAttribute('t'),
      alwaysCalculate: isOn(f.getAttribute('ca')) || isOn(f.getAttribute('aca'))
    });
  }
  return formulas;
}

/** Worksheet parts. */
export const FORMULA_HOST_PART = /^xl\/worksheets\/[^/]+\.xml$/;

/**
 * Analyses one worksheet's formulas against the workbook's calculation settings.
 *
 * `calc` is passed in rather than re-read per sheet: the settings are a property of the
 * workbook, and a per-sheet read would repeat the same finding once per worksheet.
 */
export function formulaFindingsForSheet(
  sheetXml: string,
  part: string,
  calc: CalcSettings | null
): Finding[] {
  const formulas = readFormulas(sheetXml);
  if (formulas.length === 0) return [];
  const problems: Finding[] = [];

  // --- shared formulas: a follower with no master has lost its formula ---------
  const masters = new Set(
    formulas.filter(f => f.kind === 'shared' && f.range !== null && f.sharedIndex !== null).map(f => f.sharedIndex!)
  );
  for (const f of formulas) {
    if (f.kind !== 'shared' || f.sharedIndex === null) continue;
    // No separate "skip masters" guard: a master's own si is in `masters` by
    // construction, so the check below already exempts it. An explicit guard here
    // would read as protection and provably never fire.
    if (masters.has(f.sharedIndex)) continue;

    problems.push(formulaFinding(
      'shared-master-missing', part,
      `Cell ${f.cell} is a shared-formula follower for si="${f.sharedIndex}", but no cell in this sheet carries the master (a shared formula with a ref attribute and the formula text). A follower stores no formula text of its own — its formula exists only as an offset from the master — so the formula is gone. The cell still displays its cached value ${f.cachedValue === null ? '' : `of ${f.cachedValue} `}and nothing looks wrong. Excel repairs this quietly on open; other readers see an empty formula element.`,
      `Restore the master cell for si="${f.sharedIndex}", or rewrite ${f.cell} with its own explicit formula text.`,
      { cell: f.cell, si: f.sharedIndex }
    ));
  }

  // --- array masters must declare their range ---------------------------------
  for (const f of formulas) {
    if (f.kind === 'array' && f.range === null) {
      problems.push(formulaFinding(
        'array-master-missing-ref', part,
        `Cell ${f.cell} holds an array formula with no ref attribute, so the range it spills into is undeclared. Excel infers one; another reader has nothing to infer from and may apply the formula to a single cell.`,
        `Add a ref attribute naming the range this array formula occupies.`,
        { cell: f.cell }
      ));
    }
  }

  // --- per-cell observations ---------------------------------------------------
  for (const f of formulas) {
    if (f.valueType === 'e' && f.cachedValue !== null && ERROR_VALUES.has(f.cachedValue)) {
      problems.push(formulaFinding(
        'cached-error-value', part,
        `Cell ${f.cell} has a stored error value of ${f.cachedValue}. Unlike most findings here this one IS visible — the error text is what a reader sees — but it is frozen in the file, so every export and every conversion carries it forward.`,
        `Fix the formula in ${f.cell}. ${f.cachedValue === '#REF!' ? 'A #REF! specifically means the formula points at cells that were deleted; the original reference is not recoverable from the file.' : ''}`.trim(),
        { cell: f.cell, error: f.cachedValue }
      ));
    }

    if (f.cachedValue === null && f.kind !== 'shared') {
      problems.push(formulaFinding(
        'formula-without-value', part,
        `Cell ${f.cell} holds a formula but no cached value, so any reader that does not calculate shows an empty cell where a number belongs.`,
        'Open and save the workbook in Excel to populate the cached values, or accept that non-calculating readers render this cell empty.',
        { cell: f.cell }
      ));
    }

    const upper = f.text.toUpperCase();
    const volatile = VOLATILE_FUNCTIONS.find(fn => upper.includes(`${fn}(`));
    if (volatile) {
      problems.push(formulaFinding(
        'volatile-formula', part,
        `Cell ${f.cell} uses ${volatile}(), whose result depends on when it is calculated rather than on its inputs. Its cached value was correct at save time and is almost certainly wrong now.`,
        'No action needed if the workbook is opened in Excel, which recalculates it. A converter should treat this value as meaningless rather than as data.',
        { cell: f.cell, function: volatile }
      ));
    }

    // [1]Sheet1!A1 - the bracketed index refers to an entry in externalLinks.
    if (/\[\d+\]/.test(f.text)) {
      problems.push(formulaFinding(
        'external-reference', part,
        `Cell ${f.cell} references another workbook. The value shown is a cached copy of that workbook's data from the last time both were open together — the source is outside this package, so nothing here can confirm it is current, or that it still exists.`,
        'Confirm the source workbook is available to whoever opens this file, or replace the link with static values.',
        { cell: f.cell }
      ));
    }
  }

  // --- workbook-level calculation state ----------------------------------------
  if (calc?.fullCalcOnLoad) {
    problems.push(formulaFinding(
      'stale-by-declaration', part,
      `The workbook is marked fullCalcOnLoad, so Excel will recalculate everything when it opens. The values stored in this file are stale by Excel's own declaration — anything reading them without a calculation engine is reading numbers Excel has already disowned.`,
      'No action needed in Excel. A converter or extractor should treat every cached value in this workbook as provisional.'
    ));
  }
  if (calc?.calcMode === 'manual') {
    problems.push(formulaFinding(
      'manual-calculation', part,
      'Automatic calculation is turned off for this workbook. Values drift from their formulas by design, and Excel itself will not correct them until someone presses F9 — so a cell can display a number its own formula would not produce, indefinitely.',
      'Set calcMode to auto, or confirm that manual calculation is deliberate and that consumers know the values are not live.'
    ));
  }
  if (calc?.calcCompleted === false) {
    problems.push(formulaFinding(
      'calculation-incomplete', part,
      'The last calculation of this workbook did not finish, so some cached values are from a partial pass and may not be self-consistent.',
      'Open in Excel and allow calculation to complete, then save.'
    ));
  }

  return problems;
}

/** Every formula finding across the workbook. */
export function formulaFindings(parts: PackageParts): Finding[] {
  const calc = readCalcSettings(parts);
  const sheets = Object.keys(parts).filter(p => FORMULA_HOST_PART.test(p));
  const problems: Finding[] = [];

  for (const [index, path] of sheets.entries()) {
    // Workbook-level settings belong to the workbook, so report them once - against
    // the first sheet - rather than once per worksheet.
    problems.push(...formulaFindingsForSheet(parts[path], path, index === 0 ? calc : null));
  }
  return problems;
}

/**
 * Evidence lines for the AI panel.
 *
 * Leads with the cache framing, because every other statement depends on the reader
 * understanding that the numbers are stored rather than computed.
 */
export function computeFormulaEvidenceForMarkup(
  parts: Record<string, string>
): { lines: string[]; unresolved: string[] } | null {
  const sheetPath = Object.keys(parts).find(p => FORMULA_HOST_PART.test(p));
  if (sheetPath === undefined) return null;

  const formulas = readFormulas(parts[sheetPath]);
  const calc = readCalcSettings(parts);
  const problems = formulaFindingsForSheet(parts[sheetPath], sheetPath, calc);
  if (formulas.length === 0) return null;

  const lines: string[] = [];
  const unresolved: string[] = [];

  const shared = formulas.filter(f => f.kind === 'shared').length;
  lines.push(
    `${sheetPath} contains ${formulas.length} formula(s)${shared > 0 ? `, ${shared} of them shared` : ''}. ` +
      'Each cell stores both the formula and the value from the last recalculation; every reader without a calculation engine displays the stored value.'
  );

  if (calc) {
    lines.push(
      `Workbook calculation: mode ${calc.calcMode ?? 'automatic (default)'}` +
        `${calc.fullCalcOnLoad ? ', marked to fully recalculate on load' : ''}` +
        `${calc.calcId ? `, last calculated by engine ${calc.calcId}` : ''}.`
    );
  }

  lines.push(...renderFindings(problems));

  // The honest limit, and the reason this module never says a number is wrong.
  unresolved.push(
    'No formula was evaluated. Whether a cached value equals what its formula would now produce cannot be determined without a calculation engine, which this engine does not have.'
  );
  if (formulas.some(f => /\[\d+\]/.test(f.text))) {
    unresolved.push(
      'Formulas reference an external workbook whose contents are outside this package, so the cached values taken from it cannot be checked at all.'
    );
  }

  return { lines, unresolved };
}
