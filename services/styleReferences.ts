/**
 * Dangling style and format references — the fallback that looks like a decision.
 *
 * Every formatting system in OOXML is a table of definitions plus references into it,
 * and **every one of them falls back silently when a reference misses**:
 *
 *   - A `w:pStyle` naming a style that is not in `styles.xml` does not error. Word
 *     applies **Normal**. A paragraph that was supposed to be Heading 1 renders as body
 *     text — right font, right size, wrong meaning — and the document opens cleanly.
 *   - A cell whose `@s` indexes past the end of `cellXfs` gets the default format.
 *   - A `@numFmtId` with no matching `numFmt` falls back to General, so a currency
 *     column quietly renders as bare numbers.
 *
 * None of this is visible as breakage. It is visible as *plainer formatting*, which
 * looks like a design choice rather than a fault — and that is precisely why it survives
 * review. It is also the single most common defect in **generated** documents, because a
 * generator that writes a style reference and forgets the style definition produces a
 * file that opens, renders, and passes every structural check ever written for it.
 *
 * WHY THIS IS A SEPARATE ANALYZER RATHER THAN PART OF THE RESOLVERS.
 *
 * `wordStyleResolver` and `excelStyleResolver` answer *"what does this element resolve
 * to?"* — a question about one element, asked when someone selects it. This asks *"does
 * every reference in this package land somewhere?"* — a question about the whole file,
 * asked when nobody is looking at anything in particular. The resolvers were `explain`
 * only and contributed no findings at all; this is the missing half.
 *
 * ⚠️ **`w:numId="0"` IS NOT A DANGLING REFERENCE.** The specification is explicit that 0
 * "shall never be used to point to a numbering definition instance" — it means *remove
 * numbering from this paragraph*. Reporting it as a broken lookup would fire on every
 * document that has ever had a list removed from it, which is most of them.
 *
 * ⚠️ **Excel number formats 0–163 are built in and are not declared anywhere.** Only ids
 * **164 and above** are custom and must appear in `numFmts`. Checking every id against
 * the table would report every ordinary workbook as broken.
 */

import { W_NAMESPACE, parseStyles, type StyleSheet } from './wordStyleResolver';
import { S_NAMESPACE, parseExcelStyles, BUILT_IN_NUMBER_FORMATS, type ExcelStyleSheet } from './excelStyleResolver';
import { finding, renderFindings, type Finding, type Severity } from './findings';
import type { PackageParts } from './packageIntegrity';

/**
 * The lowest `numFmtId` that must be declared.
 *
 * Ids below this are built into Excel and appear in no part of the file.
 */
export const FIRST_CUSTOM_NUMBER_FORMAT_ID = 164;

/**
 * Severity and silence per kind.
 *
 * All silent. Every one of these resolves to a default that renders perfectly well —
 * the document is not broken-looking, it is *plainer than intended*, which nobody
 * reports as a bug.
 */
const REFERENCE_RULES = {
  'missing-paragraph-style': { severity: 'error', silent: true },
  'missing-character-style': { severity: 'error', silent: true },
  'missing-table-style': { severity: 'error', silent: true },
  'missing-numbering': { severity: 'error', silent: true },
  'missing-based-on': { severity: 'warning', silent: true },
  'missing-linked-style': { severity: 'warning', silent: true },
  'no-stylesheet': { severity: 'error', silent: true },
  'cell-format-out-of-range': { severity: 'error', silent: true },
  'missing-number-format': { severity: 'error', silent: true },
  'component-out-of-range': { severity: 'error', silent: true }
} as const satisfies Record<string, { severity: Severity; silent: boolean }>;

export type StyleReferenceProblemKind = keyof typeof REFERENCE_RULES;

const referenceFinding = (
  kind: StyleReferenceProblemKind,
  part: string,
  message: string,
  remediation: string,
  subject?: Record<string, string>
): Finding => finding(`styleRef/${kind}`, part, message, remediation, { ...REFERENCE_RULES[kind], subject });

const isW = (el: Element, local: string) => el.namespaceURI === W_NAMESPACE && el.localName === local;
const wAttr = (el: Element, name: string) => el.getAttributeNS(W_NAMESPACE, name);

const parseXml = (xml: string | undefined): Document | null => {
  if (xml === undefined) return null;
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
};

/** Word style references, keyed by the element that carries them. */
const WORD_STYLE_REFERENCES: Record<string, StyleReferenceProblemKind> = {
  pStyle: 'missing-paragraph-style',
  rStyle: 'missing-character-style',
  tblStyle: 'missing-table-style'
};

const WORD_REFERENCE_LABEL: Record<string, string> = {
  pStyle: 'paragraph style',
  rStyle: 'character style',
  tblStyle: 'table style'
};

/** What Word falls back to when each kind of reference misses. */
const WORD_FALLBACK: Record<string, string> = {
  pStyle: 'Word applies the Normal style instead, so the paragraph renders as ordinary body text',
  rStyle: 'Word applies no character style, so the run renders with only its paragraph’s formatting',
  tblStyle: 'Word applies no table style, so the table loses its banding, borders and header formatting'
};

/**
 * Checks every style and numbering reference in one Word body part.
 *
 * `numbering` may be null when the package has no `numbering.xml`; in that case
 * numbering references are not judged, because a missing part is `packageIntegrity`'s
 * finding to make and reporting it here as well would double-count one fault.
 */
export function checkWordReferences(
  doc: Document,
  part: string,
  sheet: StyleSheet,
  numbering: Document | null
): Finding[] {
  const problems: Finding[] = [];
  const root = doc.documentElement;
  if (!root) return problems;

  const seen = new Set<string>();

  for (const el of Array.from(root.querySelectorAll('*'))) {
    const kind = WORD_STYLE_REFERENCES[el.localName];
    if (kind === undefined || el.namespaceURI !== W_NAMESPACE) continue;

    const id = wAttr(el, 'val');
    if (id === null || sheet.styles.has(id)) continue;

    // One finding per missing style, not one per use: a heading style referenced by
    // forty paragraphs is one broken style, and forty identical findings would bury
    // everything else in the report.
    const key = `${kind}:${id}`;
    if (seen.has(key)) continue;
    seen.add(key);

    problems.push(referenceFinding(
      kind, part,
      `A ${WORD_REFERENCE_LABEL[el.localName]} named "${id}" is referenced but not defined in styles.xml. This does not fail — ${WORD_FALLBACK[el.localName]}. The document opens cleanly and simply looks plainer than intended, which reads as a design choice rather than a fault.`,
      `Add a w:style with w:styleId="${id}" to styles.xml, or change the reference to a style that exists.`,
      { styleId: id }
    ));
  }

  // --- numbering -------------------------------------------------------------
  if (numbering?.documentElement) {
    const declared = new Set(
      Array.from(numbering.documentElement.querySelectorAll('*'))
        .filter(el => isW(el, 'num'))
        .map(el => wAttr(el, 'numId'))
        .filter((v): v is string => v !== null)
    );

    for (const el of Array.from(root.querySelectorAll('*'))) {
      if (!isW(el, 'numId')) continue;
      const id = wAttr(el, 'val');
      if (id === null) continue;
      // 0 means REMOVE numbering, not "look up definition 0". Treating it as a lookup
      // reports every document that has ever had a list removed from it.
      if (id === '0' || declared.has(id)) continue;
      const key = `num:${id}`;
      if (seen.has(key)) continue;
      seen.add(key);

      problems.push(referenceFinding(
        'missing-numbering', part,
        `A paragraph references numbering definition ${id}, which numbering.xml does not declare. The paragraph renders with no list formatting at all — no number, no bullet, no indent — while still claiming to be a list item.`,
        `Add a w:num with w:numId="${id}" to numbering.xml, or remove the w:numPr from the paragraphs referencing it.`,
        { numId: id }
      ));
    }
  }

  return problems;
}

/**
 * Checks the stylesheet's internal references.
 *
 * A style whose `basedOn` is missing does not inherit — it silently becomes a root
 * style, so everything it was supposed to pick up from its parent is simply absent.
 */
export function checkStylesheetIntegrity(sheet: StyleSheet, part: string): Finding[] {
  const problems: Finding[] = [];

  for (const [id, definition] of sheet.styles) {
    const el = definition.element;
    if (!el) continue;

    for (const [child, kind, consequence] of [
      ['basedOn', 'missing-based-on', 'so it inherits nothing and silently becomes a root style — every property it was meant to pick up from its parent is simply absent'],
      ['link', 'missing-linked-style', 'so the paired paragraph/character relationship is broken and applying one will not apply the other']
    ] as const) {
      const ref = Array.from(el.children).find(c => isW(c, child));
      const target = ref ? wAttr(ref, 'val') : null;
      if (target === null || sheet.styles.has(target)) continue;

      problems.push(referenceFinding(
        kind, part,
        `Style "${id}" declares w:${child}="${target}", which is not defined, ${consequence}.`,
        `Define a style with w:styleId="${target}", or remove the w:${child} from "${id}".`,
        { styleId: id, target }
      ));
    }
  }

  return problems;
}

/** Checks every format reference in one worksheet against the workbook stylesheet. */
export function checkExcelReferences(doc: Document, part: string, sheet: ExcelStyleSheet): Finding[] {
  const problems: Finding[] = [];
  const root = doc.documentElement;
  if (!root) return problems;

  const seen = new Set<string>();

  for (const cell of Array.from(root.querySelectorAll('*'))) {
    if (cell.namespaceURI !== S_NAMESPACE || cell.localName !== 'c') continue;
    const raw = cell.getAttribute('s');
    if (raw === null) continue;
    const index = Number.parseInt(raw, 10);
    if (Number.isNaN(index) || (index >= 0 && index < sheet.cellXfs.length)) continue;

    const key = `xf:${raw}`;
    if (seen.has(key)) continue;
    seen.add(key);

    problems.push(referenceFinding(
      'cell-format-out-of-range', part,
      `Cell ${cell.getAttribute('r') ?? '(unnamed)'} uses format index ${raw}, but cellXfs declares only ${sheet.cellXfs.length} format(s) (valid indices 0–${Math.max(sheet.cellXfs.length - 1, 0)}). Excel falls back to the default format, so the cell renders as unstyled General — no currency symbol, no date formatting, no fill — and looks merely plain rather than wrong.`,
      `Add the missing cellXfs entries, or repoint the cell at an index that exists.`,
      { index: raw }
    ));
  }

  return problems;
}

/**
 * Checks the workbook stylesheet's own internal references.
 *
 * `cellXfs` entries index into the fonts, fills and borders tables, and a `numFmtId`
 * at or above 164 must be declared in `numFmts`.
 */
export function checkExcelStylesheetIntegrity(sheet: ExcelStyleSheet, part: string): Finding[] {
  const problems: Finding[] = [];
  const seen = new Set<string>();

  const components: Array<[string, Element[], string]> = [
    ['fontId', sheet.fonts, 'font'],
    ['fillId', sheet.fills, 'fill'],
    ['borderId', sheet.borders, 'border']
  ];

  sheet.cellXfs.forEach((xf, position) => {
    for (const [attribute, table, label] of components) {
      const raw = xf.getAttribute(attribute);
      if (raw === null) continue;
      const index = Number.parseInt(raw, 10);
      if (Number.isNaN(index) || (index >= 0 && index < table.length)) continue;

      const key = `${attribute}:${raw}`;
      if (seen.has(key)) continue;
      seen.add(key);

      problems.push(referenceFinding(
        'component-out-of-range', part,
        `Cell format ${position} references ${label} ${raw}, but only ${table.length} ${label}(s) are declared. Excel substitutes the default, so every cell using this format loses that aspect of its styling and renders as if it had never been set.`,
        `Add the missing ${label} entries to styles.xml, or repoint the format at an index that exists.`,
        { format: String(position), [attribute]: raw }
      ));
    }

    const rawNumFmt = xf.getAttribute('numFmtId');
    if (rawNumFmt === null) return;
    const numFmtId = Number.parseInt(rawNumFmt, 10);
    if (Number.isNaN(numFmtId)) return;
    // Ids below 164 are built into Excel and declared nowhere; only custom ids must
    // appear in numFmts. Checking all of them reports every ordinary workbook.
    if (numFmtId < FIRST_CUSTOM_NUMBER_FORMAT_ID) return;
    if (sheet.numFmts.has(numFmtId)) return;

    const key = `numFmt:${rawNumFmt}`;
    if (seen.has(key)) return;
    seen.add(key);

    problems.push(referenceFinding(
      'missing-number-format', part,
      `Cell format ${position} references custom number format ${rawNumFmt}, which is not declared in numFmts. Excel falls back to General, so a column of currency or dates renders as bare numbers — readable, wrong, and easy to mistake for the intended formatting.`,
      `Declare a numFmt with numFmtId="${rawNumFmt}", or repoint the format at a built-in id below ${FIRST_CUSTOM_NUMBER_FORMAT_ID}.`,
      { format: String(position), numFmtId: rawNumFmt }
    ));
  });

  return problems;
}

const WORD_BODY = /^word\/(?:document\d*|header[^/]*|footer[^/]*|footnotes\d*|endnotes\d*)\.xml$/;
const EXCEL_SHEET = /^xl\/worksheets\/[^/]+\.xml$/;

/** True when this package has a stylesheet worth checking references against. */
export const hasStyleReferences = (parts: PackageParts): boolean =>
  parts['word/styles.xml'] !== undefined || parts['xl/styles.xml'] !== undefined;

/** Every dangling style or format reference in the package. */
export function styleReferenceFindings(parts: PackageParts): Finding[] {
  const problems: Finding[] = [];

  const wordStylesXml = parts['word/styles.xml'];
  if (wordStylesXml !== undefined) {
    const sheet = parseStyles(wordStylesXml);
    const numbering = parseXml(parts['word/numbering.xml']);
    problems.push(...checkStylesheetIntegrity(sheet, 'word/styles.xml'));
    for (const path of Object.keys(parts).filter(p => WORD_BODY.test(p))) {
      const doc = parseXml(parts[path]);
      if (doc) problems.push(...checkWordReferences(doc, path, sheet, numbering));
    }
  }

  const excelStylesXml = parts['xl/styles.xml'];
  if (excelStylesXml !== undefined) {
    const sheet = parseExcelStyles(excelStylesXml);
    problems.push(...checkExcelStylesheetIntegrity(sheet, 'xl/styles.xml'));
    for (const path of Object.keys(parts).filter(p => EXCEL_SHEET.test(p))) {
      const doc = parseXml(parts[path]);
      if (doc) problems.push(...checkExcelReferences(doc, path, sheet));
    }
  }

  return problems;
}

/**
 * Evidence lines for the AI panel.
 *
 * Says how many definitions exist as well as how many references miss, because
 * "3 references to styles that are not defined, out of a stylesheet of 12" is a
 * different situation from the same three out of two hundred.
 */
export function computeStyleReferenceEvidenceForMarkup(
  parts: Record<string, string>
): { lines: string[]; unresolved: string[] } | null {
  if (!hasStyleReferences(parts)) return null;
  const problems = styleReferenceFindings(parts);

  const lines: string[] = [];
  const unresolved: string[] = [];

  const wordStylesXml = parts['word/styles.xml'];
  if (wordStylesXml !== undefined) {
    lines.push(`word/styles.xml defines ${parseStyles(wordStylesXml).styles.size} style(s).`);
  }
  const excelStylesXml = parts['xl/styles.xml'];
  if (excelStylesXml !== undefined) {
    const sheet = parseExcelStyles(excelStylesXml);
    lines.push(
      `xl/styles.xml declares ${sheet.cellXfs.length} cell format(s), ${sheet.fonts.length} font(s), ` +
        `${sheet.fills.length} fill(s), ${sheet.borders.length} border(s) and ${sheet.numFmts.size} custom number format(s). ` +
        `Number format ids below ${FIRST_CUSTOM_NUMBER_FORMAT_ID} are built into Excel and are declared nowhere — ` +
        `${BUILT_IN_NUMBER_FORMATS.size} of them are known to this engine.`
    );
  }

  if (problems.length === 0) {
    lines.push('Every style and format reference in this package resolves to a definition that exists.');
  } else {
    lines.push(
      `${problems.length} reference(s) do not resolve. Each falls back to a default rather than failing, so the document renders as plainer than intended rather than as broken.`
    );
    lines.push(...renderFindings(problems));
  }

  // Only the parts supplied were checked, and the checker cannot see a style that
  // lives in a stylesheet it was not given.
  unresolved.push(
    'References were checked only against the stylesheets supplied with this request. A style defined in a part not loaded here would be reported as missing when it is not.'
  );

  return { lines, unresolved };
}
