/**
 * Composes the WordprocessingML cascade over a whole package.
 *
 * `wordStyleResolver`, `wordNumbering` and `wordTableStyles` each solve one piece and
 * deliberately take their inputs as arguments, which keeps them pure and testable but
 * leaves nobody able to actually use them. This module is the join: give it the parts
 * of a `.docx` and a paragraph, and it answers what Word will render and why.
 *
 * The output is designed to be handed to a language model as **pre-verified evidence**
 * rather than retrieved context. Every line is computed from the document, so a caller
 * can present it as verified instead of merely grounded — and, just as importantly, the
 * `unresolved` list says where the computation stopped, so an answer built on it can
 * decline to assert what was not established.
 */

import {
  parseStyles,
  resolveRunProperties,
  resolveParagraphProperties,
  explainResolution,
  W_NAMESPACE,
  type StyleSheet,
  type ResolvedFormatting
} from './wordStyleResolver';
import {
  parseNumbering,
  resolveNumbering,
  readNumberingReference,
  type NumberingDefinitions,
  type ResolvedNumbering
} from './wordNumbering';
import {
  readTableLook,
  readBandSizes,
  readConditionalBlocks,
  applicableConditionalFormats,
  type ConditionalFormatType
} from './wordTableStyles';
import { resolveAlternateContent, MODERN_CONSUMER_NAMESPACES } from './markupCompatibility';
import type { PackageParts } from './packageIntegrity';

const STYLES_PART = 'word/styles.xml';
const NUMBERING_PART = 'word/numbering.xml';
const DOCUMENT_PART = 'word/document.xml';

export interface WordDocumentContext {
  styles: StyleSheet;
  numbering: NumberingDefinitions | null;
  /** `word/document.xml`, with markup compatibility already resolved. */
  document: Document | null;
  /** Parts that were expected but absent, so a caller can say what it could not use. */
  unresolved: string[];
}

/**
 * Loads the parts the cascade needs.
 *
 * Missing parts are recorded rather than thrown: a document with no `numbering.xml` is
 * perfectly normal, and a document with no `styles.xml` still resolves against
 * document defaults. What is not acceptable is silently reporting formatting as
 * complete when a part it depends on was absent.
 */
export const loadWordContext = (parts: PackageParts): WordDocumentContext => {
  const unresolved: string[] = [];

  const stylesXml = parts[STYLES_PART];
  if (stylesXml === undefined) {
    unresolved.push(`${STYLES_PART} is not in the package; only direct formatting can be resolved`);
  }
  const styles = stylesXml
    ? parseStyles(stylesXml)
    : { docDefaults: {}, styles: new Map() };

  const numberingXml = parts[NUMBERING_PART];
  const numbering = numberingXml ? parseNumbering(numberingXml) : null;
  if (!numberingXml) {
    unresolved.push(`${NUMBERING_PART} is not in the package; list formatting cannot be resolved`);
  }

  const documentXml = parts[DOCUMENT_PART];
  let document: Document | null = null;
  if (documentXml === undefined) {
    unresolved.push(`${DOCUMENT_PART} is not in the package`);
  } else {
    const parsed = new DOMParser().parseFromString(documentXml, 'application/xml');
    if (parsed.getElementsByTagName('parsererror').length > 0) {
      unresolved.push(`${DOCUMENT_PART} is not well-formed XML`);
    } else {
      // Resolve mc:AlternateContent before anything walks the tree, or a shape
      // written twice is counted twice. See ./markupCompatibility.
      document = resolveAlternateContent(parsed, MODERN_CONSUMER_NAMESPACES).document;
    }
  }

  return { styles, numbering, document, unresolved };
};

const childOf = (parent: Element | null | undefined, localName: string): Element | null => {
  if (!parent) return null;
  return Array.from(parent.children).find(
    el => el.namespaceURI === W_NAMESPACE && el.localName === localName
  ) ?? null;
};

const valOf = (element: Element | null): string | null =>
  element?.getAttributeNS(W_NAMESPACE, 'val') ?? null;

/** Walks up from an element to the nearest ancestor with the given local name. */
const ancestor = (element: Element, localName: string): Element | null => {
  let current: Element | null = element.parentElement;
  while (current) {
    if (current.namespaceURI === W_NAMESPACE && current.localName === localName) return current;
    current = current.parentElement;
  }
  return null;
};

/**
 * Determines a cell's position within its table so table-style conditional formatting
 * can be gated correctly.
 *
 * Returns null when the paragraph is not in a table, which is the common case and not
 * an error.
 */
const tablePositionOf = (paragraph: Element): {
  tbl: Element;
  row: number;
  col: number;
  rowCount: number;
  colCount: number;
} | null => {
  const tc = ancestor(paragraph, 'tc');
  const tr = tc ? ancestor(tc, 'tr') : null;
  const tbl = tr ? ancestor(tr, 'tbl') : null;
  if (!tc || !tr || !tbl) return null;

  const rows = Array.from(tbl.children).filter(
    el => el.namespaceURI === W_NAMESPACE && el.localName === 'tr'
  );
  const cells = Array.from(tr.children).filter(
    el => el.namespaceURI === W_NAMESPACE && el.localName === 'tc'
  );

  return {
    tbl,
    row: rows.indexOf(tr),
    col: cells.indexOf(tc),
    rowCount: rows.length,
    colCount: cells.length
  };
};

export interface FormattingAnalysis {
  /** Effective paragraph properties. */
  paragraph: ResolvedFormatting;
  /** Effective run properties, when a run was supplied. */
  run: ResolvedFormatting | null;
  /** The resolved numbering level, when the paragraph is numbered. */
  numbering: ResolvedNumbering | null;
  /** Conditional formats applied, when the paragraph is in a table. */
  tableFormats: ConditionalFormatType[];
  /** Everything the analysis could not establish, in plain language. */
  unresolved: string[];
  /** Ordered, human-readable account suitable for handing to a model verbatim. */
  explanation: string[];
}

/**
 * Resolves the effective formatting of a paragraph, and optionally a run inside it.
 *
 * This is the answer to "will this render the way I want?" — computed, not retrieved.
 */
export const analyzeParagraphFormatting = (
  context: WordDocumentContext,
  paragraph: Element,
  run?: Element
): FormattingAnalysis => {
  const unresolved = [...context.unresolved];

  const pPr = childOf(paragraph, 'pPr');
  const paragraphStyleId = valOf(childOf(pPr, 'pStyle')) ?? undefined;
  const directRPr = run ? childOf(run, 'rPr') : null;
  const characterStyleId = valOf(childOf(directRPr, 'rStyle')) ?? undefined;

  // --- Layer 3: numbering -------------------------------------------------
  // Direct numbering on the paragraph overrides numbering inherited from its style,
  // so only fall back to the style when the paragraph says nothing.
  let numberingRef = readNumberingReference(pPr);
  if (numberingRef.numId === null && paragraphStyleId) {
    const style = context.styles.styles.get(paragraphStyleId);
    numberingRef = readNumberingReference(style?.pPr ?? null);
  }

  let numbering: ResolvedNumbering | null = null;
  if (numberingRef.numId !== null) {
    if (!context.numbering) {
      unresolved.push(`paragraph references numId ${numberingRef.numId} but the numbering part is missing`);
    } else {
      numbering = resolveNumbering(context.numbering, context.styles, numberingRef);
      if (numbering === null && numberingRef.numId === '0') {
        // Not a failure: numId 0 is an explicit instruction to remove numbering.
      } else if (numbering && numbering.lvl === null) {
        unresolved.push(`numbering level ${numberingRef.ilvl ?? 0} of numId ${numberingRef.numId} could not be resolved`);
      }
      for (const problem of numbering?.problems ?? []) unresolved.push(problem);
    }
  }

  // --- Layer 2: table style conditional formatting -------------------------
  const position = tablePositionOf(paragraph);
  let tableFormats: ConditionalFormatType[] = [];
  let tableStyleLayers: { type: ConditionalFormatType; pPr?: Element; rPr?: Element }[] | undefined;

  if (position) {
    const tblPr = childOf(position.tbl, 'tblPr');
    const tableStyleId = valOf(childOf(tblPr, 'tblStyle'));
    const styleDefinition = tableStyleId ? context.styles.styles.get(tableStyleId) : undefined;

    if (tableStyleId && !styleDefinition) {
      unresolved.push(`table references style "${tableStyleId}", which is not defined in styles.xml`);
    }

    // Band sizes live on the *style's* tblPr, not the table's.
    const styleElement = styleDefinition?.element ?? null;
    const bands = readBandSizes(childOf(styleElement, 'tblPr'));
    for (const note of bands.notes) unresolved.push(`${tableStyleId ?? 'table style'}: ${note}`);

    const look = readTableLook(tblPr);
    tableFormats = applicableConditionalFormats(look, bands, position);

    const blocks = readConditionalBlocks(styleElement);
    tableStyleLayers = tableFormats.map(type => {
      const block = blocks.get(type);
      return { type, pPr: childOf(block ?? null, 'pPr') ?? undefined, rPr: childOf(block ?? null, 'rPr') ?? undefined };
    });
  }

  const paragraphResult = resolveParagraphProperties(context.styles, {
    paragraphStyleId,
    directPPr: pPr ?? undefined,
    insideTable: position !== null,
    tableStyle: tableStyleLayers,
    numbering
  });

  const runResult = run
    ? resolveRunProperties(context.styles, {
        paragraphStyleId,
        characterStyleId,
        directRPr: directRPr ?? undefined,
        insideTable: position !== null,
        tableStyle: tableStyleLayers,
        numbering
      })
    : null;

  const explanation: string[] = [];
  if (position) {
    explanation.push(
      `Paragraph is in a table at row ${position.row + 1} of ${position.rowCount}, column ${position.col + 1} of ${position.colCount}.`,
      `Conditional formats applied, in Word's order: ${tableFormats.join(' → ')}.`
    );
  }
  if (numbering) {
    explanation.push(
      `Numbered: numId ${numbering.numId}, level ${numbering.ilvl}` +
      (numbering.numFmt ? `, format ${numbering.numFmt}` : '') +
      (numbering.lvlText ? `, pattern "${numbering.lvlText}"` : '') + '.'
    );
  }
  explanation.push('Paragraph properties:', ...explainResolution(paragraphResult).map(l => `  ${l}`));
  if (runResult) {
    explanation.push('Run properties:', ...explainResolution(runResult).map(l => `  ${l}`));
  }
  if (unresolved.length > 0) {
    explanation.push(
      'Not established by this analysis (do not assert these):',
      ...unresolved.map(u => `  ${u}`)
    );
  }

  return {
    paragraph: paragraphResult,
    run: runResult,
    numbering,
    tableFormats,
    unresolved,
    explanation
  };
};

/**
 * Convenience entry point: analyze the nth paragraph of a loaded package.
 *
 * Paragraph order is document order after markup-compatibility resolution, so the
 * index is stable against the same input.
 */
export const analyzeParagraphAt = (
  context: WordDocumentContext,
  index: number
): FormattingAnalysis | null => {
  if (!context.document) return null;
  const paragraphs = Array.from(
    context.document.getElementsByTagNameNS(W_NAMESPACE, 'p')
  );
  const paragraph = paragraphs[index];
  if (!paragraph) return null;
  const firstRun = Array.from(paragraph.children).find(
    el => el.namespaceURI === W_NAMESPACE && el.localName === 'r'
  );
  return analyzeParagraphFormatting(context, paragraph, firstRun);
};
