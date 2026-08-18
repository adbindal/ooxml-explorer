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

/**
 * Matches the parts of a `.docx` that carry `w:p` content.
 *
 * Body text is only one of several stories in a Word file: headers, footers, footnotes,
 * endnotes and comments are sibling parts holding paragraphs of their own. All of them
 * resolve against the *same* `word/styles.xml` and `word/numbering.xml`, so nothing
 * about the cascade changes — but a part that is never loaded can never be searched,
 * and a selection inside one silently produces no evidence at all.
 *
 * Matched by shape rather than by a fixed list of names. The `header1.xml` numbering is
 * only a producer convention; the part name a relationship actually points at is free
 * to be anything, and a file that uses an unusual but legal name should still resolve
 * rather than fall back to guesswork.
 *
 * Two deliberate exclusions. Sub-folders (`word/glossary/document.xml`) are left out
 * because a glossary document has its *own* `word/glossary/styles.xml`, so resolving it
 * against the main stylesheet would report formatting Word never applies. And the
 * `comments` family is held to a digits-only suffix so that the w15/w16 side-cars —
 * `commentsExtended.xml`, `commentsIds.xml` — stay out; they carry no paragraphs and
 * would only add parse cost.
 */
const BODY_PART_PATTERN =
  /^word\/(?:document\d*|header[^/]*|footer[^/]*|footnotes\d*|endnotes\d*|comments\d*)\.xml$/;

/** A story within the document, with markup compatibility already resolved. */
export interface WordBodyPart {
  /** The part path as it appears in the package, e.g. `word/header1.xml`. */
  path: string;
  document: Document;
}

export interface WordDocumentContext {
  styles: StyleSheet;
  numbering: NumberingDefinitions | null;
  /**
   * `word/document.xml` specifically, kept for callers that mean the main story.
   * Null when the package has no such part, even if other body parts loaded.
   */
  document: Document | null;
  /**
   * Every body part that parsed, `word/document.xml` first and the rest in path order.
   * Ordering is fixed so that anything indexing into it is stable across runs.
   */
  bodyParts: WordBodyPart[];
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

  // A document with no lists has no numbering part, and that is entirely normal - not
  // something left unresolved. Flagging it unconditionally capped every list-free
  // document below the Verified tier over a part nothing referenced. The genuine case,
  // a paragraph that references a numId when the part is absent, is reported per
  // paragraph in analyzeParagraphFormatting where the reference is actually seen.
  const numberingXml = parts[NUMBERING_PART];
  const numbering = numberingXml ? parseNumbering(numberingXml) : null;

  const bodyPaths = Object.keys(parts)
    .filter(path => BODY_PART_PATTERN.test(path))
    .sort((a, b) => (a === DOCUMENT_PART ? -1 : b === DOCUMENT_PART ? 1 : a.localeCompare(b)));

  const bodyParts: WordBodyPart[] = [];
  for (const path of bodyPaths) {
    const parsed = new DOMParser().parseFromString(parts[path], 'application/xml');
    if (parsed.getElementsByTagName('parsererror').length > 0) {
      // A malformed header is not fatal to the rest of the analysis, but it must be
      // reported: it caps the tier below Verified rather than letting an answer rest on
      // a story that was quietly skipped.
      unresolved.push(`${path} is not well-formed XML`);
      continue;
    }
    // Resolve mc:AlternateContent before anything walks the tree, or a shape
    // written twice is counted twice. Applies to every story, not just the body —
    // a text box in a header is written exactly the same way. See ./markupCompatibility.
    bodyParts.push({ path, document: resolveAlternateContent(parsed, MODERN_CONSUMER_NAMESPACES).document });
  }

  // Only a total absence of stories is worth reporting. A missing `document.xml` on its
  // own is not: a header paragraph resolves against `styles.xml` and `numbering.xml`
  // alone, so noting the body part it never consulted would cap a fully computed answer
  // below Verified for a gap that had no bearing on it.
  if (bodyParts.length === 0) {
    unresolved.push(
      `${DOCUMENT_PART} is not in the package, and neither is any header, footer or notes part; there is no paragraph content to resolve`
    );
  }

  const document = bodyParts.find(part => part.path === DOCUMENT_PART)?.document ?? null;

  return { styles, numbering, document, bodyParts, unresolved };
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
  // Where the numbering came from decides where it sits in the cascade, and the two
  // placements give opposite answers - see CascadeContext.numberingSource.
  let numberingSource: 'paragraph' | 'style' = 'paragraph';
  if (numberingRef.numId === null && paragraphStyleId) {
    const style = context.styles.styles.get(paragraphStyleId);
    numberingRef = readNumberingReference(style?.pPr ?? null);
    numberingSource = 'style';
  }

  // A w:numPr carrying numId="0" says this paragraph is NOT a numbered item, and Word
  // additionally discards the indentation inherited from the style hierarchy when it
  // sees one. Without this a cancelled list item keeps its list indent and sits out of
  // line with the body text around it.
  const cancelsNumbering = numberingRef.numId === '0';

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
    numbering,
    numberingSource
  });

  if (cancelsNumbering) {
    paragraphResult.properties.delete('ind');
    paragraphResult.trace.push({
      layer: 'numbering:cancelled',
      contributed: ['ind'],
      note: 'numId="0" cancels numbering, and Word also discards the indentation inherited from the style hierarchy when it sees one.'
    });
  }

  const runResult = run
    ? resolveRunProperties(context.styles, {
        paragraphStyleId,
        characterStyleId,
        directRPr: directRPr ?? undefined,
        insideTable: position !== null,
        tableStyle: tableStyleLayers,
        numbering,
        numberingSource
      })
    : null;

  const explanation: string[] = [];
  // Which story the paragraph belongs to is part of the answer, not decoration: the
  // same style renders differently in a header (different section, different defaults
  // in the reader's head), and a reader told only "sz = 32" cannot tell whether the
  // analysis looked at the paragraph they had selected.
  const ownerPart = context.bodyParts.find(part => part.document === paragraph.ownerDocument);
  if (ownerPart) explanation.push(`Paragraph is in ${ownerPart.path}.`);
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
 * Convenience entry point: analyze the nth paragraph of `word/document.xml`.
 *
 * **Deliberately the main story only**, even though the context now holds headers,
 * footers and notes. An index is only meaningful over a sequence a caller can see, and
 * there is no such thing as "the 4th paragraph of the document" once five stories are
 * concatenated in an order the file itself does not define — the number would silently
 * mean something different as soon as a header were added. Callers that need a
 * paragraph in another part locate it by markup (`locateParagraphByMarkup`) or walk
 * `context.bodyParts` themselves and call `analyzeParagraphFormatting` directly, which
 * resolves any story's paragraph exactly the same way.
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

/**
 * Normalizes markup for comparison.
 *
 * Two sources of noise have to go. The editor pretty-prints what it displays, so a
 * caller's snippet is rarely byte-identical to the same element in the parsed
 * document. And serializing a subtree re-emits the namespace declarations it inherited
 * from the root, so a serialized `<w:p>` carries an `xmlns:w` that the caller's snippet
 * does not. Neither difference is meaningful here.
 */
const normalizeMarkup = (xml: string): string =>
  xml
    .replace(/\s+xmlns(:[A-Za-z0-9_-]+)?="[^"]*"/g, '')
    .replace(/>\s+</g, '><')
    .replace(/\s+/g, ' ')
    .trim();

export interface LocatedParagraph {
  paragraph: Element;
  run?: Element;
  /**
   * The body part the paragraph was found in. Null when a bare `Document` was searched,
   * because then there is no package to name the part from — reported as null rather
   * than defaulted to `word/document.xml`, which would be a guess presented as fact.
   */
  part: string | null;
}

/**
 * Finds the paragraph a snippet of markup belongs to, and the run inside it if the
 * snippet is one.
 *
 * Searches every loaded body part, not just the body: a header paragraph is as
 * selectable as any other, and searching only `document.xml` means the caller sees
 * "Unverified" for markup that was fully resolvable.
 *
 * Returns null unless the match is **unambiguous**. A document routinely contains many
 * identical paragraphs — empty ones especially — and resolving the formatting of the
 * wrong one would produce a confidently wrong answer under a "Verified" badge. Refusing
 * to guess costs a fallback to the ordinary explanation; guessing costs correctness.
 * Ambiguity is counted across parts, not within one: the same empty paragraph in
 * `document.xml` and in `header1.xml` is exactly as unresolvable as two of them in the
 * body, and widening the search would otherwise quietly turn refusals into coin flips.
 *
 * Accepts a whole context (all stories) or a single `Document` for callers that already
 * hold one.
 */
export const locateParagraphByMarkup = (
  source: Document | WordDocumentContext,
  rawXml: string
): LocatedParagraph | null => {
  const needle = normalizeMarkup(rawXml);
  if (needle === '') return null;

  const stories: { path: string | null; document: Document }[] =
    'bodyParts' in source ? source.bodyParts : [{ path: null, document: source }];

  const serialize = (el: Element) => normalizeMarkup(new XMLSerializer().serializeToString(el));
  const candidates = stories.flatMap(story =>
    Array.from(story.document.getElementsByTagNameNS(W_NAMESPACE, 'p')).map(paragraph => ({
      part: story.path,
      paragraph,
      markup: serialize(paragraph)
    }))
  );

  // The snippet may itself be a paragraph.
  const exact = candidates.filter(c => c.markup === needle);
  if (exact.length === 1) return { paragraph: exact[0].paragraph, part: exact[0].part };
  if (exact.length > 1) return null;

  const containing = candidates.filter(c => c.markup.includes(needle));
  if (containing.length !== 1) return null;

  const { paragraph, part } = containing[0];
  const runs = Array.from(paragraph.getElementsByTagNameNS(W_NAMESPACE, 'r'));
  const matchingRuns = runs.filter(r => serialize(r) === needle || serialize(r).includes(needle));

  return {
    paragraph,
    part,
    // Only attach a run when it is unambiguous; otherwise resolve the paragraph alone.
    run: matchingRuns.length === 1 ? matchingRuns[0] : undefined
  };
};

/**
 * One-call entry point for a UI: given the package parts and a snippet of markup,
 * produce evidence ready to hand to the AI layer.
 *
 * The package may hold any mix of body parts — a header on its own is enough, since
 * headers resolve against the same stylesheet as the body. What matters is that at
 * least one story parsed; `document.xml` in particular is not required.
 *
 * Returns null when nothing could be computed — a non-Word package, an unlocatable
 * snippet — so the caller falls back to its ordinary path rather than showing an empty
 * "Verified" answer.
 */
export const computeEvidenceForMarkup = (
  parts: PackageParts,
  rawXml: string
): { lines: string[]; unresolved: string[]; part: string | null } | null => {
  const context = loadWordContext(parts);
  if (context.bodyParts.length === 0) return null;

  const located = locateParagraphByMarkup(context, rawXml);
  if (!located) return null;

  const analysis = analyzeParagraphFormatting(context, located.paragraph, located.run);
  return { lines: analysis.explanation, unresolved: analysis.unresolved, part: located.part };
};
