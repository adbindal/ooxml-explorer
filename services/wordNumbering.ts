/**
 * WordprocessingML numbering resolution — cascade layer 3.
 *
 * Numbering is the most-indirected structure in the format and the most common source
 * of "my list renders wrong" bugs. The visible "3.2.1" exists nowhere in the file; it
 * is computed. What *is* in the file is a chain of references that a resolver has to
 * walk correctly, and three separate patterns for walking it.
 *
 *   Pattern 1 — direct
 *     w:p/w:pPr/w:numPr/{w:ilvl,w:numId} → w:num[@numId] → @abstractNumId
 *       → w:abstractNum[@abstractNumId] → w:lvl[@ilvl]
 *
 *   Pattern 2 — style-linked
 *     The paragraph carries only w:pStyle. The w:lvl carries w:pStyle pointing back at
 *     it. This is how Heading 1/2/3 auto-numbering works.
 *
 *   Pattern 3 — named list styles, and the one that silently returns nothing
 *     The w:abstractNum reached in pattern 1 has NO w:lvl children at all — only a
 *     w:numStyleLink naming a style, whose pPr/numPr/numId points at a *different*
 *     w:num, reaching a *different* w:abstractNum that carries w:styleLink and the
 *     actual levels. A resolver that expects levels on the first abstractNum finds
 *     none and renders the list unnumbered.
 *
 * Two rules that look like edge cases and are not:
 *
 *   - `w:numId w:val="0"` is not a lookup. The specification is explicit that 0 "shall
 *     never be used to point to a numbering definition instance" and instead means the
 *     *removal* of numbering. A resolver that looks up num 0, finds nothing, and falls
 *     back to the style's numbering renders bullets on a paragraph that asked to have
 *     none.
 *
 *   - numId shares a *counter*; abstractNumId shares a *definition*. Two w:num entries
 *     pointing at the same abstractNum are two independent counters using one format.
 *     Confusing the two produces either "the list restarts in the middle" or "the
 *     second list continues from the first" — the classic generated-document bug.
 */

import { W_NAMESPACE, type StyleSheet } from './wordStyleResolver';

/** Word will not open a file whose ilvl falls outside this range. */
export const MIN_ILVL = 0;
export const MAX_ILVL = 8;

export interface NumberingDefinitions {
  /** `abstractNumId` → `w:abstractNum` element. */
  abstractNums: Map<string, Element>;
  /** `numId` → `w:num` element. */
  nums: Map<string, Element>;
}

export interface ResolvedNumbering {
  numId: string;
  ilvl: number;
  /** The abstract definition finally reached, after any numStyleLink hop. */
  abstractNumId: string | null;
  /** The winning `w:lvl` element, or null when nothing resolved. */
  lvl: Element | null;
  /** `w:start`, after any `w:startOverride`. */
  start: string | null;
  numFmt: string | null;
  lvlText: string | null;
  /** Separator between number and text: `tab` (default), `space` or `nothing`. */
  suff: string;
  /** `w:isLgl` — render every level as decimal regardless of each level's numFmt. */
  isLgl: boolean;
  /** `w:lvl/w:pPr`, which overrides the paragraph style's indentation. */
  pPr: Element | null;
  rPr: Element | null;
  /** Ordered account of how the level was reached. */
  trace: string[];
  /** Set when the markup is legal per schema but Word will reject or misbehave. */
  problems: string[];
}

const childrenOf = (parent: Element, localName: string): Element[] =>
  Array.from(parent.children).filter(
    el => el.namespaceURI === W_NAMESPACE && el.localName === localName
  );

const firstChild = (parent: Element, localName: string): Element | null =>
  childrenOf(parent, localName)[0] ?? null;

const valOf = (element: Element | null): string | null =>
  element?.getAttributeNS(W_NAMESPACE, 'val') ?? null;

/** Parses `numbering.xml` into lookup maps. */
export const parseNumbering = (numberingXml: string): NumberingDefinitions => {
  const doc = new DOMParser().parseFromString(numberingXml, 'application/xml');
  const abstractNums = new Map<string, Element>();
  const nums = new Map<string, Element>();

  for (const el of Array.from(doc.getElementsByTagNameNS(W_NAMESPACE, 'abstractNum'))) {
    const id = el.getAttributeNS(W_NAMESPACE, 'abstractNumId');
    if (id !== null) abstractNums.set(id, el);
  }
  for (const el of Array.from(doc.getElementsByTagNameNS(W_NAMESPACE, 'num'))) {
    const id = el.getAttributeNS(W_NAMESPACE, 'numId');
    if (id !== null) nums.set(id, el);
  }

  return { abstractNums, nums };
};

/** Finds `w:lvl[@w:ilvl="n"]` among an element's children. */
const levelWithIlvl = (parent: Element, ilvl: number): Element | null =>
  childrenOf(parent, 'lvl').find(
    el => el.getAttributeNS(W_NAMESPACE, 'ilvl') === String(ilvl)
  ) ?? null;

/**
 * Follows `w:numStyleLink` from an abstract definition that carries no levels.
 *
 * The hop leaves the numbering part entirely: numStyleLink names a style, the style's
 * `pPr/numPr/numId` names a `w:num`, and that num reaches the abstract definition that
 * actually holds the levels. Returns the abstractNumId finally reached.
 */
const followNumStyleLink = (
  abstractNum: Element,
  numbering: NumberingDefinitions,
  styles: StyleSheet,
  trace: string[],
  seen: Set<string>
): string | null => {
  const linkedStyleId = valOf(firstChild(abstractNum, 'numStyleLink'));
  if (!linkedStyleId) return null;

  trace.push(`numStyleLink → style:${linkedStyleId}`);
  const style = styles.styles.get(linkedStyleId);
  if (!style?.pPr) {
    trace.push(`style:${linkedStyleId} not found or has no pPr`);
    return null;
  }

  const numPr = firstChild(style.pPr, 'numPr');
  const linkedNumId = numPr ? valOf(firstChild(numPr, 'numId')) : null;
  if (!linkedNumId) {
    trace.push(`style:${linkedStyleId} has no numPr/numId`);
    return null;
  }

  trace.push(`style:${linkedStyleId} → numId ${linkedNumId}`);
  return resolveAbstractNumId(linkedNumId, numbering, styles, trace, seen);
};

/**
 * Maps a numId to the abstractNumId that holds its levels, following numStyleLink.
 *
 * Cycles are tolerated rather than fatal — a malformed numbering part should degrade,
 * not hang the tab.
 */
const resolveAbstractNumId = (
  numId: string,
  numbering: NumberingDefinitions,
  styles: StyleSheet,
  trace: string[],
  seen: Set<string>
): string | null => {
  if (seen.has(numId)) {
    trace.push(`cycle detected at numId ${numId}`);
    return null;
  }
  seen.add(numId);

  const num = numbering.nums.get(numId);
  if (!num) {
    trace.push(`numId ${numId} not found in numbering part`);
    return null;
  }

  const abstractNumId = valOf(firstChild(num, 'abstractNumId'));
  if (!abstractNumId) {
    trace.push(`numId ${numId} has no abstractNumId`);
    return null;
  }

  const abstractNum = numbering.abstractNums.get(abstractNumId);
  if (!abstractNum) {
    trace.push(`abstractNumId ${abstractNumId} not found`);
    return null;
  }

  // Pattern 3: no levels here, only a link onward.
  if (childrenOf(abstractNum, 'lvl').length === 0) {
    const linked = followNumStyleLink(abstractNum, numbering, styles, trace, seen);
    if (linked) return linked;
  }

  return abstractNumId;
};

/**
 * Resolves the effective numbering level for a paragraph.
 *
 * Returns null when the paragraph is genuinely not numbered — including the `numId=0`
 * case, which is an explicit instruction to remove numbering rather than a failed
 * lookup. Callers must not treat null as "fall back to the style's numbering".
 */
export const resolveNumbering = (
  numbering: NumberingDefinitions,
  styles: StyleSheet,
  input: { numId: string | null; ilvl?: number | null }
): ResolvedNumbering | null => {
  const trace: string[] = [];
  const problems: string[] = [];
  const numId = input.numId;

  if (numId === null || numId === '') return null;

  // Explicit removal, not a lookup. See the module comment.
  if (numId === '0') {
    return null;
  }

  const ilvl = input.ilvl ?? 0;
  if (ilvl < MIN_ILVL || ilvl > MAX_ILVL) {
    // Schema-legal (w:ilvl is an integer) but Word refuses to open the file.
    problems.push(`ilvl ${ilvl} is outside the permitted range ${MIN_ILVL}-${MAX_ILVL}; Word will not open this file`);
  }

  trace.push(`numId ${numId}`);
  const abstractNumId = resolveAbstractNumId(numId, numbering, styles, trace, new Set());

  const result: ResolvedNumbering = {
    numId,
    ilvl,
    abstractNumId,
    lvl: null,
    start: null,
    numFmt: null,
    lvlText: null,
    suff: 'tab',
    isLgl: false,
    pPr: null,
    rPr: null,
    trace,
    problems
  };

  if (!abstractNumId) return result;

  // A w:lvlOverride on the w:num takes precedence over the abstract definition, and
  // comes in two forms: a complete replacement w:lvl, or only a w:startOverride that
  // restarts the counter while keeping the abstract level's formatting.
  const num = numbering.nums.get(numId);
  let startOverride: string | null = null;
  let lvl: Element | null = null;

  if (num) {
    const override = childrenOf(num, 'lvlOverride').find(
      el => el.getAttributeNS(W_NAMESPACE, 'ilvl') === String(ilvl)
    );
    if (override) {
      startOverride = valOf(firstChild(override, 'startOverride'));
      const overrideLvl = firstChild(override, 'lvl');
      if (overrideLvl) {
        lvl = overrideLvl;
        trace.push(`lvlOverride supplies a full level for ilvl ${ilvl}`);
      } else if (startOverride !== null) {
        trace.push(`lvlOverride restarts ilvl ${ilvl} at ${startOverride}`);
      }
    }
  }

  if (!lvl) {
    const abstractNum = numbering.abstractNums.get(abstractNumId);
    lvl = abstractNum ? levelWithIlvl(abstractNum, ilvl) : null;
    if (lvl) trace.push(`abstractNum ${abstractNumId} supplies ilvl ${ilvl}`);
  }

  if (!lvl) {
    trace.push(`no level definition found for ilvl ${ilvl}`);
    if (startOverride !== null) result.start = startOverride;
    return result;
  }

  result.lvl = lvl;
  result.start = startOverride ?? valOf(firstChild(lvl, 'start'));
  result.numFmt = valOf(firstChild(lvl, 'numFmt'));
  result.lvlText = valOf(firstChild(lvl, 'lvlText'));
  result.suff = valOf(firstChild(lvl, 'suff')) ?? 'tab';
  result.pPr = firstChild(lvl, 'pPr');
  result.rPr = firstChild(lvl, 'rPr');

  const isLgl = firstChild(lvl, 'isLgl');
  result.isLgl = isLgl !== null && !['0', 'false', 'off'].includes((valOf(isLgl) ?? '1').toLowerCase());

  // w:lvlText uses 1-based tokens (%1, %2) while w:ilvl is 0-based — a persistent
  // off-by-one. Bullet levels do not substitute tokens even when they contain them.
  if (result.numFmt === 'bullet' && result.lvlText?.includes('%')) {
    problems.push('lvlText contains a %N token but numFmt is "bullet"; tokens are not substituted for bullets');
  }

  return result;
};

/**
 * Extracts `{ numId, ilvl }` from a paragraph's `w:pPr`.
 *
 * Direct numbering on the paragraph overrides numbering inherited from its style, so a
 * caller that finds a `numPr` here must not also apply the style's.
 */
export const readNumberingReference = (
  pPr: Element | null | undefined
): { numId: string | null; ilvl: number | null } => {
  if (!pPr) return { numId: null, ilvl: null };
  const numPr = firstChild(pPr, 'numPr');
  if (!numPr) return { numId: null, ilvl: null };
  const rawIlvl = valOf(firstChild(numPr, 'ilvl'));
  return {
    numId: valOf(firstChild(numPr, 'numId')),
    ilvl: rawIlvl === null ? null : Number.parseInt(rawIlvl, 10)
  };
};
