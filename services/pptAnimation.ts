/**
 * Slide animation and timing — the build that silently stops building.
 *
 * An animation does not contain the thing it animates. It NAMES it, by number:
 * `p:spTgt/@spid` holds a shape id, and the shape it refers to is whichever
 * `p:cNvPr/@id` in the slide's shape tree happens to match. Nothing binds the two
 * together. Which produces the failure this module exists for:
 *
 *   DELETE THE SHAPE AND THE ANIMATION DOES NOT ERROR, WARN, OR DISAPPEAR.
 *   IT SIMPLY NEVER FIRES.
 *
 * PowerPoint opens the deck without complaint. The slide looks right in the editor,
 * right in the thumbnail, right in a PDF export — a static render never runs the timing
 * tree at all, so a screenshot diff is clean by construction. The three bullets that were
 * supposed to appear one click at a time appear all at once, or not at all, and the
 * person who finds out is the one presenting.
 *
 * The same shape-id-by-number arrangement appears three more times on one slide, and each
 * one fails the same silent way:
 *
 *   • `p:bldLst` build entries (`p:bldP`, `p:bldDgm`, `p:bldOleChart`, `p:bldGraphic`)
 *     each carry a REQUIRED `@spid` naming the shape whose text or diagram is built.
 *   • `p:cond/@evt` conditions — "start this sequence when the user clicks THAT shape" —
 *     name the triggering shape through their own `p:tgtEl/p:spTgt`.
 *   • `p:txEl/p:pRg` narrows a target to a range of paragraphs INSIDE the shape, so an
 *     animation can be aimed at a paragraph the shape no longer contains even when the
 *     shape itself is still there.
 *
 * And if two shapes on one slide share a `p:cNvPr/@id`, every one of those references
 * becomes ambiguous at once: the animation still runs, on whichever shape the consumer
 * picked.
 *
 * ⚠️ TWO CORRECTIONS TO THE COMMON DESCRIPTION OF THIS MARKUP — BOTH SCHEMA-VERIFIED.
 *
 * 1. A BEHAVIOUR DOES NOT TARGET THROUGH `p:cTn`. The path usually written as
 *    `p:cTn/p:tgtEl/p:spTgt` does not exist: `CT_TLCommonTimeNodeData` (`p:cTn`) has
 *    exactly six children — `p:stCondLst`, `p:endCondLst`, `p:endSync`, `p:iterate`,
 *    `p:childTnLst`, `p:subTnLst` — and `p:tgtEl` is not among them. The real path is
 *    `p:anim/p:cBhvr/p:tgtEl/p:spTgt`, where `p:cBhvr` (`CT_TLCommonBehaviorData`)
 *    requires both a `p:cTn` and a `p:tgtEl` as siblings. Media nodes are the analogous
 *    `p:video/p:cMediaNode/p:tgtEl`. Reading `p:cTn/p:tgtEl` finds nothing, on every deck.
 *
 * 2. THERE IS NO `p:condLst` ELEMENT. Conditions live in four differently named lists,
 *    all of type `CT_TLTimeConditionList`: `p:stCondLst` and `p:endCondLst` on `p:cTn`,
 *    and `p:prevCondLst` / `p:nextCondLst` on `p:seq`. This module finds `p:cond` by name
 *    at any depth rather than by its parent, so all four are covered and no invented
 *    parent name is relied on.
 *
 * PROVENANCE — verified against the Open XML SDK's machine-readable schema data
 * (`data/schemas/schemas_openxmlformats_org_presentationml_2006_main.json` and the
 * DrawingML equivalent):
 *   • `p:timing` (`CT_SlideTiming`) children: `p:tnLst`, `p:bldLst`, `p:extLst`. It is
 *     permitted on `p:sld`, `p:sldLayout` and `p:sldMaster` — and NOT on `p:notes`.
 *   • `p:tnLst` (`CT_RootTimeNode`) holds a single `p:par`; `p:childTnLst` / `p:subTnLst`
 *     (`CT_TimeNodeList`) hold `p:par`, `p:seq`, `p:excl`, `p:anim`, `p:animClr`,
 *     `p:animEffect`, `p:animMotion`, `p:animRot`, `p:animScale`, `p:cmd`, `p:set`,
 *     `p:audio`, `p:video`.
 *   • Every one of `p:anim`, `p:animClr`, `p:animEffect`, `p:animMotion`, `p:animRot`,
 *     `p:animScale`, `p:cmd`, `p:set` has `p:cBhvr` as its first child.
 *   • `p:tgtEl` (`CT_TLTimeTargetElement`) is a choice of `p:sldTgt`, `p:sndTgt`,
 *     `p:spTgt`, `p:inkTgt`, `p14:bmkTgt`.
 *   • `p:spTgt` (`CT_TLShapeTargetElement`) has `@spid` REQUIRED, and a choice of
 *     `p:bg`, `p:subSp`, `p:oleChartEl`, `p:txEl`, `p:graphicEl`.
 *   • `p:txEl` (`CT_TLTextTargetElement`) is a choice of `p:charRg` and `p:pRg`, both
 *     `CT_IndexRange`, whose `@st` and `@end` are REQUIRED `xsd:unsignedInt`.
 *   • `p:cond` and `p:endSync` are both `CT_TLTimeCondition`: attributes `@evt`
 *     (optional, `ST_TLTriggerEvent`) and `@delay`, children `p:tgtEl`, `p:tn`, `p:rtn`.
 *   • `p:bldLst` (`CT_BuildList`) children `p:bldP`, `p:bldDgm`, `p:bldOleChart`,
 *     `p:bldGraphic`; all four carry REQUIRED `@spid` and `@grpId`.
 *   • `p:cNvPr` is a PresentationML element of DrawingML type `CT_NonVisualDrawingProps`;
 *     `@id` is a REQUIRED `UInt32`, `@name` REQUIRED.
 *   • `p:sp` may carry at most one `p:txBody` (`a:CT_TextBody`), whose paragraphs are its
 *     direct `a:p` children.
 *
 * ⚠️ `@spid` IS NOT ALWAYS A SHAPE ID — the same trap `pptMedia.ts` documents, and it
 * applies identically to `p:spTgt`, `p:subSp`, `p:inkTgt` and all four `p:bld*` elements.
 * The SDK records two validators for each: `StringValidator` for Office2007, numeric
 * `a:ST_DrawingElementId` from Office2010 on. `p:cNvPr/@id` is `UInt32` in every version.
 * A 2007-era deck can therefore legally write `spid="_x0000_s1026"` — a VML shape name
 * that matches no `p:cNvPr/@id` BY DESIGN. Reporting those as dangling would be a
 * confident wrong answer on every deck of that vintage, so a non-numeric `@spid` yields
 * `exists: null` and no finding.
 *
 * ✅ `p:pRg` INDICES ARE ZERO-BASED AND INCLUSIVE. The schema data records only the types,
 * so this was taken from the normative example reproduced in the SDK's own reference
 * documentation for `pRg` (ISO/IEC 29500-1, 1st Edition): "Consider an animation entrance
 * of the first 3 text paragraphs" is written `<p:pRg st="0" end="2"/>`. That is what makes
 * an out-of-range check possible at all — under a one-based reading `end="2"` on a
 * two-paragraph shape would be legal, and this module would over-report.
 *
 * WHAT THIS MODULE DELIBERATELY DOES NOT DO.
 *
 * `p:video` and `p:audio` timing nodes are `pptMedia.ts`'s, including their dangling
 * `@spid`s (`media/dangling-trigger`). They are skipped here rather than re-reported, so
 * the panel does not say the same thing twice in two different vocabularies.
 *
 * NAMESPACES ARE COMPARED BY EXACT EQUALITY. `conformance.ts` rewrites ISO Strict URIs to
 * their Transitional equivalents before any analyzer runs, so Strict-tolerant matching
 * here would be dead code that looks load-bearing.
 */

import { finding, renderFindings, type Finding, type Severity } from './findings';

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/**
 * Severity and silence per kind, decided once here.
 *
 * EVERY KIND IS SILENT, and that is the finding rather than an oversight. A timing tree
 * is not drawn: a slide with a dead animation renders exactly like the same slide with a
 * working one, in the editor, in the thumbnail, and in every static export. There is no
 * visible counterpart to a missing poster frame here, because nothing about an animation
 * is visible until someone advances the slide.
 *
 * `duplicate-shape-id` is a `warning` rather than an `error` because the animation still
 * fires — on one of the two shapes, chosen by the consumer rather than by the author. A
 * wrong shape moving is a different defect from no shape moving, and calling it the same
 * would overstate it. `inverted-range` is a `warning` for the honest reason that ECMA does
 * not say what a consumer does with a range whose start is past its end: the effect is
 * undefined, not provably absent.
 */
const ANIMATION_RULES = {
  'dead-target':        { severity: 'error',   silent: true },
  'dead-trigger':       { severity: 'error',   silent: true },
  'dead-build':         { severity: 'error',   silent: true },
  'paragraph-out-of-range': { severity: 'error', silent: true },
  'duplicate-shape-id': { severity: 'warning', silent: true },
  'inverted-range':     { severity: 'warning', silent: true }
} as const satisfies Record<string, { severity: Severity; silent: boolean }>;

export type AnimationProblemKind = keyof typeof ANIMATION_RULES;

const animationFinding = (
  kind: AnimationProblemKind,
  part: string,
  message: string,
  remediation: string,
  subject?: Record<string, string>
): Finding => finding(`animation/${kind}`, part, message, remediation, { ...ANIMATION_RULES[kind], subject });

/**
 * The behaviour nodes, keyed by local name.
 *
 * All eight share the `p:cBhvr` wrapper and therefore the same target path. `p:cmd` is
 * included because a command behaviour is aimed at a shape exactly like the others and
 * fails exactly like the others when that shape is gone.
 */
const BEHAVIOUR_NODES: Readonly<Record<string, string>> = {
  anim: 'a property animation',
  animClr: 'a colour animation',
  animEffect: 'an entrance or exit effect',
  animMotion: 'a motion path',
  animRot: 'a rotation',
  animScale: 'a scale',
  set: 'a property change',
  cmd: 'a command'
};

/** The `p:bldLst` entries, keyed by local name. All four carry a required `@spid`. */
const BUILD_NODES: Readonly<Record<string, string>> = {
  bldP: 'a paragraph build',
  bldDgm: 'a diagram build',
  bldOleChart: 'an OLE chart build',
  bldGraphic: 'a graphic build'
};

/** What a `p:tgtEl` selected. Only `shape` is ever judged. */
export type AnimationTargetKind = 'shape' | 'slide' | 'sound' | 'ink' | 'other' | 'none';

export interface AnimationTarget {
  kind: AnimationTargetKind;
  /** `p:spTgt/@spid` verbatim. */
  shapeId: string | null;
  /**
   * Whether a shape with that id is in this part's shape tree.
   *
   * `null` means unknowable — no shape target at all, or the Office 2007 string form of
   * `@spid` that cannot be matched against a `p:cNvPr/@id`.
   */
  exists: boolean | null;
  /** `p:txEl/p:pRg`, zero-based and inclusive. */
  paragraphRange: { start: number; end: number } | null;
  /** True when the target narrows to characters (`p:charRg`) rather than paragraphs. */
  characterRange: boolean;
  /** `p:spTgt/p:subSp/@spid` — a sub-shape id inside a group or diagram, never judged. */
  subShapeId: string | null;
}

/** One behaviour node in the timing tree. */
export interface Animation {
  part: string;
  element: Element;
  /** `p:anim`, `p:animEffect`, … */
  label: string;
  /** `p:cBhvr/p:cTn/@id`, for naming the node in a message. */
  timeNodeId: string | null;
  target: AnimationTarget;
  problems: Finding[];
}

/** One `p:cond` (or `p:endSync`) that starts or stops timing when something happens. */
export interface AnimationTrigger {
  part: string;
  element: Element;
  /** `p:cond` or `p:endSync`. */
  label: string;
  /** `@evt` verbatim, e.g. `onClick`. Optional in the schema. */
  event: string | null;
  /** The condition list it sits in — `p:stCondLst`, `p:nextCondLst`, … */
  list: string | null;
  target: AnimationTarget;
  problems: Finding[];
}

/** One `p:bldLst` entry: which shape gets built, and how. */
export interface AnimationBuild {
  part: string;
  element: Element;
  /** `p:bldP`, `p:bldDgm`, `p:bldOleChart`, `p:bldGraphic`. */
  label: string;
  shapeId: string | null;
  exists: boolean | null;
  problems: Finding[];
}

export interface SlideShape {
  id: string;
  name: string | null;
  /** The `p:sp`, `p:pic`, `p:grpSp`, … that owns the `p:cNvPr`. */
  element: Element;
  /**
   * Paragraphs in the shape's text body.
   *
   * `null` — not `0` — when the shape kind carries no direct `p:txBody`: a graphic frame
   * holding a table or chart has text, but not at a paragraph index this can count, so
   * "how many paragraphs does it have" has no answer here rather than the answer zero.
   */
  paragraphCount: number | null;
}

export interface AnimationIndex {
  part: string;
  animations: Animation[];
  triggers: AnimationTrigger[];
  builds: AnimationBuild[];
  /** Every shape in the tree that carries an id, first occurrence wins. */
  shapes: Map<string, SlideShape>;
  /** Ids carried by more than one shape, and how many shapes carry each. */
  duplicateShapeIds: Map<string, number>;
  /** Every finding for this part: each item's own, plus the slide-level duplicate ids. */
  problems: Finding[];
}

const rootOf = (node: Document | Element): ParentNode =>
  'documentElement' in node && node.documentElement ? node.documentElement : (node as Element);

const childrenNamed = (parent: Element, ns: string, local: string): Element[] =>
  Array.from(parent.children).filter(c => c.namespaceURI === ns && c.localName === local);

const childNamed = (parent: Element, ns: string, local: string): Element | null =>
  childrenNamed(parent, ns, local)[0] ?? null;

const descendantsNamed = (root: ParentNode, ns: string, local: string): Element[] =>
  Array.from(root.querySelectorAll('*')).filter(el => el.namespaceURI === ns && el.localName === local);

/** Shape kinds a `p:cNvPr` can hang off, in the order the schema lists them under `p:spTree`. */
const SHAPE_KINDS = new Set(['sp', 'grpSp', 'graphicFrame', 'cxnSp', 'pic', 'contentPart', 'spTree']);

/** The `p:sp`/`p:pic`/… that owns a `p:cNvPr`, found by walking out of its `p:nv*Pr` wrapper. */
const owningShape = (cNvPr: Element): Element | null => {
  let node: Element | null = cNvPr.parentElement;
  while (node !== null) {
    if (node.namespaceURI === P_NS && SHAPE_KINDS.has(node.localName)) return node;
    node = node.parentElement;
  }
  return null;
};

/**
 * Paragraph count for a shape, or `null` when the question does not apply.
 *
 * Only `p:sp` carries a direct `p:txBody`, and `a:CT_TextBody`'s paragraphs are its direct
 * `a:p` children. A `p:graphicFrame` (table, chart, diagram) also contains `a:p` elements,
 * deep inside `a:tbl` — counting those would produce a paragraph total that `p:pRg` does
 * not index, so those shapes report `null` and are never judged.
 */
const paragraphCountOf = (shape: Element): number | null => {
  if (shape.localName !== 'sp') return null;
  const txBody = childNamed(shape, P_NS, 'txBody');
  return txBody === null ? 0 : childrenNamed(txBody, A_NS, 'p').length;
};

interface ShapeIndex {
  byId: Map<string, SlideShape>;
  duplicates: Map<string, number>;
}

/**
 * Every id-bearing shape in the part.
 *
 * Scoped to `p:cSld/p:spTree` when there is one. A `p:cNvPr` outside the shape tree is not
 * a shape an animation can target, and counting it would both invent targets and invent
 * duplicate ids.
 */
const readShapes = (root: ParentNode): ShapeIndex => {
  const spTree = descendantsNamed(root, P_NS, 'spTree')[0];
  const scope: ParentNode = spTree ?? root;

  const byId = new Map<string, SlideShape>();
  const counts = new Map<string, number>();

  for (const cNvPr of descendantsNamed(scope, P_NS, 'cNvPr')) {
    const id = cNvPr.getAttribute('id');
    if (id === null) continue;
    const shape = owningShape(cNvPr);
    if (shape === null) continue;

    counts.set(id, (counts.get(id) ?? 0) + 1);
    if (byId.has(id)) continue;
    byId.set(id, {
      id,
      name: cNvPr.getAttribute('name'),
      element: shape,
      paragraphCount: paragraphCountOf(shape)
    });
  }

  const duplicates = new Map([...counts].filter(([, count]) => count > 1));
  return { byId, duplicates };
};

const NO_TARGET: AnimationTarget = {
  kind: 'none',
  shapeId: null,
  exists: null,
  paragraphRange: null,
  characterRange: false,
  subShapeId: null
};

const readIndexRange = (range: Element): { start: number; end: number } | null => {
  const start = Number(range.getAttribute('st'));
  const end = Number(range.getAttribute('end'));
  // Both are required unsignedInt. A missing or non-numeric one is a schema-validity
  // problem, not an animation problem, so it is left unjudged rather than guessed at.
  if (!Number.isInteger(start) || !Number.isInteger(end)) return null;
  return { start, end };
};

/** Reads one `p:tgtEl`, resolving its shape target against the slide's shapes. */
const readTarget = (tgtEl: Element | null, shapes: ShapeIndex): AnimationTarget => {
  if (tgtEl === null) return NO_TARGET;

  if (childNamed(tgtEl, P_NS, 'sldTgt') !== null) return { ...NO_TARGET, kind: 'slide' };
  if (childNamed(tgtEl, P_NS, 'sndTgt') !== null) return { ...NO_TARGET, kind: 'sound' };
  if (childNamed(tgtEl, P_NS, 'inkTgt') !== null) return { ...NO_TARGET, kind: 'ink' };

  const spTgt = childNamed(tgtEl, P_NS, 'spTgt');
  if (spTgt === null) return { ...NO_TARGET, kind: tgtEl.children.length === 0 ? 'none' : 'other' };

  const shapeId = spTgt.getAttribute('spid');
  // Office 2007 wrote a VML shape name here; Office 2010 on writes ST_DrawingElementId.
  // Only the numeric form can be compared with p:cNvPr/@id at all.
  const numeric = shapeId !== null && /^\d+$/.test(shapeId);
  const exists = shapeId === null || !numeric ? null : shapes.byId.has(shapeId);

  const txEl = childNamed(spTgt, P_NS, 'txEl');
  const pRg = txEl === null ? null : childNamed(txEl, P_NS, 'pRg');
  const subSp = childNamed(spTgt, P_NS, 'subSp');

  return {
    kind: 'shape',
    shapeId,
    exists,
    paragraphRange: pRg === null ? null : readIndexRange(pRg),
    characterRange: txEl !== null && childNamed(txEl, P_NS, 'charRg') !== null,
    subShapeId: subSp?.getAttribute('spid') ?? null
  };
};

/**
 * The paragraph-range checks, shared by behaviours and triggers.
 *
 * Only run when the shape itself resolved: an out-of-range paragraph on a shape that does
 * not exist is the same defect reported twice, and the dead target is the one worth
 * reading. Ranges are zero-based and inclusive (see PROVENANCE).
 */
const paragraphProblems = (target: AnimationTarget, shapes: ShapeIndex, part: string, describe: string): Finding[] => {
  const range = target.paragraphRange;
  if (range === null || target.exists !== true || target.shapeId === null) return [];

  const problems: Finding[] = [];
  if (range.start > range.end) {
    problems.push(animationFinding(
      'inverted-range', part,
      `${describe} selects paragraphs ${range.start} to ${range.end} of shape ${target.shapeId} — a range that ends before it starts. ECMA does not define what a consumer does with an inverted range, so the effect is undefined rather than reliably absent; in practice nothing is animated.`,
      'Correct p:pRg so st is less than or equal to end.',
      { spid: target.shapeId, st: String(range.start), end: String(range.end) }
    ));
    return problems;
  }

  const shape = shapes.byId.get(target.shapeId);
  const count = shape?.paragraphCount ?? null;
  if (count === null) return problems;

  if (range.end >= count) {
    problems.push(animationFinding(
      'paragraph-out-of-range', part,
      `${describe} targets paragraphs ${range.start}–${range.end} of shape ${target.shapeId}${
        shape?.name ? ` ("${shape.name}")` : ''
      }, which has ${count === 0 ? 'no text at all' : `${count} paragraph(s), numbered 0 to ${count - 1}`}. The paragraphs it names are not there, so this part of the build has nothing to reveal — and the slide renders identically whether it runs or not.`,
      count === 0
        ? `Point the animation at a shape that has text, or delete it.`
        : `Narrow p:pRg to the paragraphs the shape has (0 to ${count - 1}), or restore the text that was removed.`,
      { spid: target.shapeId, st: String(range.start), end: String(range.end), paragraphs: String(count) }
    ));
  }
  return problems;
};

/**
 * Everything in one part's timing tree, and whether each reference still lands.
 *
 * `partPath` is used only to label findings; the markup is read from `doc`, which follows
 * `wordFields.ts` rather than `pptMedia.ts` because nothing here needs the package —
 * every id an animation names has to be in the same part.
 */
export function readAnimations(doc: Document | Element, partPath = ''): AnimationIndex {
  const root = rootOf(doc);
  const empty: AnimationIndex = {
    part: partPath,
    animations: [],
    triggers: [],
    builds: [],
    shapes: new Map(),
    duplicateShapeIds: new Map(),
    problems: []
  };

  const timings = descendantsNamed(root, P_NS, 'timing');
  if (timings.length === 0) return empty;

  const shapes = readShapes(root);
  const animations: Animation[] = [];
  const triggers: AnimationTrigger[] = [];
  const builds: AnimationBuild[] = [];

  for (const timing of timings) {
    for (const node of Array.from(timing.querySelectorAll('*'))) {
      if (node.namespaceURI !== P_NS) continue;

      // --- behaviour nodes ---------------------------------------------------
      const behaviour = BEHAVIOUR_NODES[node.localName];
      if (behaviour !== undefined) {
        // p:cBhvr is where BOTH the timing (p:cTn) and the target (p:tgtEl) live. There
        // is no p:tgtEl under p:cTn — see the correction at the top of this file.
        const cBhvr = childNamed(node, P_NS, 'cBhvr');
        const target = readTarget(cBhvr === null ? null : childNamed(cBhvr, P_NS, 'tgtEl'), shapes);
        const cTn = cBhvr === null ? null : childNamed(cBhvr, P_NS, 'cTn');
        const label = `p:${node.localName}`;
        const describe = `${label}, ${behaviour},`;

        const problems: Finding[] = [];
        if (target.exists === false) {
          problems.push(animationFinding(
            'dead-target', partPath,
            `${describe} targets shape id ${target.shapeId}, and no shape in ${partPath || 'this part'} has that id. The animation is dead: PowerPoint reports nothing, the slide looks exactly as intended in the editor and in any static export, and the effect simply never happens when the deck is presented.`,
            `Point p:spTgt/@spid at the p:cNvPr/@id of the shape this should animate, or delete the behaviour and its time node.`,
            { spid: target.shapeId ?? '', node: label }
          ));
        }
        problems.push(...paragraphProblems(target, shapes, partPath, describe));

        animations.push({
          part: partPath,
          element: node,
          label,
          timeNodeId: cTn?.getAttribute('id') ?? null,
          target,
          problems
        });
        continue;
      }

      // --- conditions --------------------------------------------------------
      // p:cond and p:endSync are the same type (CT_TLTimeCondition) and fail the same
      // way, so both are read. A condition that triggers on a time node (p:tn) or a
      // runtime node (p:rtn) rather than an element has no p:tgtEl and nothing to check.
      if (node.localName === 'cond' || node.localName === 'endSync') {
        const target = readTarget(childNamed(node, P_NS, 'tgtEl'), shapes);
        const event = node.getAttribute('evt');
        const parent = node.parentElement;
        const list = parent?.namespaceURI === P_NS ? `p:${parent.localName}` : null;
        const label = `p:${node.localName}`;
        const describe = `${label}${event === null ? '' : ` on ${event}`}`;

        const problems: Finding[] = [];
        if (target.exists === false) {
          problems.push(animationFinding(
            'dead-trigger', partPath,
            `A ${describe} condition${list === null ? '' : ` in ${list}`} waits on shape id ${target.shapeId}, and no shape in ${partPath || 'this part'} has that id. The event it waits for can never be raised, so whatever this condition was meant to start or stop never starts or stops — a click that does nothing, with no indication that anything is missing.`,
            `Point the condition's p:spTgt/@spid at the p:cNvPr/@id of the shape that should trigger it, or remove the condition.`,
            { spid: target.shapeId ?? '', ...(event === null ? {} : { evt: event }) }
          ));
        }
        problems.push(...paragraphProblems(target, shapes, partPath, `A ${describe} condition`));

        triggers.push({ part: partPath, element: node, label, event, list, target, problems });
        continue;
      }

      // --- build entries -----------------------------------------------------
      const build = BUILD_NODES[node.localName];
      if (build !== undefined) {
        const spid = node.getAttribute('spid');
        const numeric = spid !== null && /^\d+$/.test(spid);
        const exists = spid === null || !numeric ? null : shapes.byId.has(spid);
        const label = `p:${node.localName}`;

        const problems: Finding[] = [];
        if (exists === false) {
          problems.push(animationFinding(
            'dead-build', partPath,
            `${label}, ${build}, names shape id ${spid}, and no shape in ${partPath || 'this part'} has that id. The build entry describes how a shape that is not on the slide should be revealed, so it does nothing at all — and the slide is drawn identically with it and without it.`,
            `Point the build entry's @spid at a p:cNvPr/@id on this slide, or delete the entry from p:bldLst.`,
            { spid: spid ?? '', node: label }
          ));
        }

        builds.push({ part: partPath, element: node, label, shapeId: spid, exists, problems });
      }
    }
  }

  // --- slide-level: ambiguous ids -------------------------------------------
  // Reported only when something actually targets by id, because that is when the
  // ambiguity has a consequence. A duplicate id on a slide with no timing tree is a
  // different (and much quieter) problem than a duplicate id an animation depends on.
  const problems: Finding[] = [
    ...animations.flatMap(a => a.problems),
    ...triggers.flatMap(t => t.problems),
    ...builds.flatMap(b => b.problems)
  ];

  const targeted = new Set(
    [
      ...animations.map(a => a.target.shapeId),
      ...triggers.map(t => t.target.shapeId),
      ...builds.map(b => b.shapeId)
    ].filter((id): id is string => id !== null)
  );

  for (const [id, count] of shapes.duplicates) {
    if (!targeted.has(id)) continue;
    problems.push(animationFinding(
      'duplicate-shape-id', partPath,
      `${count} shapes on this slide share p:cNvPr/@id ${id}, and the timing tree targets that id. Every reference to it — animations, click triggers, build entries — is ambiguous: the effect still runs, on whichever of the shapes the consumer happens to resolve first, which need not be the one the author animated and need not be the same one in two different consumers.`,
      `Give each shape a unique p:cNvPr/@id and repoint the timing tree at the intended one.`,
      { spid: id, shapes: String(count) }
    ));
  }

  return {
    part: partPath,
    animations,
    triggers,
    builds,
    shapes: shapes.byId,
    duplicateShapeIds: shapes.duplicates,
    problems
  };
}

/**
 * The animations that will never run.
 *
 * The count worth leading with: "3 of 11 animations will never run" is a sentence someone
 * can act on, where eleven individual findings are a list someone skims.
 */
export function deadAnimations(index: AnimationIndex): Animation[] {
  return index.animations.filter(a => a.target.exists === false);
}

/** Every finding this module produces for one part. */
export function animationFindings(doc: Document | Element, partPath = ''): Finding[] {
  return readAnimations(doc, partPath).problems;
}

/**
 * Parts whose schema permits `p:timing`: slides, layouts and masters.
 *
 * Notes slides are deliberately absent — `CT_NotesSlide` has no `p:timing` child, so
 * matching them would add a part that can never contain anything to find.
 */
export const ANIMATION_HOST_PART = /^ppt\/(?:slides|slideLayouts|slideMasters)\/[^/]+\.xml$/;

const parseXml = (xml: string): Document | null => {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
};

/**
 * Evidence lines for the AI panel.
 *
 * Scans for the first host part that actually has a timing tree rather than the first one
 * that merely could, for the reason `pptMedia.ts` gives: a bundle routinely carries a
 * layout and a master alongside the slide, and taking the first match blind would report
 * "no animations" for a deck that has some, depending on how the bundle was assembled.
 */
export function computeAnimationEvidenceForMarkup(
  parts: Record<string, string>,
  rawXml: string
): { lines: string[]; unresolved: string[] } | null {
  let hostPath: string | null = null;
  let index: AnimationIndex | null = null;

  for (const path of Object.keys(parts).filter(p => ANIMATION_HOST_PART.test(p))) {
    const doc = parseXml(parts[path]);
    if (doc === null) continue;
    const found = readAnimations(doc, path);
    if (found.animations.length === 0 && found.triggers.length === 0 && found.builds.length === 0) continue;
    hostPath = path;
    index = found;
    break;
  }
  if (index === null || hostPath === null) return null;
  // ⚠️ INCOMPLETE. The agent writing this module was cut off here, so everything below
  // this point is missing: `idx` was bound for narrowing and never used. Do not register
  // this analyzer's `explain` entry until this function actually builds its lines.
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const idx: AnimationIndex = index;

  const lines: string[] = [];
  const unresolved: string[] = [];

  const dead = deadAnimations(index);
  lines.push(
    `${hostPath} has ${index.animations.length} animation behaviour(s) in its timing tree, ` +
      `${index.triggers.length} timing condition(s), and ${index.builds.length} build entry(ies) in p:bldLst.`
  );
  lines.push(
    dead.length === 0
      ? 'Every animation names a shape id that exists in this slide\'s shape tree.'
      : `${dead.length} of ${index.animations.length} animations will never run: they target shape ids (${[
          ...new Set(dead.map(a => a.target.shapeId ?? 'unstated'))
        ].join(', ')}) that no shape on this slide has. Nothing about the slide looks wrong — a timing tree is never drawn, so no rendering, thumbnail or export can reveal this.`
  );

  const deadTriggers = index.triggers.filter(t => t.target.exists === false).length;
  if (deadTriggers > 0) {
    lines.push(`${deadTriggers} timing condition(s) wait on a shape that is not on the slide, so the event they wait for cannot be raised.`);
  }
  const deadBuilds = index.builds.filter(b => b.exists === false).length;
  if (deadBuilds > 0) {
    lines.push(`${deadBuilds} build entry(ies) describe how to reveal a shape that is not on the slide.`);
  }

  // Describe the shape the user has open, when the selected markup names one.
  const selectedSpid = /spid="([^"]*)"/.exec(rawXml)?.[1];
  if (selectedSpid !== undefined) {
    const shape = index.shapes.get(selectedSpid);
    lines.push(
      shape === undefined
        ? `The selected markup targets shape id ${selectedSpid}, which no shape in ${hostPath} carries.`
        : `The selected markup targets shape id ${selectedSpid}, which is ${
            shape.name === null ? 'an unnamed shape' : `"${shape.name}"`
          }${shape.paragraphCount === null ? '' : ` with ${shape.paragraphCount} paragraph(s)`}.`
    );
  }

  lines.push(...renderFindings(index.problems));

  if (index.animations.length > 0 || index.triggers.length > 0) {
    unresolved.push(
      'Whether an animation that resolves actually looks right — its timing, order, direction, or whether two effects collide — is not determined here; only whether the shape it names still exists.'
    );
  }
  const unjudged = [...index.animations, ...index.triggers].filter(
    n => n.target.kind === 'shape' && n.target.exists === null && n.target.shapeId !== null
  );
  if (unjudged.length > 0 || index.builds.some(b => b.exists === null && b.shapeId !== null)) {
    unresolved.push(
      'An @spid uses the Office 2007 string form, which cannot be matched against p:cNvPr/@id, so whether it still points at a live shape is unverified.'
    );
  }
  if ([...index.animations, ...index.triggers].some(n => n.target.subShapeId !== null)) {
    unresolved.push(
      'A target narrows to a sub-shape (p:subSp) inside a group or diagram. Sub-shape ids are not p:cNvPr ids and are not resolved here.'
    );
  }
  if (
    [...index.animations, ...index.triggers].some(
      n => n.target.paragraphRange !== null && index.shapes.get(n.target.shapeId ?? '')?.paragraphCount === null
    )
  ) {
    unresolved.push(
      'A paragraph range targets a shape that is not a p:sp — a table, chart or diagram frame. Its text is not indexed the way p:pRg counts paragraphs, so the range was not checked.'
    );
  }

  return { lines, unresolved };
}
