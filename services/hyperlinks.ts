/**
 * Hyperlinks in all three formats — the link that goes nowhere and looks exactly like
 * the one that works.
 *
 * A hyperlink has no appearance of its own. The blue, underlined, styled text is a
 * character style applied to the run; the *destination* lives in an attribute nobody
 * renders. So:
 *
 *   A BROKEN LINK IS PIXEL-IDENTICAL TO A WORKING ONE.
 *
 * The document opens, the text is right, a screenshot diff is clean. It fails when
 * somebody clicks — and the person who clicks is almost never the person who could have
 * fixed it. This is the same shape as the OLE preview problem (`oleObjects.ts`) and the
 * media poster problem (`pptMedia.ts`): "it renders correctly" is not evidence the
 * reference survived.
 *
 * TWO KINDS, AND THEY FAIL DIFFERENTLY.
 *
 * **External** — an `r:id` naming a relationship with `TargetMode="External"`. The
 * honest report here is not "broken" or "fine": it is that **nothing inside the package
 * can tell you**. A URL may 404, may have been a typo from the start, may require a VPN.
 * This module reports the target and says so, and *never fetches anything*. What it can
 * still check, and does, is the packaging around the URL: an `r:id` naming a
 * relationship that does not exist, or an external relationship whose `Target` is empty,
 * are faults the package settles by itself.
 *
 * **Internal** — an anchor naming something in the same document. **This half is fully
 * checkable**, and it is the valuable one: the anchor either names a bookmark, sheet,
 * defined name or slide that is in the package, or it does not, and there is no third
 * answer that needs the network.
 *
 * ONE CONCEPT, THREE INCOMPATIBLE SPELLINGS.
 *
 * A converter cannot ask one question of all three. Where the destination lives, what it
 * is called, and what an "internal" destination even means all change per format:
 *
 *   Word         w:hyperlink/@r:id        external, via the relationship
 *                w:hyperlink/@w:anchor    internal — a BOOKMARK NAME in this document
 *                w:hyperlink/@w:docLocation  a location inside the r:id target
 *
 *   Excel        x:hyperlink/@r:id        external, via the relationship
 *                x:hyperlink/@location    internal — a CELL REFERENCE or DEFINED NAME
 *                x:hyperlink/@ref         the cell range the link covers (REQUIRED)
 *
 *   PowerPoint   a:hlinkClick/@r:id       external, via the relationship
 *                a:hlinkClick/@action     a `ppaction://` verb; a slide jump is
 *                                         `ppaction://hlinksldjump` PLUS an r:id whose
 *                                         relationship points at the slide part
 *
 * So the internal destination is a *name* in Word, a *formula-ish expression* in Excel,
 * and a *relationship to a part* in PowerPoint. Only Excel puts the link's own position
 * in the element; Word and PowerPoint get it from where the element sits in the tree.
 * There is no attribute the three share except `r:id`.
 *
 * ⚠️ THE ELEMENTS ARE NOT WHERE YOU EXPECT THEM. `w:hyperlink` is a *run container* — it
 * wraps runs the way `w:p` does — while `x:hyperlink` is a LEAF that names its cells by
 * address from a `x:hyperlinks` list at the end of the sheet, nowhere near the cell.
 * `a:hlinkClick` is DrawingML and hangs off either a shape's `a:cNvPr` (whole-shape link)
 * or a run's `a:rPr` (text link). Code that expects "the link wraps the linked thing"
 * is right in exactly one of the three formats.
 *
 * WHY AN EXTERNAL LINK PRODUCES NO FINDING.
 *
 * `problems` is for defects. "This URL cannot be checked from inside the file" is not a
 * defect — it is the honest limit of what a package can answer, and a document with
 * forty working external links would otherwise produce forty findings that mean nothing.
 * It goes to `unresolved` in the evidence instead, where it caps the confidence tier
 * rather than pretending to be damage. `hyperlinkResolves()` returns `null` for it,
 * which is deliberately NOT `false` — the same three-state discipline as
 * `oleDataIsPresent`.
 *
 * PROVENANCE — what is verified and what is not.
 *
 * Verified against the Open XML SDK's machine-readable schema data
 * (`data/schemas/schemas_openxmlformats_org_{wordprocessingml,spreadsheetml,drawingml}_2006_main.json`):
 *   • `w:CT_Hyperlink/w:hyperlink` attributes: `w:tgtFrame` (MaxLength 255),
 *     `w:tooltip` (MaxLength 260), `w:docLocation` (MaxLength 255), `w:history`,
 *     `w:anchor` (MaxLength 255), `r:id`. None is required.
 *     A second type, `w:CT_HyperlinkRuby/w:hyperlink`, carries the identical attribute
 *     set — so the same element name means two schema types, and matching on the local
 *     name (as this module does) covers both.
 *   • `w:hyperlink` parents: `w:p`, `w:customXml`, `w:fldSimple`, `w:hyperlink` (it
 *     nests), `w:bdo`, `w:dir`, `w:sdtContent`.
 *   • `x:CT_Hyperlink/x:hyperlink` is a LEAF element; attributes `ref` (**required**),
 *     `r:id`, `location`, `tooltip`, `display` — all unprefixed except `r:id`.
 *     Its only parent is `x:hyperlinks`, whose only parent is `x:worksheet`.
 *   • `a:CT_Hyperlink` is the shared type behind `a:hlinkClick`, `a:hlinkMouseOver` and
 *     `a:hlinkHover`. Attributes: `r:id`, `invalidUrl`, `action`, `tgtFrame`, `tooltip`,
 *     `history`, `highlightClick`, `endSnd` — all unprefixed except `r:id`, and **none
 *     required**, so an action-only link with no relationship is legal markup.
 *   • `a:hlinkClick` parents: `a:cNvPr` (`a:CT_NonVisualDrawingProps`, which is also the
 *     type of `p:cNvPr`) and `a:rPr` / `a:defRPr` / `a:endParaRPr`.
 *   • The SDK documents `@invalidUrl` as "In case the url is invalid so we can't create
 *     a relationship, we'll save it here, r:id will point to a NULL one" — which is why
 *     an EMPTY `r:id` is handled here as "names no relationship" rather than as a
 *     dangling one. An empty id is a producer saying "there is nothing to point at".
 *
 * Verified against Open XML SDK source
 * (`src/DocumentFormat.OpenXml.Framework/Packaging/HyperlinkRelationship.cs`):
 *   • the hyperlink relationship type is
 *     `http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink`.
 *     This is the one relationship Type URI this module matches on, and it is matched
 *     only for the "declared but referenced by nothing" check — every destination is
 *     resolved by `Id` alone, exactly as `oleObjects.ts` does.
 *
 * ⚠️ NOT VERIFIED AGAINST ANY NORMATIVE SOURCE — the `ppaction://` verb vocabulary.
 * In the schema `@action` is an untyped string with no enumeration, so the SDK cannot
 * confirm any verb. The table in `PPT_ACTIONS` was cross-checked against two independent
 * implementations (python-pptx's `pptx/action.py`, LibreOffice's `oox`
 * `hyperlinkcontext.cxx`) and NOT against ECMA-376 or MS-ODRAWXML. It is therefore used
 * only to *add* a check for verbs it knows; an unrecognised verb is reported verbatim
 * and never called wrong.
 *
 * ⚠️ NOT VERIFIED — which destination wins when a `w:hyperlink` carries BOTH `r:id` and
 * `w:anchor`. The schema permits it and states no precedence; no observed Word behaviour
 * was confirmed. It is reported as ambiguous rather than resolved in either direction.
 *
 * ⚠️ NOT VERIFIED — whether Word resolves `w:anchor` against bookmarks in *other* parts
 * (a link in a header aiming at a bookmark in the body). To avoid a confident wrong
 * answer, an anchor is looked up across every Word body part in the package and only
 * called dangling when it is in none of them.
 *
 * ⚠️ NOT VERIFIED — what Excel does with an `x:hyperlink` missing its required `@ref`
 * (repair, ignore, or refuse to open). It is reported as an error but NOT marked silent,
 * because `silent` claims the file still renders exactly as intended and that claim
 * cannot be supported here.
 *
 * NAMESPACES ARE COMPARED BY EXACT EQUALITY. `conformance.ts` rewrites ISO Strict URIs to
 * their Transitional equivalents before any analyzer runs, so Strict-tolerant matching
 * here would be dead code that looks load-bearing.
 *
 * NOT THIS MODULE'S JOB: the `HYPERLINK` *field* (`{ HYPERLINK "http://…" \l "Anchor" }`),
 * which is a different encoding of the same idea and belongs to `wordFields.ts`, where
 * it is already cross-checked against bookmarks. The two never overlap: a field is
 * `w:instrText`, an element hyperlink is `w:hyperlink`.
 */

import { relsPathFor, resolveTarget, type PackageParts } from './packageIntegrity';
import { finding, renderFindings, type Finding, type Severity } from './findings';
import { readBookmarks, MAX_BOOKMARK_NAME_LENGTH } from './wordBookmarks';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const X_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const HYPERLINK_RELATIONSHIP_TYPE = `${R_NS}/hyperlink`;

/**
 * Word's destination for "the top of this document".
 *
 * ⚠️ OBSERVED, NOT SCHEMA. `_top` is not a bookmark and no `w:bookmarkStart` will ever
 * declare it, so checking it against the bookmark index would report a dangling anchor on
 * every document that uses Word's own "Place in This Document → Top of the Document".
 * Suppressing it costs nothing — `_top` is never a real bookmark name, so there is no
 * true positive to lose — while reporting it would be a confident wrong answer.
 */
const WORD_RESERVED_ANCHORS = new Set(['_top']);

/**
 * `ppaction://` verbs, and whether the verb needs a relationship to mean anything.
 *
 * ⚠️ NOT from any normative source — see the header. Used only to add checks for verbs
 * that are recognised; anything else is reported verbatim and not judged.
 */
const PPT_ACTIONS: Readonly<Record<string, { needsRelationship: boolean; description: string }>> = {
  hlinksldjump: { needsRelationship: true, description: 'jumps to a specific slide in this presentation' },
  hlinkshowjump: { needsRelationship: false, description: 'jumps to a relative position in the show (next, previous, first, last, end)' },
  hlinkpres: { needsRelationship: true, description: 'opens a slide in another presentation' },
  hlinkfile: { needsRelationship: true, description: 'opens a file on the same computer' },
  customshow: { needsRelationship: false, description: 'starts a custom slide show' },
  ole: { needsRelationship: true, description: 'performs an OLE verb on an embedded object' },
  macro: { needsRelationship: false, description: 'runs an embedded VBA macro' },
  program: { needsRelationship: true, description: 'launches an external program' }
};

/**
 * Severity and silence per kind, decided once here.
 *
 * Almost every one of these is SILENT, and that is the point of the analyzer: a link's
 * destination has no rendering, so breaking it changes nothing anyone can see. The two
 * exceptions are `missing-cell-range` — where what Excel does is unverified, so the
 * safer under-claim is "not silent" — and `unreferenced-relationship`, which is a note
 * about dead weight rather than damage.
 *
 * `relationship-not-external` is a warning rather than an error because it is ambiguous
 * by nature: OPC says a relationship without `TargetMode` resolves as a path inside the
 * package, and a producer that forgot the attribute and a producer that genuinely meant
 * an in-package target write identical markup.
 */
const HYPERLINK_RULES = {
  'relationship-missing':      { severity: 'error',   silent: true },
  'empty-external-target':     { severity: 'error',   silent: true },
  'internal-target-missing':   { severity: 'error',   silent: true },
  'dangling-anchor':           { severity: 'error',   silent: true },
  'dangling-location':         { severity: 'error',   silent: true },
  'action-needs-relationship': { severity: 'error',   silent: true },
  'no-destination':            { severity: 'error',   silent: true },
  'missing-cell-range':        { severity: 'error',   silent: false },
  'relationship-not-external': { severity: 'warning', silent: true },
  'ambiguous-destination':     { severity: 'warning', silent: true },
  'slide-jump-off-target':     { severity: 'warning', silent: true },
  'invalid-url':               { severity: 'warning', silent: true },
  'empty-cell-range':          { severity: 'warning', silent: true },
  'unreferenced-relationship': { severity: 'note',    silent: true }
} as const satisfies Record<string, { severity: Severity; silent: boolean }>;

export type HyperlinkProblemKind = keyof typeof HYPERLINK_RULES;

const hyperlinkFinding = (
  kind: HyperlinkProblemKind,
  part: string,
  message: string,
  remediation: string,
  subject?: Record<string, string>
): Finding => finding(`hyperlink/${kind}`, part, message, remediation, { ...HYPERLINK_RULES[kind], subject });

export type HyperlinkFormat = 'word' | 'excel' | 'powerpoint';

/**
 * What the link points at.
 *
 * `external` and `internal` are the two the brief cares about. `action` is PowerPoint's
 * third case — a `ppaction://` verb that is neither a URL nor a name — and `none` is a
 * hyperlink element that states no destination at all.
 */
export type HyperlinkKind = 'external' | 'internal' | 'action' | 'none';

export interface Hyperlink {
  format: HyperlinkFormat;
  /** The package part this link was found in. Relationships resolve relative to it. */
  part: string;
  element: Element;
  /** `w:hyperlink`, `x:hyperlink`, `a:hlinkClick`, `a:hlinkMouseOver`, `a:hlinkHover`. */
  label: string;
  kind: HyperlinkKind;
  /** Which attribute the destination was read from — differs per format, so state it. */
  destinationEvidence: string;
  /**
   * `@r:id` verbatim. `''` is a real, distinct value: PowerPoint writes an empty id
   * alongside `@invalidUrl` to mean "there is deliberately no relationship here".
   */
  relationshipId: string | null;
  /** The internal destination as written: a bookmark name, or an Excel `@location`. */
  anchor: string | null;
  /** `a:hlinkClick/@action` verbatim, or null. */
  action: string | null;
  /** The URI a `TargetMode="External"` relationship carries. Never fetched. */
  externalTarget: string | null;
  /** Resolved part path when the relationship is NOT external. */
  internalTarget: string | null;
  /** `w:docLocation` — a position inside the r:id target, not in this document. */
  documentLocation: string | null;
  tooltip: string | null;
  targetFrame: string | null;
  /** Excel only: `@ref`, the cell range the link covers. Required by the schema. */
  cellRange: string | null;
  /**
   * Whether the destination resolves. THREE STATES, and the third is the whole point.
   *
   * `true`  — checked, and the thing it names is in the package.
   * `false` — checked, and it is not.
   * `null`  — CANNOT BE CHECKED from inside the package. An external URL is always this,
   *           and so is an internal destination whose index part (`xl/workbook.xml`) is
   *           itself absent. "We cannot check" and "it is broken" are different answers
   *           and only one of them is a defect.
   */
  destinationResolves: boolean | null;
  problems: Finding[];
}

interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

const parseXml = (xml: string): Document | null => {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
};

const descendants = (root: ParentNode, ns: string, names: ReadonlySet<string>): Element[] =>
  Array.from(root.querySelectorAll('*')).filter(el => el.namespaceURI === ns && names.has(el.localName));

/**
 * `@r:id` as written, preserving the difference between absent (`null`) and empty (`''`).
 * `getAttributeNS` collapses neither, but `?? null` on a `''` would — so this is spelled
 * out rather than defaulted.
 */
const relId = (el: Element): string | null => el.getAttributeNS(R_NS, 'id');

const readRelationships = (parts: PackageParts, ownerPart: string): Map<string, Relationship> | null => {
  const relsXml = parts[relsPathFor(ownerPart)];
  if (relsXml === undefined) return null;
  const doc = parseXml(relsXml);
  if (!doc) return null;

  const map = new Map<string, Relationship>();
  for (const rel of Array.from(doc.getElementsByTagName('Relationship'))) {
    const id = rel.getAttribute('Id');
    if (id === null) continue;
    map.set(id, {
      id,
      type: rel.getAttribute('Type') ?? '',
      target: rel.getAttribute('Target') ?? '',
      external: (rel.getAttribute('TargetMode') ?? '').toLowerCase() === 'external'
    });
  }
  return map;
};

/** Word body parts, each carrying its own bookmarks. */
const WORD_BODY_PART = /^word\/(?:document\d*|header[^/]*|footer[^/]*|footnotes\d*|endnotes\d*|comments\d*)\.xml$/;

/**
 * Every bookmark name declared anywhere in the package's Word body parts.
 *
 * Deliberately package-wide rather than part-local: see the header. Whether Word actually
 * follows an anchor from a header into the body was not confirmed, so the wider set is
 * used — it can only cause a missed report, never a false one.
 */
const allBookmarkNames = (parts: PackageParts): Set<string> => {
  const names = new Set<string>();
  for (const [path, xml] of Object.entries(parts)) {
    if (!WORD_BODY_PART.test(path)) continue;
    const doc = parseXml(xml);
    if (!doc?.documentElement) continue;
    for (const bookmark of readBookmarks(doc, path).bookmarks) {
      if (bookmark.name !== '') names.add(bookmark.name);
    }
  }
  return names;
};

/** What the workbook says exists, for checking an Excel `@location`. */
interface WorkbookIndex {
  sheetNames: Set<string>;
  definedNames: Set<string>;
}

/** `null` when there is no `xl/workbook.xml` — cannot check, which is not the same as broken. */
const readWorkbookIndex = (parts: PackageParts): WorkbookIndex | null => {
  const xml = parts['xl/workbook.xml'];
  if (xml === undefined) return null;
  const doc = parseXml(xml);
  if (!doc?.documentElement) return null;

  const sheetNames = new Set<string>();
  for (const sheet of descendants(doc.documentElement, X_NS, new Set(['sheet']))) {
    const name = sheet.getAttribute('name');
    if (name !== null) sheetNames.add(name);
  }
  const definedNames = new Set<string>();
  for (const defined of descendants(doc.documentElement, X_NS, new Set(['definedName']))) {
    const name = defined.getAttribute('name');
    if (name !== null) definedNames.add(name);
  }
  return { sheetNames, definedNames };
};

/** `A1` / `$A$1` → 1-based column and row, or null when it is not a cell address. */
const parseCellAddress = (raw: string): { column: number; row: number } | null => {
  const match = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/.exec(raw);
  if (!match) return null;
  const letters = match[1].toUpperCase();
  let column = 0;
  for (const character of letters) column = column * 26 + (character.charCodeAt(0) - 64);
  return { column, row: Number(match[2]) };
};

interface CellRectangle {
  left: number;
  right: number;
  top: number;
  bottom: number;
}

/**
 * `@ref` as rectangles. `ST_Ref` is a single range, but a space-separated list is parsed
 * too so that a producer writing `ST_Sqref` into the attribute is read rather than
 * silently treated as a broken single range.
 */
const parseRanges = (ref: string): CellRectangle[] => {
  const rectangles: CellRectangle[] = [];
  for (const piece of ref.trim().split(/\s+/)) {
    if (piece === '') continue;
    const [fromRaw, toRaw] = piece.split(':');
    const from = parseCellAddress(fromRaw);
    const to = toRaw === undefined ? from : parseCellAddress(toRaw);
    if (!from || !to) continue;
    rectangles.push({
      left: Math.min(from.column, to.column),
      right: Math.max(from.column, to.column),
      top: Math.min(from.row, to.row),
      bottom: Math.max(from.row, to.row)
    });
  }
  return rectangles;
};

/**
 * Does any cell inside `ref` actually carry content?
 *
 * "Content" is a `x:v` (value), `x:is` (inline string) or `x:f` (formula) child. A bare
 * `<x:c r="A1" s="3"/>` is a *styled* empty cell and does not count — it is exactly what
 * is left behind when someone deletes the text a link was attached to.
 *
 * Returns `null` when the question cannot be asked: a `@ref` that does not parse as a
 * range gives no rectangle to test, and reporting "no content" for it would be an answer
 * to a question that was never posed.
 */
const rangeCoversContent = (worksheet: Element, ref: string): boolean | null => {
  const rectangles = parseRanges(ref);
  if (rectangles.length === 0) return null;

  const contentBearing = new Set(['v', 'is', 'f']);
  for (const cell of descendants(worksheet, X_NS, new Set(['c']))) {
    const address = cell.getAttribute('r');
    if (address === null) continue;
    const at = parseCellAddress(address);
    if (!at) continue;
    if (!rectangles.some(r => at.column >= r.left && at.column <= r.right && at.row >= r.top && at.row <= r.bottom)) {
      continue;
    }
    const hasContent = Array.from(cell.children).some(
      child => child.namespaceURI === X_NS && contentBearing.has(child.localName)
    );
    if (hasContent) return true;
  }
  return false;
};

const formatOf = (partPath: string): HyperlinkFormat =>
  partPath.startsWith('xl/') ? 'excel' : partPath.startsWith('ppt/') ? 'powerpoint' : 'word';

/** Everything the relationship half of a destination settles, and what it could not. */
interface ResolvedRelationship {
  externalTarget: string | null;
  internalTarget: string | null;
  /** null when there is no relationship reference to resolve at all. */
  resolves: boolean | null;
}

/**
 * Resolves an `r:id` and says which link in the chain broke.
 *
 * The external case returns `resolves: null` on purpose: the relationship is intact and
 * the URI is well-formed as far as the package is concerned, and whether it *works* is
 * not a question a zip file can answer.
 */
const resolveRelationship = (
  parts: PackageParts,
  ownerPart: string,
  rels: Map<string, Relationship> | null,
  id: string,
  label: string,
  problems: Finding[]
): ResolvedRelationship => {
  const rel = rels?.get(id);
  if (rel === undefined) {
    problems.push(hyperlinkFinding(
      'relationship-missing', ownerPart,
      rels === null
        ? `${label} references relationship "${id}", but ${ownerPart} has no relationship part at all, so the destination cannot be resolved. The link is styled and underlined exactly like a working one.`
        : `${label} references relationship "${id}", which ${relsPathFor(ownerPart)} does not declare. The link has no destination — and nothing about how it renders says so.`,
      `Add a Relationship with Id="${id}" to ${relsPathFor(ownerPart)}, or remove the reference to it.`,
      { relationshipId: id }
    ));
    return { externalTarget: null, internalTarget: null, resolves: false };
  }

  if (rel.external) {
    if (rel.target === '') {
      // Checkable, and a real fault: the package says "the destination is out there"
      // and then names nothing at all.
      problems.push(hyperlinkFinding(
        'empty-external-target', ownerPart,
        `${label} points at relationship "${id}", which is marked TargetMode="External" and carries an empty Target. There is no URI to open, so the link is dead in a way the package itself settles — no network access needed to know it.`,
        `Set the Target of relationship "${id}" to the intended URI, or delete the link.`,
        { relationshipId: id }
      ));
      return { externalTarget: '', internalTarget: null, resolves: false };
    }
    // Reported, never fetched. See the header.
    return { externalTarget: rel.target, internalTarget: null, resolves: null };
  }

  const target = resolveTarget(ownerPart, rel.target);
  const exists = parts[target] !== undefined;
  return { externalTarget: null, internalTarget: target, resolves: exists };
};

/** Word: `w:hyperlink`, whose internal destination is a bookmark NAME. */
const readWordHyperlink = (
  el: Element,
  parts: PackageParts,
  ownerPart: string,
  rels: Map<string, Relationship> | null,
  bookmarkNames: Set<string>
): Hyperlink => {
  const problems: Finding[] = [];
  const id = relId(el);
  const anchor = el.getAttributeNS(W_NS, 'anchor');
  const docLocation = el.getAttributeNS(W_NS, 'docLocation');

  const relationship =
    id !== null && id !== ''
      ? resolveRelationship(parts, ownerPart, rels, id, 'w:hyperlink', problems)
      : { externalTarget: null, internalTarget: null, resolves: null as boolean | null };

  if (id !== null && id !== '' && relationship.internalTarget !== null) {
    reportNonExternal(problems, ownerPart, 'w:hyperlink', id, relationship);
  }

  let anchorResolves: boolean | null = null;
  if (anchor !== null && anchor !== '' && !WORD_RESERVED_ANCHORS.has(anchor)) {
    anchorResolves = bookmarkNames.has(anchor);
    if (!anchorResolves) {
      // The whole reason this analyzer exists, in one finding.
      const tooLong =
        anchor.length > MAX_BOOKMARK_NAME_LENGTH
          ? ` The anchor is ${anchor.length} characters, and a bookmark name cannot exceed ${MAX_BOOKMARK_NAME_LENGTH}, so no bookmark could ever have carried this name.`
          : '';
      problems.push(hyperlinkFinding(
        'dangling-anchor', ownerPart,
        `w:hyperlink jumps to bookmark "${anchor}", and no bookmark of that name is declared anywhere in this package. The link text is still blue and underlined, the document still opens, and clicking does nothing.${tooLong}`,
        `Restore a bookmark named "${anchor}", or repoint w:anchor at a bookmark that exists.`,
        { anchor }
      ));
    }
  }

  const hasRelationship = id !== null && id !== '';
  if (hasRelationship && anchor !== null && anchor !== '') {
    problems.push(hyperlinkFinding(
      'ambiguous-destination', ownerPart,
      `w:hyperlink carries BOTH r:id="${id}" (a relationship destination) and w:anchor="${anchor}" (a bookmark in this document). The schema permits both and states no precedence, and WordprocessingML already has a separate attribute — w:docLocation — for naming a position inside the r:id target, so this is not the encoding for "that file, at that spot". Which destination a consumer follows is undetermined.`,
      'Keep r:id for an external target (with w:docLocation if a position inside it is needed), or w:anchor for an internal one — not both.',
      { relationshipId: id, anchor }
    ));
  }

  if (!hasRelationship && (anchor === null || anchor === '')) {
    problems.push(hyperlinkFinding(
      'no-destination', ownerPart,
      'w:hyperlink states no destination: it has neither r:id nor w:anchor. The runs inside it are still styled as a hyperlink, so the text reads as a link and behaves as ordinary text.',
      'Add r:id for an external target or w:anchor for a bookmark in this document, or unwrap the runs.'
    ));
  }

  const kind: HyperlinkKind = hasRelationship
    ? relationship.externalTarget !== null
      ? 'external'
      : 'internal'
    : anchor !== null && anchor !== ''
      ? 'internal'
      : 'none';

  return {
    format: formatOf(ownerPart),
    part: ownerPart,
    element: el,
    label: 'w:hyperlink',
    kind,
    destinationEvidence: hasRelationship ? 'w:hyperlink/@r:id' : 'w:hyperlink/@w:anchor',
    relationshipId: id,
    anchor,
    action: null,
    externalTarget: relationship.externalTarget,
    internalTarget: relationship.internalTarget,
    documentLocation: docLocation,
    tooltip: el.getAttributeNS(W_NS, 'tooltip'),
    targetFrame: el.getAttributeNS(W_NS, 'tgtFrame'),
    cellRange: null,
    destinationResolves: combineResolution(hasRelationship ? relationship.resolves : null, anchorResolves, kind),
    problems
  };
};

/** Excel: `x:hyperlink`, a leaf that names its cells by address from the end of the sheet. */
const readExcelHyperlink = (
  el: Element,
  parts: PackageParts,
  ownerPart: string,
  rels: Map<string, Relationship> | null,
  worksheet: Element,
  workbook: WorkbookIndex | null
): Hyperlink => {
  const problems: Finding[] = [];
  const id = relId(el);
  // Unprefixed: verified against x:CT_Hyperlink, where only r:id is namespaced.
  const location = el.getAttribute('location');
  const ref = el.getAttribute('ref');

  const relationship =
    id !== null && id !== ''
      ? resolveRelationship(parts, ownerPart, rels, id, 'x:hyperlink', problems)
      : { externalTarget: null, internalTarget: null, resolves: null as boolean | null };

  if (id !== null && id !== '' && relationship.internalTarget !== null) {
    reportNonExternal(problems, ownerPart, 'x:hyperlink', id, relationship);
  }

  let locationResolves: boolean | null = null;
  if (location !== null && location !== '') {
    locationResolves = checkExcelLocation(location, workbook, ownerPart, problems);
  }

  const hasRelationship = id !== null && id !== '';
  if (!hasRelationship && (location === null || location === '')) {
    problems.push(hyperlinkFinding(
      'no-destination', ownerPart,
      `x:hyperlink over ${ref === null ? 'an unstated range' : `"${ref}"`} states no destination: it has neither r:id nor @location. The cells keep their hyperlink styling and clicking them goes nowhere.`,
      'Add r:id for an external target or @location for a place in this workbook, or delete the hyperlink entry.'
    ));
  }

  if (ref === null) {
    problems.push(hyperlinkFinding(
      'missing-cell-range', ownerPart,
      'x:hyperlink has no @ref, which the schema makes required. Nothing states which cells the link covers, so it is attached to no part of the sheet.',
      'Add @ref naming the cell or range the hyperlink applies to.'
    ));
  } else if (rangeCoversContent(worksheet, ref) === false) {
    problems.push(hyperlinkFinding(
      'empty-cell-range', ownerPart,
      `x:hyperlink covers "${ref}", and no cell in that range holds a value, an inline string or a formula. The link is still live — Excel makes the region clickable — but there is nothing on the sheet to click, which is what a hyperlink looks like after the text it was attached to was deleted.`,
      `Put the link text back in ${ref}, or remove the hyperlink entry.`,
      { ref }
    ));
  }

  const kind: HyperlinkKind = hasRelationship
    ? relationship.externalTarget !== null
      ? 'external'
      : 'internal'
    : location !== null && location !== ''
      ? 'internal'
      : 'none';

  return {
    format: 'excel',
    part: ownerPart,
    element: el,
    label: 'x:hyperlink',
    kind,
    destinationEvidence: hasRelationship ? 'x:hyperlink/@r:id' : 'x:hyperlink/@location',
    relationshipId: id,
    anchor: location,
    action: null,
    externalTarget: relationship.externalTarget,
    internalTarget: relationship.internalTarget,
    documentLocation: null,
    tooltip: el.getAttribute('tooltip'),
    targetFrame: null,
    cellRange: ref,
    destinationResolves: combineResolution(hasRelationship ? relationship.resolves : null, locationResolves, kind),
    problems
  };
};

/**
 * An Excel `@location` against what the workbook declares.
 *
 * Three shapes, and only two are checkable:
 *   `Sheet1!A1` / `'My Sheet'!$A$1`  → the sheet must exist. `#REF!A1` lands here too,
 *                                      and correctly fails.
 *   `A1`                             → a cell in the same sheet; always in range, nothing
 *                                      to check, so this reports nothing.
 *   `Total_Sales`                    → a defined name, which must exist.
 *
 * Returns `null` when there is no `xl/workbook.xml` to check against — not `false`.
 */
const checkExcelLocation = (
  location: string,
  workbook: WorkbookIndex | null,
  ownerPart: string,
  problems: Finding[]
): boolean | null => {
  if (workbook === null) return null;

  const bang = location.lastIndexOf('!');
  if (bang !== -1) {
    const raw = location.slice(0, bang);
    // A sheet name containing spaces or punctuation is single-quoted, and a literal
    // quote inside it is doubled.
    const sheet = raw.startsWith("'") && raw.endsWith("'") && raw.length >= 2
      ? raw.slice(1, -1).replace(/''/g, "'")
      : raw;
    if (workbook.sheetNames.has(sheet)) return true;
    problems.push(hyperlinkFinding(
      'dangling-location', ownerPart,
      `x:hyperlink jumps to "${location}", and this workbook has no sheet named "${sheet}". The cell still shows hyperlink styling; clicking it produces a "Reference is not valid" box and nothing else changes.`,
      `Restore the sheet named "${sheet}", or repoint @location at a sheet that exists.`,
      { location, sheet }
    ));
    return false;
  }

  // A bare cell address is a position in the sheet the link already lives on.
  if (parseCellAddress(location) !== null) return true;

  if (workbook.definedNames.has(location)) return true;
  problems.push(hyperlinkFinding(
    'dangling-location', ownerPart,
    `x:hyperlink jumps to "${location}", which is neither a cell address nor a defined name in this workbook. The cell looks like a working link and clicking it goes nowhere.`,
    `Define a name "${location}" in xl/workbook.xml, or repoint @location at an address or a name that exists.`,
    { location }
  ));
  return false;
};

/** PowerPoint and other DrawingML: `a:hlinkClick` and its mouse-over siblings. */
const readDrawingHyperlink = (
  el: Element,
  parts: PackageParts,
  ownerPart: string,
  rels: Map<string, Relationship> | null
): Hyperlink => {
  const problems: Finding[] = [];
  const label = `a:${el.localName}`;
  const id = relId(el);
  // Unprefixed: verified against a:CT_Hyperlink, where only r:id is namespaced.
  const action = el.getAttribute('action');
  const invalidUrl = el.getAttribute('invalidUrl');

  const relationship =
    id !== null && id !== ''
      ? resolveRelationship(parts, ownerPart, rels, id, label, problems)
      : { externalTarget: null, internalTarget: null, resolves: null as boolean | null };

  // The empty r:id is not a dangling reference — the schema documents it as what a
  // producer writes when it could not build a relationship for the URL it was given.
  if (invalidUrl !== null && invalidUrl !== '') {
    problems.push(hyperlinkFinding(
      'invalid-url', ownerPart,
      `${label} carries @invalidUrl="${invalidUrl}" — the producing application could not turn that string into a relationship, so it stored it here and left r:id empty. The shape or run is still styled as a link and has no destination.`,
      `Correct "${invalidUrl}" to a well-formed URI and write it as an external relationship, or remove the hyperlink.`,
      { invalidUrl }
    ));
  }

  const verb = action === null ? null : /^ppaction:\/\/([A-Za-z]+)/.exec(action)?.[1] ?? null;
  const known = verb === null ? undefined : PPT_ACTIONS[verb];
  const hasRelationship = id !== null && id !== '';

  if (known !== undefined && known.needsRelationship && !hasRelationship) {
    problems.push(hyperlinkFinding(
      'action-needs-relationship', ownerPart,
      `${label} declares action "${action}", which ${known.description} and therefore needs an r:id naming what to open — and it has none. The click is a no-op, and the shape looks exactly as it would if it worked.`,
      `Add r:id pointing at the relationship for the ${verb === 'hlinksldjump' ? 'target slide' : 'target'}, or drop the action.`,
      { action: action ?? '' }
    ));
  }

  // A slide jump is the fully checkable PowerPoint case: the relationship has to land on
  // a slide part that is actually in the package.
  if (verb === 'hlinksldjump' && relationship.internalTarget !== null) {
    if (relationship.resolves === false) {
      problems.push(hyperlinkFinding(
        'internal-target-missing', ownerPart,
        `${label} jumps to slide "${relationship.internalTarget}", which is not in the package. The shape still renders and the deck still opens; the jump silently does nothing.`,
        `Restore ${relationship.internalTarget}, or repoint the relationship at a slide that exists.`,
        { target: relationship.internalTarget }
      ));
    } else if (!/^ppt\/slides\/[^/]+\.xml$/.test(relationship.internalTarget)) {
      problems.push(hyperlinkFinding(
        'slide-jump-off-target', ownerPart,
        `${label} declares a slide jump but its relationship resolves to "${relationship.internalTarget}", which is not a slide part. A jump to a layout, a master or a notes page is not a destination the show can navigate to.`,
        'Point the relationship at a part under ppt/slides/.',
        { target: relationship.internalTarget }
      ));
    }
  } else if (hasRelationship && relationship.internalTarget !== null) {
    reportNonExternal(problems, ownerPart, label, id, relationship);
  }

  if (!hasRelationship && action === null && (invalidUrl === null || invalidUrl === '')) {
    problems.push(hyperlinkFinding(
      'no-destination', ownerPart,
      `${label} states no destination: no r:id and no @action. The shape or run keeps its hyperlink styling and clicking it does nothing.`,
      'Add r:id for an external target, or @action for a jump inside the presentation, or remove the element.'
    ));
  }

  const kind: HyperlinkKind =
    verb === 'hlinksldjump'
      ? 'internal'
      : action !== null
        ? 'action'
        : hasRelationship
          ? relationship.externalTarget !== null
            ? 'external'
            : 'internal'
          : 'none';

  return {
    format: formatOf(ownerPart),
    part: ownerPart,
    element: el,
    label,
    kind,
    destinationEvidence: action !== null ? `${label}/@action` : `${label}/@r:id`,
    relationshipId: id,
    anchor: null,
    action,
    externalTarget: relationship.externalTarget,
    internalTarget: relationship.internalTarget,
    documentLocation: null,
    tooltip: el.getAttribute('tooltip'),
    targetFrame: el.getAttribute('tgtFrame'),
    cellRange: null,
    destinationResolves:
      kind === 'external' ? null : kind === 'none' ? false : hasRelationship ? relationship.resolves : null,
    problems
  };
};

/**
 * A hyperlink relationship without `TargetMode="External"`.
 *
 * OPC resolves such a Target as a path inside the package, so the "link" points at a part
 * rather than at the web. A producer that forgot the attribute and one that genuinely
 * meant an in-package destination write identical markup, which is why this is a warning
 * that describes the situation rather than an error that assumes intent.
 */
const reportNonExternal = (
  problems: Finding[],
  ownerPart: string,
  label: string,
  id: string,
  relationship: ResolvedRelationship
): void => {
  const target = relationship.internalTarget;
  if (target === null) return;
  if (relationship.resolves === false) {
    problems.push(hyperlinkFinding(
      'internal-target-missing', ownerPart,
      `${label} points at relationship "${id}", which has no TargetMode="External" and so resolves to the package path "${target}" — and no such part is in the package. The link renders normally and has nothing behind it.`,
      `Add TargetMode="External" to relationship "${id}" if the target is a URI, or restore ${target}.`,
      { relationshipId: id, target }
    ));
  } else {
    problems.push(hyperlinkFinding(
      'relationship-not-external', ownerPart,
      `${label} points at relationship "${id}", which carries no TargetMode="External". OPC therefore resolves it as a path inside the package ("${target}") rather than as a URI. If a web or file address was intended, every consumer will follow the wrong thing, and the link looks identical either way.`,
      `Add TargetMode="External" to relationship "${id}" if it names a URI; leave it as it is if an in-package part really was meant.`,
      { relationshipId: id, target }
    ));
  }
};

/**
 * Merges the relationship answer and the internal-anchor answer into one three-state
 * verdict.
 *
 * `false` wins over everything — one broken half is a broken link. `null` beats `true`
 * only when nothing was actually checked, so a link with a working anchor and an
 * unverifiable URL reports `null`: the honest answer is that it is not fully confirmed.
 */
const combineResolution = (
  fromRelationship: boolean | null,
  fromAnchor: boolean | null,
  kind: HyperlinkKind
): boolean | null => {
  if (fromRelationship === false || fromAnchor === false) return false;
  if (kind === 'none') return false;
  if (fromRelationship === null && fromAnchor === null) return null;
  if (fromRelationship === null && kind === 'external') return null;
  return true;
};

const WORD_LINK_ELEMENTS = new Set(['hyperlink']);
const EXCEL_LINK_ELEMENTS = new Set(['hyperlink']);
const DRAWING_LINK_ELEMENTS = new Set(['hlinkClick', 'hlinkMouseOver', 'hlinkHover']);

/**
 * Every hyperlink in one part, whichever format it belongs to.
 *
 * `partPath` must be the part the XML came from — relationships resolve relative to it,
 * and a hyperlink's `r:id` means nothing without it.
 *
 * All three element families are scanned in every part on purpose: a `.docx` can contain
 * `a:hlinkClick` on a shape inside `word/document.xml`, and an `.xlsx` can contain one
 * inside `xl/drawings/drawing1.xml`. Keying the scan off the file extension would miss
 * exactly the links that are hardest to find by hand.
 */
export function readHyperlinks(parts: PackageParts, partPath: string): Hyperlink[] {
  const xml = parts[partPath];
  if (xml === undefined) return [];
  const doc = parseXml(xml);
  if (!doc?.documentElement) return [];

  const root = doc.documentElement;
  const rels = readRelationships(parts, partPath);
  const found: Hyperlink[] = [];

  const wordLinks = descendants(root, W_NS, WORD_LINK_ELEMENTS);
  if (wordLinks.length > 0) {
    const bookmarkNames = allBookmarkNames(parts);
    for (const el of wordLinks) found.push(readWordHyperlink(el, parts, partPath, rels, bookmarkNames));
  }

  const excelLinks = descendants(root, X_NS, EXCEL_LINK_ELEMENTS);
  if (excelLinks.length > 0) {
    const workbook = readWorkbookIndex(parts);
    for (const el of excelLinks) found.push(readExcelHyperlink(el, parts, partPath, rels, root, workbook));
  }

  for (const el of descendants(root, A_NS, DRAWING_LINK_ELEMENTS)) {
    found.push(readDrawingHyperlink(el, parts, partPath, rels));
  }

  return found;
}

/**
 * Hyperlink relationships the part declares and nothing in it uses.
 *
 * This is what a deleted link leaves behind: the markup goes, the relationship stays. It
 * is dead weight rather than damage — hence a note — but it is also the fingerprint of an
 * edit that removed a link, which is worth being able to see when comparing two files.
 *
 * References are gathered from EVERY `r:`-namespace attribute in the part, not only from
 * hyperlink elements, so a relationship consumed by some other construct is never
 * mistaken for an orphan.
 */
const unreferencedRelationshipFindings = (parts: PackageParts, partPath: string): Finding[] => {
  const xml = parts[partPath];
  if (xml === undefined) return [];
  const doc = parseXml(xml);
  if (!doc?.documentElement) return [];
  const rels = readRelationships(parts, partPath);
  if (rels === null) return [];

  const used = new Set<string>();
  for (const el of Array.from(doc.documentElement.querySelectorAll('*'))) {
    for (const attribute of Array.from(el.attributes)) {
      if (attribute.namespaceURI === R_NS) used.add(attribute.value);
    }
  }

  const findings: Finding[] = [];
  for (const rel of rels.values()) {
    if (rel.type !== HYPERLINK_RELATIONSHIP_TYPE || used.has(rel.id)) continue;
    findings.push(hyperlinkFinding(
      'unreferenced-relationship', partPath,
      `${relsPathFor(partPath)} declares hyperlink relationship "${rel.id}" targeting "${rel.target}", and nothing in ${partPath} references it. Nothing is broken by it, but this is what is left behind when a link is deleted from the markup and its relationship is not.`,
      `Delete relationship "${rel.id}", or restore the markup that used to reference it.`,
      { relationshipId: rel.id, target: rel.target }
    ));
  }
  return findings;
};

/** Every finding this module produces for one part. */
export function hyperlinkFindings(parts: PackageParts, partPath: string): Finding[] {
  return [
    ...readHyperlinks(parts, partPath).flatMap(link => link.problems),
    ...unreferencedRelationshipFindings(parts, partPath)
  ];
}

/**
 * The links that render correctly and are broken anyway.
 *
 * For hyperlinks this is very nearly all of them, which is the point: `silent` is what
 * separates a link that fails on click from one whose failure is visible on the page, and
 * a link's destination is never on the page.
 */
export function findSilentlyBrokenHyperlinks(links: readonly Hyperlink[]): Hyperlink[] {
  return links.filter(link => link.problems.some(problem => problem.silent));
}

/**
 * Whether the destination resolves, stated plainly.
 *
 * `null` means unknowable from the package alone — an external URL always is, because
 * answering would mean leaving the file, which this code does not do. That is deliberately
 * not `false`: "we cannot check" and "it is broken" are different answers and only one of
 * them is a defect.
 */
export function hyperlinkResolves(link: Hyperlink): boolean | null {
  return link.destinationResolves;
}

/** Body parts that can carry a hyperlink, in any of the three formats. */
export const HYPERLINK_HOST_PART =
  /^(?:word\/(?:document\d*|header[^/]*|footer[^/]*|footnotes\d*|endnotes\d*|comments\d*)\.xml|xl\/(?:worksheets|drawings)\/[^/]+\.xml|ppt\/(?:slides|slideLayouts|slideMasters|notesSlides)\/[^/]+\.xml)$/;

/**
 * Evidence lines for the AI panel, format-agnostic.
 *
 * Prefers the part the user actually has open — `rawXml` is that part's text — and falls
 * back to the first host part that contains any links. Picking the first *possible* host
 * blind would report "no hyperlinks" for a deck whose links are on slide 4, since bundles
 * carry layouts and masters alongside the slide and key order is insertion order.
 */
export function computeHyperlinkEvidenceForMarkup(
  parts: Record<string, string>,
  rawXml: string
): { lines: string[]; unresolved: string[] } | null {
  const hosts = Object.keys(parts).filter(path => HYPERLINK_HOST_PART.test(path));
  // The open part first, then anything else that has links.
  const ordered = [...hosts.filter(path => parts[path] === rawXml), ...hosts.filter(path => parts[path] !== rawXml)];

  let hostPath: string | null = null;
  let links: Hyperlink[] = [];
  for (const path of ordered) {
    const found = readHyperlinks(parts, path);
    if (found.length === 0) continue;
    hostPath = path;
    links = found;
    break;
  }
  if (hostPath === null) return null;

  const lines: string[] = [];
  const unresolved: string[] = [];
  const orphans = unreferencedRelationshipFindings(parts, hostPath);

  const external = links.filter(link => link.kind === 'external');
  const internal = links.filter(link => link.kind === 'internal');
  lines.push(
    `${hostPath} contains ${links.length} hyperlink(s): ${internal.length} pointing inside the document, ` +
      `${external.length} pointing outside it. A broken hyperlink renders identically to a working one — the ` +
      `styling lives on the runs, the destination lives in an attribute nothing draws.`
  );

  for (const link of links) {
    const destination =
      link.externalTarget !== null
        ? `the external target "${link.externalTarget}"`
        : link.anchor !== null && link.anchor !== ''
          ? `"${link.anchor}" inside this document`
          : link.internalTarget !== null
            ? link.action !== null
              ? `the package part "${link.internalTarget}" via the action "${link.action}"`
              : `the package part "${link.internalTarget}"`
            : link.action !== null
              ? `the action "${link.action}"`
              : 'nothing';
    lines.push(
      `A ${link.format} hyperlink (${link.label}${link.cellRange === null ? '' : ` over ${link.cellRange}`}) ` +
        `read from ${link.destinationEvidence}, pointing at ${destination}.`
    );

    if (hyperlinkResolves(link) === null && link.kind === 'external') {
      lines.push(
        'Whether that destination is reachable cannot be determined from the package: it is a URI outside the file, and nothing here was fetched to find out.'
      );
      unresolved.push(
        `The external target "${link.externalTarget}" cannot be verified from inside the package — no URL was fetched, so whether it resolves, redirects or 404s is unknown.`
      );
    }

    lines.push(...renderFindings(link.problems));
  }

  lines.push(...renderFindings(orphans));

  const silent = findSilentlyBrokenHyperlinks(links);
  if (silent.length > 0) {
    lines.push(
      `${silent.length} of these link(s) will render exactly as intended and are broken anyway — hyperlink styling is applied to the text, never to the destination, so no visual check will catch this before someone clicks.`
    );
  }

  if (links.some(link => link.format === 'excel' && link.anchor !== null) && parts['xl/workbook.xml'] === undefined) {
    unresolved.push(
      'xl/workbook.xml is not in this bundle, so Excel @location values could not be checked against the sheet names and defined names the workbook declares.'
    );
  }

  if (links.some(link => link.relationshipId !== null && link.anchor !== null && link.anchor !== '')) {
    unresolved.push(
      'Which destination a consumer follows when a hyperlink carries both a relationship id and an internal anchor was not verified against ECMA-376 or against observed Word behaviour.'
    );
  }

  return { lines, unresolved };
}
