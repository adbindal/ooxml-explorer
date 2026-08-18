/**
 * WordprocessingML bookmarks — the range model, and why they break silently.
 *
 * A bookmark is not a thing in the document; it is a *pair* of empty marker elements
 * that happen to sit around some content:
 *
 *   <w:bookmarkStart w:id="3" w:name="Chapter2"/> …content… <w:bookmarkEnd w:id="3"/>
 *
 * Three consequences follow, and each one is a class of bug.
 *
 * 1. THE PAIR IS MATCHED BY @w:id, NOT BY @w:name.
 *    Only the start carries a name; the end carries an id and nothing else. So an end
 *    with no matching start — or a start whose end was deleted when the surrounding
 *    content was edited — does not produce a malformed bookmark. It produces *no
 *    bookmark at all*, and every hyperlink, cross-reference and TOC entry aimed at that
 *    name resolves to nothing. The document still opens. Nothing looks wrong until
 *    someone clicks a link.
 *
 * 2. THE MARKERS ARE NOT INSIDE RUNS.
 *    They are direct children of w:p (or of w:tbl, w:tr, w:tc, w:body…). Any code that
 *    walks runs to find content will not see them, which is exactly the bug this
 *    project's own diff had: bookmarks were being stripped as noise alongside
 *    revision ids, so a document that lost every cross-reference target compared equal.
 *
 * 3. A BOOKMARK CAN SPAN ANYTHING.
 *    Start and end are independent, so a bookmark can begin mid-paragraph and end
 *    inside a different table cell. Nesting and overlap are both legal. "What text
 *    does this bookmark cover?" therefore cannot be answered by reading any single
 *    element — it needs a walk between two points in document order.
 *
 * THE CORRUPTION CLASS: @w:id IS SHARED WITH TRACKED CHANGES.
 *
 * `w:id` on bookmarkStart is not a bookmark-private counter. The same attribute names
 * revisions (w:ins, w:del), permissions (w:permStart), and every *Change element. A
 * generator that numbers its tracked changes from 1 — the obvious thing to do — will
 * collide with bookmark ids in any document that already has bookmarks, and documents
 * routinely have hundreds. Word then rejects the file as corrupt while lenient readers
 * (macOS Preview, most libraries) open it happily, so the bug reaches users having
 * passed every test.
 *
 * ⚠️ PROVENANCE: this is *observed Word behaviour* reported against a real generator
 * (anthropics/skills issue #489), not a rule stated in ECMA-376. The specification
 * assigns w:id to each element type without saying the space is shared, and Word's
 * rejection is stricter than the schema. Encoded here as behaviour, not as a citation.
 * The safe generator rule: scan for the maximum w:id across ALL element types and
 * start above it. See `findMarkupIdCollisions`.
 *
 * Verified against the Open XML SDK schema data (`CT_Bookmark`, `CT_MarkupRange`,
 * `CT_MoveBookmark`, `CT_Markup`): @w:name and @w:id are both *required* on
 * bookmarkStart, @w:name is capped at 40 characters, and @w:id accepts a non-negative
 * number or a number ≤ -2 — so -1 alone is out of range, a quirk of the union type.
 */

import { W_NAMESPACE } from './wordStyleResolver';
import { finding, renderFindings, type Finding, type Severity } from './findings';

/** `CT_Bookmark/@w:name` carries a StringValidator with MaxLength 40. */
export const MAX_BOOKMARK_NAME_LENGTH = 40;

/**
 * Word writes this to remember where the cursor was and rewrites it on every save.
 * The one bookmark that genuinely is noise — see `NOISE_BOOKMARK_NAMES` in ooxmlDiff.
 */
export const LAST_EDIT_BOOKMARK = '_GoBack';

/** Range markers that pair a start and an end by @w:id, keyed start → end. */
const RANGE_KINDS = {
  bookmarkStart: 'bookmarkEnd',
  moveFromRangeStart: 'moveFromRangeEnd',
  moveToRangeStart: 'moveToRangeEnd'
} as const;

export type BookmarkKind = keyof typeof RANGE_KINDS;

/**
 * Every element type that draws its @w:id from the shared markup id space.
 * Derived from the SDK schema by selecting types declaring a `w:id` attribute; the
 * *Change and math-revision types inherit theirs through CT_Markup / CT_TrackChange.
 */
const MARKUP_ID_ELEMENTS = new Set([
  'bookmarkStart', 'bookmarkEnd',
  'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd',
  'ins', 'del', 'moveFrom', 'moveTo',
  'permStart', 'permEnd',
  'cellMerge', 'numberingChange',
  'pPrChange', 'rPrChange', 'sectPrChange',
  'tblPrChange', 'tblPrExChange', 'tblGridChange', 'tcPrChange', 'trPrChange'
]);

/** Prefixes Word uses for bookmarks it generates itself and hides from the UI. */
const GENERATED_PREFIXES: Array<[string, string]> = [
  ['_Toc', 'a table-of-contents entry'],
  ['_Ref', 'a cross-reference target'],
  ['_Hlk', 'an editing-session marker Word leaves behind']
];

/**
 * Severity and silence for each kind, decided once here rather than at each call site.
 *
 * Almost everything a bookmark gets wrong is SILENT: the document opens, the text is
 * all present, and the page looks exactly right. What breaks is navigation — a
 * hyperlink, a cross-reference, a TOC entry — which nobody notices until they click.
 * `name-too-long` is the exception that Word itself will refuse.
 */
const BOOKMARK_RULES = {
  'unmatched-start':  { severity: 'error',   silent: true },
  'unmatched-end':    { severity: 'warning', silent: true },
  'duplicate-id':     { severity: 'error',   silent: true },
  'duplicate-name':   { severity: 'warning', silent: true },
  'missing-name':     { severity: 'error',   silent: true },
  'name-too-long':    { severity: 'error',   silent: false },
  'id-out-of-range':  { severity: 'error',   silent: false },
  'reversed-range':   { severity: 'warning', silent: true },
  'id-collision':     { severity: 'error',   silent: true }
} as const satisfies Record<string, { severity: Severity; silent: boolean }>;

export type BookmarkProblemKind = keyof typeof BOOKMARK_RULES;

/** Builds a bookmark finding, applying the severity table above. */
const bookmarkFinding = (
  kind: BookmarkProblemKind,
  part: string,
  message: string,
  remediation: string,
  subject?: Record<string, string>
): Finding =>
  finding(`bookmark/${kind}`, part, message, remediation, { ...BOOKMARK_RULES[kind], subject });

export interface Bookmark {
  id: string;
  name: string;
  kind: BookmarkKind;
  start: Element;
  /** null when no end carries this id — the bookmark does not exist for Word. */
  end: Element | null;
  /** @w:colFirst / @w:colLast: present only when the bookmark covers table columns. */
  colFirst: string | null;
  colLast: string | null;
  isColumnRange: boolean;
  /** Names beginning with `_` are hidden from Word's bookmark dialog. */
  hidden: boolean;
  /** Set when the name matches a prefix Word generates, explaining what made it. */
  generatedBy: string | null;
}

export interface BookmarkIndex {
  bookmarks: Bookmark[];
  /** Names are not unique in practice, so this maps to a list. */
  byName: Map<string, Bookmark[]>;
  byId: Map<string, Bookmark>;
  problems: Finding[];
}

export interface IdCollision {
  id: string;
  /** Local names of the elements sharing this id, in document order. */
  elements: string[];
}

const attr = (el: Element, name: string) => el.getAttributeNS(W_NAMESPACE, name);

/** All descendants of `root` in the w namespace whose local name is in `names`. */
const collect = (root: ParentNode, names: Set<string> | string[]): Element[] => {
  const wanted = names instanceof Set ? names : new Set(names);
  return Array.from(root.querySelectorAll('*')).filter(
    el => el.namespaceURI === W_NAMESPACE && wanted.has(el.localName)
  );
};

/**
 * `w:id` is `ST_NonNegativeDecimalNumber` unioned with signed numbers ≤ -2.
 * -1 is therefore the one integer the union excludes.
 */
const idIsInRange = (raw: string): boolean => {
  if (!/^-?\d+$/.test(raw)) return false;
  const n = Number(raw);
  return n >= 0 || n <= -2;
};

const classifyName = (name: string) => {
  for (const [prefix, description] of GENERATED_PREFIXES) {
    if (name.startsWith(prefix)) return description;
  }
  if (name === LAST_EDIT_BOOKMARK) {
    return "Word's record of the last edit position; it moves on every save";
  }
  return null;
};

/**
 * Index every bookmark-like range in a document part, pairing starts to ends by id
 * and reporting what could not be paired rather than throwing.
 *
 * Pass the parsed `word/document.xml` — or any body part. Headers, footers, footnotes
 * and endnotes each carry their own bookmarks in their own id space, so index them
 * separately; ids are unique *within a part*, not across the package.
 */
export function readBookmarks(doc: Document | Element, part = ''): BookmarkIndex {
  const root: ParentNode = 'documentElement' in doc && doc.documentElement ? doc.documentElement : (doc as Element);
  const problems: Finding[] = [];
  const bookmarks: Bookmark[] = [];
  const byName = new Map<string, Bookmark[]>();
  const byId = new Map<string, Bookmark>();

  const startNames = Object.keys(RANGE_KINDS);
  const endNames = Object.values(RANGE_KINDS);
  const starts = collect(root, startNames);
  const ends = collect(root, endNames);

  // Ends are matched by id *within their own kind* — a moveToRangeEnd does not close a
  // bookmarkStart even if the ids happen to agree.
  const endsByKind = new Map<string, Map<string, Element[]>>();
  for (const end of ends) {
    const id = attr(end, 'id');
    if (id === null) continue;
    const forKind = endsByKind.get(end.localName) ?? new Map<string, Element[]>();
    forKind.set(id, [...(forKind.get(id) ?? []), end]);
    endsByKind.set(end.localName, forKind);
  }

  const claimedEnds = new Set<Element>();

  for (const start of starts) {
    const kind = start.localName as BookmarkKind;
    const id = attr(start, 'id');
    const name = attr(start, 'name');

    if (id === null) {
      problems.push(bookmarkFinding(
        'unmatched-start', part,
        `A <w:${kind}> has no w:id, so nothing can close it. @w:id is required.`,
        'Give the start an id and add a matching end element.',
        name === null ? undefined : { name }
      ));
      continue;
    }

    if (!idIsInRange(id)) {
      problems.push(bookmarkFinding(
        'id-out-of-range', part,
        `w:id "${id}" is outside the permitted range. The type admits any non-negative number, or a negative number of -2 or below; -1 and non-integers are excluded.`,
        'Renumber this marker to a non-negative integer above every other w:id in the part.',
        name === null ? { id } : { id, name }
      ));
    }

    if (name === null) {
      problems.push(bookmarkFinding(
        'missing-name', part,
        `<w:${kind} w:id="${id}"> has no w:name. The name is required, and it is the only handle a hyperlink or cross-reference has.`,
        'Add a w:name, or delete the marker pair if nothing references it.',
        { id }
      ));
    } else if (name.length > MAX_BOOKMARK_NAME_LENGTH) {
      problems.push(bookmarkFinding(
        'name-too-long', part,
        `Bookmark name is ${name.length} characters; the maximum is ${MAX_BOOKMARK_NAME_LENGTH}.`,
        `Shorten the name to ${MAX_BOOKMARK_NAME_LENGTH} characters or fewer, and update every reference to it.`,
        { id, name }
      ));
    }

    const candidates = endsByKind.get(RANGE_KINDS[kind])?.get(id) ?? [];
    const end = candidates.find(candidate => !claimedEnds.has(candidate)) ?? null;
    if (end) claimedEnds.add(end);

    const resolvedName = name ?? '';
    const colFirst = attr(start, 'colFirst');
    const colLast = attr(start, 'colLast');

    const bookmark: Bookmark = {
      id,
      name: resolvedName,
      kind,
      start,
      end,
      colFirst,
      colLast,
      isColumnRange: colFirst !== null || colLast !== null,
      hidden: resolvedName.startsWith('_'),
      generatedBy: classifyName(resolvedName)
    };

    if (!end) {
      problems.push(bookmarkFinding(
        'unmatched-start', part,
        `Bookmark "${resolvedName}" opens but never closes — no <w:${RANGE_KINDS[kind]} w:id="${id}"/> exists. Word treats the bookmark as absent, so hyperlinks, cross-references and TOC entries aimed at this name resolve to nothing. The document still opens and looks correct.`,
        `Insert <w:${RANGE_KINDS[kind]} w:id="${id}"/> at the point the range should end.`,
        { id, name: resolvedName }
      ));
    } else if (start.compareDocumentPosition(end) & Node.DOCUMENT_POSITION_PRECEDING) {
      problems.push(bookmarkFinding(
        'reversed-range', part,
        `Bookmark "${resolvedName}" has its end before its start in document order, so it covers no content.`,
        'Move the end marker after the start, or swap the two.',
        { id, name: resolvedName }
      ));
    }

    if (byId.has(id)) {
      const first = byId.get(id)!;
      problems.push(bookmarkFinding(
        'duplicate-id', part,
        `w:id "${id}" starts two ranges ("${first.name}" and "${resolvedName}"). Ends match by id, so the second start competes with the first for the same end marker and one of the two bookmarks silently loses its range.`,
        'Renumber one of them to an id unused anywhere in the part.',
        { id, name: resolvedName }
      ));
    } else {
      byId.set(id, bookmark);
    }

    const sameName = byName.get(resolvedName);
    if (sameName && resolvedName !== '') {
      problems.push(bookmarkFinding(
        'duplicate-name', part,
        `Two bookmarks are named "${resolvedName}". Names are the reference handle, so anything pointing at this name reaches only one of them — and which one is not something the markup determines.`,
        'Rename one, or delete the redundant range.',
        { id, name: resolvedName }
      ));
      sameName.push(bookmark);
    } else {
      byName.set(resolvedName, [bookmark]);
    }

    bookmarks.push(bookmark);
  }

  for (const end of ends) {
    if (claimedEnds.has(end)) continue;
    const id = attr(end, 'id');
    problems.push(bookmarkFinding(
      'unmatched-end', part,
      `A <w:${end.localName}${id === null ? '' : ` w:id="${id}"`}/> closes a range that was never opened. Because starts carry the name and ends do not, there is no way to tell which bookmark was lost — only that one was.`,
      'Delete the stray end, or restore the start that used to match it.',
      id === null ? undefined : { id }
    ));
  }

  return { bookmarks, byName, byId, problems };
}

/**
 * The text a bookmark covers, which is the question people actually ask and the one
 * no single element answers.
 *
 * Returns null when the range never closed — deliberately distinct from `''`, which is
 * a real answer meaning the bookmark is present but covers nothing (an insertion point,
 * which is what a cross-reference target usually is).
 *
 * Deleted text (`w:delText`) is excluded because it is not part of the document as
 * read, and field instructions (`w:instrText`) because they are code, not content.
 */
export function bookmarkText(bookmark: Bookmark): string | null {
  if (!bookmark.end) return null;
  const root = bookmark.start.ownerDocument?.documentElement;
  if (!root) return null;

  const parts: string[] = [];
  for (const t of collect(root, ['t'])) {
    const afterStart = bookmark.start.compareDocumentPosition(t) & Node.DOCUMENT_POSITION_FOLLOWING;
    const beforeEnd = bookmark.end.compareDocumentPosition(t) & Node.DOCUMENT_POSITION_PRECEDING;
    if (afterStart && beforeEnd) parts.push(t.textContent ?? '');
  }
  return parts.join('');
}

/**
 * Every @w:id used by more than one element in the shared markup id space.
 *
 * This is the check that catches the corruption class described at the top of the file:
 * a generator numbering its tracked changes from 1 collides with existing bookmark ids,
 * Word refuses to open the result, and every lenient reader says the file is fine.
 *
 * Collisions *within* one element type are also reported — two bookmarkStarts sharing
 * an id break range matching just as surely.
 */
export function findMarkupIdCollisions(doc: Document | Element): IdCollision[] {
  const root: ParentNode = 'documentElement' in doc && doc.documentElement ? doc.documentElement : (doc as Element);
  const seen = new Map<string, string[]>();

  for (const el of collect(root, MARKUP_ID_ELEMENTS)) {
    const id = attr(el, 'id');
    if (id === null) continue;
    seen.set(id, [...(seen.get(id) ?? []), el.localName]);
  }

  const collisions: IdCollision[] = [];
  for (const [id, elements] of seen) {
    if (elements.length > 1) collisions.push({ id, elements });
  }
  return collisions;
}

/**
 * The id a generator should start from so that nothing it writes can collide.
 * One past the highest id in the part, across every element type — never one past the
 * highest *bookmark* id, which is the mistake that produces the corruption.
 */
export function nextSafeMarkupId(doc: Document | Element): number {
  const root: ParentNode = 'documentElement' in doc && doc.documentElement ? doc.documentElement : (doc as Element);
  let max = -1;
  for (const el of collect(root, MARKUP_ID_ELEMENTS)) {
    const raw = attr(el, 'id');
    if (raw === null || !/^-?\d+$/.test(raw)) continue;
    max = Math.max(max, Number(raw));
  }
  return max + 1;
}

/** Word body parts carry bookmarks; each is its own id space. */
const WORD_BODY_PART = /^word\/(?:document\d*|header[^/]*|footer[^/]*|footnotes\d*|endnotes\d*|comments\d*)\.xml$/;

/**
 * Evidence lines for the AI panel: what the bookmarks in this part actually are, and
 * what is broken about them.
 *
 * Everything here is computed, never asserted by the model — the same contract the
 * formatting analyses hold to. Facts the markup does not settle go to `unresolved`
 * so the tier is capped below Verified rather than papered over.
 */
export function computeBookmarkEvidenceForMarkup(
  parts: Record<string, string>,
  rawXml: string
): { lines: string[]; unresolved: string[] } | null {
  const entry = Object.entries(parts).find(([path]) => WORD_BODY_PART.test(path));
  if (!entry) return null;

  const [path, xml] = entry;
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) return null;

  const index = readBookmarks(doc, path);
  const collisions = findMarkupIdCollisions(doc);
  if (index.bookmarks.length === 0 && index.problems.length === 0 && collisions.length === 0) return null;

  const lines: string[] = [];
  const unresolved: string[] = [];

  const visible = index.bookmarks.filter(b => !b.hidden);
  lines.push(
    `${path} declares ${index.bookmarks.length} bookmark range(s): ${visible.length} user-visible, ` +
      `${index.bookmarks.length - visible.length} hidden (names beginning with "_", which Word omits from its bookmark dialog).`
  );

  // Name the selected bookmark when the user has one open, since that is what the
  // question is almost certainly about.
  const selectedName = /<w:bookmark(?:Start|End)\b[^>]*w:name="([^"]*)"/.exec(rawXml)?.[1];
  const selected = selectedName ? index.byName.get(selectedName)?.[0] : undefined;
  if (selected) {
    const covered = bookmarkText(selected);
    lines.push(
      covered === null
        ? `Bookmark "${selected.name}" (w:id="${selected.id}") never closes, so it covers nothing and does not exist as far as Word is concerned.`
        : covered === ''
          ? `Bookmark "${selected.name}" (w:id="${selected.id}") covers no text — an insertion point, which is what a cross-reference target normally is.`
          : `Bookmark "${selected.name}" (w:id="${selected.id}") covers: "${covered.slice(0, 200)}${covered.length > 200 ? '…' : ''}".`
    );
    if (selected.generatedBy) {
      lines.push(`"${selected.name}" is generated by Word, not authored: it marks ${selected.generatedBy}.`);
    }
    const references = findBookmarkReferences(doc, selected.name);
    lines.push(
      references.length === 0
        ? `Nothing in ${path} references "${selected.name}" by name. That is not itself an error — many bookmarks are navigation aids — but a cross-reference that used to point here would look exactly like this.`
        : `${references.length} reference(s) to "${selected.name}" found in this part (hyperlink anchors and field instructions).`
    );
    unresolved.push(
      `References from other parts of the package to "${selected.name}" were not checked; only ${path} was read.`
    );
  }

  lines.push(...renderFindings(index.problems));

  for (const collision of collisions) {
    lines.push(
      `w:id "${collision.id}" is used by ${collision.elements.length} elements (${collision.elements.join(', ')}). ` +
        `w:id is drawn from one space shared by bookmarks, tracked changes and permissions, and Word rejects files that reuse an id — ` +
        `while lenient readers open them without complaint. Renumber from ${nextSafeMarkupId(doc)} upward.`
    );
  }

  return { lines, unresolved };
}

/**
 * Does anything in the part reference this bookmark by name?
 *
 * Checks the three references that matter: `w:hyperlink/@w:anchor`, and field
 * instructions (`w:instrText`, `w:fldSimple/@w:instr`) which is how REF, PAGEREF and
 * TOC entries name their target. An unreferenced bookmark is not an error — plenty are
 * navigation aids — but an unreferenced bookmark that something *used* to point at is
 * how a broken cross-reference looks after the fact.
 */
export function findBookmarkReferences(doc: Document | Element, name: string): Element[] {
  const root: ParentNode = 'documentElement' in doc && doc.documentElement ? doc.documentElement : (doc as Element);
  const found: Element[] = [];

  for (const link of collect(root, ['hyperlink'])) {
    if (attr(link, 'anchor') === name) found.push(link);
  }
  // Field instructions name their target as a bare token, so match on word boundaries
  // rather than substring: bookmark "Ch1" must not match a reference to "Ch10".
  const token = new RegExp(`(^|[\\s"])${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}($|[\\s"\\\\])`);
  for (const field of collect(root, ['instrText'])) {
    if (token.test(field.textContent ?? '')) found.push(field);
  }
  for (const field of collect(root, ['fldSimple'])) {
    if (token.test(attr(field, 'instr') ?? '')) found.push(field);
  }
  return found;
}
