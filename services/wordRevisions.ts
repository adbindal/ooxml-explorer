/**
 * WordprocessingML tracked changes — the document with unaccepted revisions says TWO
 * DIFFERENT THINGS, and nothing in the file says which one you are reading.
 *
 * A `.docx` with revisions in it holds both the text before the edits and the text after
 * them, interleaved:
 *
 *   <w:ins w:id="4" w:author="Ann" w:date="…"><w:r><w:t>new wording</w:t></w:r></w:ins>
 *   <w:del w:id="5" w:author="Ann" w:date="…"><w:r><w:delText>old wording</w:delText></w:r></w:del>
 *
 * Accept everything and it reads "new wording". Reject everything and it reads "old
 * wording". Both readings are in the file, and every tool that extracts text silently
 * picks one — usually without saying so, and usually not the one the reader assumes.
 *
 * THREE CONSEQUENCES, EACH A CLASS OF BUG.
 *
 * 1. DELETED TEXT IS STILL IN THE FILE.
 *    `w:del` does not remove anything; it wraps runs whose text moved from `w:t` into
 *    `w:delText`. The words are still there, readable by anyone who unzips the package.
 *    Deleting a paragraph and sending the file is not redaction. Worse in the other
 *    direction: an extractor that walks text nodes without checking their ancestors
 *    emits `w:delText` as though it were live content, so the output contains sentences
 *    the document does not say.
 *
 * 2. PARAGRAPH-MARK REVISIONS CHANGE THE STRUCTURE, NOT THE WORDS.
 *    `w:pPr/w:rPr/w:del` marks the *paragraph mark itself* as deleted. Accepting it
 *    merges this paragraph with the next one; `w:pPr/w:rPr/w:ins` rejected does the same.
 *    No run is touched, so anything that only looks inside runs sees nothing at all —
 *    and then two paragraphs become one, which moves every heading, list number and
 *    cross-reference after it.
 *
 * 3. FORMAT-ONLY REVISIONS ARE INVISIBLE IN THE TEXT ENTIRELY.
 *    `w:pPrChange`, `w:rPrChange`, `w:tblPrChange`, `w:sectPrChange` and friends record
 *    the **previous** formatting, not the new formatting. They never affect a single
 *    character, so no text comparison of any kind can find them, and accepting or
 *    rejecting them changes how the document looks with no change to what it says.
 *
 * MOVES ARE A DELETE AND AN INSERT THAT KNOW ABOUT EACH OTHER.
 *
 * `w:moveFrom`/`w:moveTo` wrap the two halves of moved content, and the surrounding
 * `w:moveFromRangeStart`/`w:moveToRangeStart` markers carry a **`w:name`** that ties the
 * two halves together. The range start and its end pair by `@w:id` like any bookmark;
 * the two *halves* pair with each other by name. Break the naming and the move degrades
 * into an unrelated deletion plus an unrelated insertion — which still accepts to the
 * same text, but reviews as two changes instead of one and loses the fact that nothing
 * was written or removed.
 *
 * THE SHARED ID SPACE. `@w:id` on `w:ins`/`w:del` is drawn from the same space as
 * bookmarks and permissions. A generator numbering its revisions from 1 collides with
 * bookmark ids in any document that has bookmarks, and Word rejects the result as
 * corrupt while lenient readers open it happily. That check already exists —
 * `findMarkupIdCollisions` in `wordBookmarks.ts` — and is called from here rather than
 * reimplemented; this module only reports the collisions a revision took part in.
 *
 * ─────────────────────────────────────────────────────────────────────────────────────
 * VERIFIED against the Open XML SDK schema data
 * (`data/schemas/schemas_openxmlformats_org_wordprocessingml_2006_main.json`, types
 * `CT_TrackChange`, `CT_RunTrackChange`, `CT_Markup`, `CT_MarkupRange`, `CT_MoveBookmark`,
 * `CT_PPrChange`, `CT_RPrChange`, `CT_ParaRPrChange`, `CT_TblGridChange`), cross-checked
 * against this repo's own schema-derived corpus in `public/rag-data.json`:
 *
 *   • `@w:author` is REQUIRED on `CT_TrackChange` and `CT_RunTrackChange` and on every
 *     *Change type, with a 255-character maximum.
 *   • ⚠️ `@w:date` is **NOT required** — the brief for this module said it was, and the
 *     schema disagrees. It carries no `RequiredValidator` on `CT_TrackChange`,
 *     `CT_RunTrackChange`, `CT_PPrChange`, `CT_RPrChange`, `CT_ParaRPrChange`,
 *     `CT_SectPrChange`, `CT_Tbl*Change`, `CT_TcPrChange` or `CT_TrPrChange`. It IS
 *     required on `CT_MoveBookmark` (`w:moveFromRangeStart`, `w:moveToRangeStart`),
 *     which is the one place a missing date is a schema violation. Both cases are
 *     reported, at different severities, because they are different facts.
 *   • ⚠️ `w:tblGridChange` (`CT_TblGridChange`) declares **only `@w:id`** — no author and
 *     no date. It is the one *Change element that carries no attribution, so it is
 *     excluded from the author check rather than reported as missing one.
 *   • `@w:id` is required everywhere and takes `ST_NonNegativeDecimalNumber` unioned with
 *     signed numbers ≤ -2, so -1 is the single integer the type excludes.
 *   • `@w:name` is required on `CT_MoveBookmark`; the matching `w:moveFromRangeEnd` /
 *     `w:moveToRangeEnd` are `CT_MarkupRange` and carry **no name** — only `@w:id`.
 *   • `w:pPr/w:rPr` is `CT_ParaRPr`, whose children include `w:ins`, `w:del`,
 *     `w:moveFrom`, `w:moveTo`. Run properties (`CT_RPr`) do **not** admit them, so a
 *     revision element inside an `w:rPr` is always a paragraph-mark revision.
 *   • `w:ins`/`w:del` also appear inside `w:trPr` (row inserted/deleted) and `w:ins`
 *     inside `w:numPr`; `w:cellIns`, `w:cellDel`, `w:cellMerge` inside `w:tcPr`.
 *   • `CT_R` admits `w:t`, `w:delText` and `w:delInstrText` — all three are `CT_Text`.
 *   • `word/settings.xml` carries `w:trackRevisions`, `w:doNotTrackMoves`,
 *     `w:doNotTrackFormatting` (all `CT_OnOff`) and `w:revisionView`
 *     (`CT_TrackChangesView`) with `@w:markup`, `@w:comments`, `@w:insDel`
 *     ("Display Content Revisions"), `@w:formatting`, `@w:inkAnnotations`.
 *
 * ⚠️ NOT VERIFIED, and treated accordingly:
 *   • That deleted runs *must* use `w:delText` rather than `w:t` is a semantic rule, not
 *     a schema constraint — `CT_R` permits both regardless of what encloses the run. A
 *     `w:t` inside a `w:del` is therefore reported as a hazard, not as invalid markup.
 *   • What Word displays for a `w:delText` that sits outside any deletion was not
 *     confirmed against Word. The finding says the markup is inconsistent and stops
 *     short of predicting the rendering.
 *   • The exact merge behaviour of two adjacent paragraph-mark revisions by different
 *     authors was not confirmed; the accept/reject model below applies each mark
 *     independently, which is the obvious reading and is documented as an assumption.
 */

import { W_NAMESPACE } from './wordStyleResolver';
import { readBookmarks, findMarkupIdCollisions, nextSafeMarkupId, type Bookmark } from './wordBookmarks';
import { finding, renderFindings, type Finding, type Severity } from './findings';

/**
 * Severity and silence per kind.
 *
 * EVERY finding here is silent, and that is not laziness — it is the defining property
 * of tracked changes. A document full of unaccepted revisions opens, renders and prints
 * exactly as intended; Word shows the markup and everything is fine. What breaks is
 * everything *downstream*: extraction, conversion, diffing, search indexing, redaction.
 * There is no screenshot and no human read-through that catches any of it.
 *
 * `id-out-of-range` is the sole visible one, because Word refuses to open a file whose
 * `@w:id` is outside the union type — that failure is loud and immediate.
 *
 * The three `note` kinds describe states that are entirely legitimate: a document under
 * review is *supposed* to have revisions in it. They are reported because a consumer
 * reading the text needs to know the text is provisional, not because anything is wrong.
 */
const REVISION_RULES = {
  'unaccepted-revisions':     { severity: 'note',    silent: true },
  'paragraph-mark-revision':  { severity: 'note',    silent: true },
  'format-only-revision':     { severity: 'note',    silent: true },
  'deleted-text-retained':    { severity: 'warning', silent: true },
  'live-text-in-deletion':    { severity: 'error',   silent: true },
  'orphan-deleted-text':      { severity: 'error',   silent: true },
  'missing-author':           { severity: 'error',   silent: true },
  'author-too-long':          { severity: 'warning', silent: true },
  'missing-date':             { severity: 'note',    silent: true },
  'missing-required-date':    { severity: 'error',   silent: true },
  'missing-id':               { severity: 'error',   silent: true },
  'id-out-of-range':          { severity: 'error',   silent: false },
  'id-collision':             { severity: 'error',   silent: true },
  'unpaired-move':            { severity: 'error',   silent: true },
  'duplicate-move-name':      { severity: 'warning', silent: true },
  'unnamed-move-range':       { severity: 'error',   silent: true },
  'revisions-hidden-in-view': { severity: 'warning', silent: true }
} as const satisfies Record<string, { severity: Severity; silent: boolean }>;

export type RevisionProblemKind = keyof typeof REVISION_RULES;

const revisionFinding = (
  kind: RevisionProblemKind,
  part: string,
  message: string,
  remediation: string,
  subject?: Record<string, string>
): Finding => finding(`revision/${kind}`, part, message, remediation, { ...REVISION_RULES[kind], subject });

/** Author string used when `@w:author` is absent, so per-author counts still add up. */
export const UNATTRIBUTED_AUTHOR = '(no w:author)';

/** `CT_TrackChange/@w:author` carries a StringValidator with MaxLength 255. */
export const MAX_AUTHOR_LENGTH = 255;

/** The four elements that wrap content, whatever they enclose. */
const CONTENT_REVISIONS = new Set(['ins', 'del', 'moveFrom', 'moveTo']);

/** Accepting the change removes the content these enclose. */
const REMOVED_BY_ACCEPT = new Set(['del', 'moveFrom']);

/** Rejecting the change removes the content these enclose. */
const REMOVED_BY_REJECT = new Set(['ins', 'moveTo']);

/**
 * Revisions that record the PREVIOUS formatting and touch no text.
 * `w:numberingChange` is grouped separately below because it is numbering, not format.
 */
const FORMAT_CHANGES = new Set([
  'pPrChange', 'rPrChange', 'sectPrChange',
  'tblPrChange', 'tblPrExChange', 'tblGridChange', 'tcPrChange', 'trPrChange'
]);

/** Table-cell revisions, all children of `w:tcPr`. */
const CELL_REVISIONS = new Set(['cellIns', 'cellDel', 'cellMerge']);

/**
 * The one revision element with no `@w:author` and no `@w:date` in the schema.
 * Checking it for attribution would report a fault on every valid table edit.
 */
const UNATTRIBUTED_ELEMENTS = new Set(['tblGridChange']);

/** Range markers naming the two halves of a move. `CT_MoveBookmark`. */
const MOVE_RANGE_STARTS = { moveFromRangeStart: 'moveFrom', moveToRangeStart: 'moveTo' } as const;

export type RevisionCategory =
  | 'content'
  | 'move'
  | 'paragraph-mark'
  | 'row'
  | 'cell'
  | 'numbering'
  | 'format';

export interface Revision {
  /** Local name of the element, e.g. `ins`, `pPrChange`, `cellMerge`. */
  kind: string;
  category: RevisionCategory;
  element: Element;
  id: string | null;
  author: string | null;
  /** ISO-ish timestamp as written. Absent on most types without violating the schema. */
  date: string | null;
  /** Every `w:t` and `w:delText` this revision encloses; `''` for the ones with no text. */
  text: string;
  /** Name of the enclosing move range, when this half sits inside one. */
  moveName: string | null;
}

export interface Move {
  name: string;
  /** The `w:moveFromRangeStart` range, as indexed by the bookmark reader. */
  from: Bookmark | null;
  to: Bookmark | null;
  /** True only when both halves exist — otherwise this is a delete and an insert. */
  paired: boolean;
}

export interface RevisionIndex {
  revisions: Revision[];
  /** Revision count per author, including `UNATTRIBUTED_AUTHOR`. */
  byAuthor: Map<string, number>;
  moves: Move[];
  problems: Finding[];
}

const isW = (el: Element, local: string) => el.namespaceURI === W_NAMESPACE && el.localName === local;
const attr = (el: Element, name: string) => el.getAttributeNS(W_NAMESPACE, name);

const rootOf = (node: Document | Element): ParentNode =>
  'documentElement' in node && node.documentElement ? node.documentElement : (node as Element);

const allElements = (root: ParentNode): Element[] =>
  Array.from(root.querySelectorAll('*')).filter(el => el.namespaceURI === W_NAMESPACE);

const childW = (parent: Element, local: string): Element | null =>
  Array.from(parent.children).find(child => isW(child, local)) ?? null;

/** `CT_OnOff`: present means on unless it says otherwise. */
const onOff = (el: Element | null): boolean => {
  if (!el) return false;
  const value = attr(el, 'val');
  return value === null || value === '1' || value === 'true' || value === 'on';
};

/**
 * `w:id` is `ST_NonNegativeDecimalNumber` unioned with signed numbers ≤ -2, so -1 is the
 * single integer the union excludes. Same rule as bookmarks — same attribute, same type.
 */
const idIsInRange = (raw: string): boolean => {
  if (!/^-?\d+$/.test(raw)) return false;
  const n = Number(raw);
  return n >= 0 || n <= -2;
};

const nearest = (el: Element, local: string): Element | null => {
  let node = el.parentElement;
  while (node) {
    if (isW(node, local)) return node;
    node = node.parentElement;
  }
  return null;
};

/**
 * The content revisions enclosing a node, innermost first.
 *
 * Revisions nest: text inserted by one author and later deleted by another sits inside a
 * `w:del` inside a `w:ins`. Both matter, and which is innermost decides whether the run
 * should be spelling its text as `w:t` or `w:delText`.
 *
 * Paragraph-mark revisions (inside `w:rPr`) are excluded: they mark the paragraph mark,
 * not the runs, and no text is ever their descendant anyway.
 */
const enclosingRevisions = (el: Element): string[] => {
  const kinds: string[] = [];
  let node = el.parentElement;
  while (node) {
    if (
      node.namespaceURI === W_NAMESPACE &&
      CONTENT_REVISIONS.has(node.localName) &&
      !(node.parentElement && isW(node.parentElement, 'rPr'))
    ) {
      kinds.push(node.localName);
    }
    node = node.parentElement;
  }
  return kinds;
};

/**
 * The revision applied to a paragraph's own mark, or null.
 * `w:p/w:pPr/w:rPr/{w:ins|w:del|w:moveFrom|w:moveTo}` — direct chain, never a descendant
 * search, or a revision on a run inside the paragraph would be mistaken for one.
 */
export const paragraphMarkRevision = (paragraph: Element): string | null => {
  const pPr = childW(paragraph, 'pPr');
  const rPr = pPr ? childW(pPr, 'rPr') : null;
  if (!rPr) return null;
  const marker = Array.from(rPr.children).find(
    child => child.namespaceURI === W_NAMESPACE && CONTENT_REVISIONS.has(child.localName)
  );
  return marker ? marker.localName : null;
};

const categoryOf = (el: Element): RevisionCategory | null => {
  const parent = el.parentElement;
  const parentName = parent && parent.namespaceURI === W_NAMESPACE ? parent.localName : '';
  const local = el.localName;

  if (CONTENT_REVISIONS.has(local)) {
    // Schema: only CT_ParaRPr admits these inside an w:rPr, so this is the paragraph mark.
    if (parentName === 'rPr') return 'paragraph-mark';
    if (parentName === 'trPr') return 'row';
    if (parentName === 'numPr') return 'numbering';
    return local === 'moveFrom' || local === 'moveTo' ? 'move' : 'content';
  }
  if (CELL_REVISIONS.has(local)) return 'cell';
  if (local === 'numberingChange') return 'numbering';
  if (FORMAT_CHANGES.has(local)) return 'format';
  return null;
};

/** Every `w:t` and `w:delText` a revision encloses, in document order. */
const revisionText = (el: Element): string =>
  Array.from(el.querySelectorAll('*'))
    .filter(child => isW(child, 't') || isW(child, 'delText'))
    .map(child => child.textContent ?? '')
    .join('');

/**
 * Index every tracked change in a body part, and everything wrong with them.
 *
 * Pass the parsed `word/document.xml`, or a header, footer, footnotes or endnotes part.
 * Each story carries its own revisions in its own `@w:id` space, so index them
 * separately — and a move whose halves land in two different parts cannot be paired from
 * either one, which is reported as unpaired rather than guessed at.
 */
export function readRevisions(doc: Document | Element, partPath = ''): RevisionIndex {
  const root = rootOf(doc);
  const elements = allElements(root);
  const problems: Finding[] = [];
  const revisions: Revision[] = [];

  // --- move ranges, so a move half can be named -------------------------------
  // readBookmarks already pairs moveFromRangeStart/End and moveToRangeStart/End by id
  // and reports the unpaired ones; its problems belong to the bookmark analyzer, so only
  // the ranges themselves are taken here.
  const ranges = readBookmarks(doc, partPath).bookmarks.filter(
    b => b.kind === 'moveFromRangeStart' || b.kind === 'moveToRangeStart'
  );

  const nameOfRangeContaining = (el: Element, half: 'moveFrom' | 'moveTo'): string | null => {
    for (const range of ranges) {
      if (MOVE_RANGE_STARTS[range.kind as keyof typeof MOVE_RANGE_STARTS] !== half) continue;
      if (!range.end) continue;
      const afterStart = range.start.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING;
      const beforeEnd = range.end.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_PRECEDING;
      if (afterStart && beforeEnd) return range.name;
    }
    return null;
  };

  // --- the revisions themselves ----------------------------------------------
  const byAuthor = new Map<string, number>();

  for (const el of elements) {
    const category = categoryOf(el);
    if (category === null) continue;

    const id = attr(el, 'id');
    const author = attr(el, 'author');
    const date = attr(el, 'date');
    const local = el.localName;

    const revision: Revision = {
      kind: local,
      category,
      element: el,
      id,
      author,
      date,
      text: revisionText(el),
      moveName:
        local === 'moveFrom' || local === 'moveTo'
          ? nameOfRangeContaining(el, local)
          : null
    };
    revisions.push(revision);

    const key = author ?? UNATTRIBUTED_AUTHOR;
    byAuthor.set(key, (byAuthor.get(key) ?? 0) + 1);

    if (id === null) {
      problems.push(revisionFinding(
        'missing-id', partPath,
        `<w:${local}> has no w:id. The attribute is required, and it is what ties a revision into the shared markup id space, so a revision without one cannot be renumbered or referred to.`,
        `Give the <w:${local}> a w:id above every other w:id in the part.`,
        author === null ? undefined : { author }
      ));
    } else if (!idIsInRange(id)) {
      problems.push(revisionFinding(
        'id-out-of-range', partPath,
        `<w:${local} w:id="${id}"> is outside the permitted range. The type admits any non-negative number, or a negative number of -2 or below; -1 and non-integers are excluded.`,
        'Renumber this revision to a non-negative integer above every other w:id in the part.',
        { id, element: local }
      ));
    }

    if (author === null && !UNATTRIBUTED_ELEMENTS.has(local)) {
      problems.push(revisionFinding(
        'missing-author', partPath,
        `<w:${local}${id === null ? '' : ` w:id="${id}"`}> has no w:author. The attribute is required by the schema, and without it the change cannot be attributed to anyone — a review that shows "who changed this" has nothing to show.`,
        'Add w:author naming whoever made the change.',
        id === null ? { element: local } : { id, element: local }
      ));
    } else if (author !== null && author.length > MAX_AUTHOR_LENGTH) {
      problems.push(revisionFinding(
        'author-too-long', partPath,
        `<w:${local}> has a w:author of ${author.length} characters; the schema caps it at ${MAX_AUTHOR_LENGTH}. What Word does with a longer one was not confirmed, so this is reported as out of spec rather than as a file Word will refuse.`,
        `Shorten w:author to ${MAX_AUTHOR_LENGTH} characters or fewer.`,
        { element: local }
      ));
    }

    // w:date is NOT required here — see the schema notes at the top of this file. Its
    // absence is reported as a note because it costs you the ordering of the review, not
    // because the file is invalid.
    if (date === null && !UNATTRIBUTED_ELEMENTS.has(local)) {
      problems.push(revisionFinding(
        'missing-date', partPath,
        `<w:${local}${id === null ? '' : ` w:id="${id}"`}> has no w:date. The schema does not require one on this type, so the file is valid — but nothing then says when the change was made, and revisions from different authors cannot be put in order.`,
        'Add w:date if the review order matters; leaving it off is permitted.',
        author === null ? { element: local } : { element: local, author }
      ));
    }
  }

  // --- move range markers, which have stricter requirements --------------------
  for (const el of elements) {
    if (!(el.localName in MOVE_RANGE_STARTS)) continue;
    const id = attr(el, 'id');
    const subject = id === null ? { element: el.localName } : { element: el.localName, id };

    if (attr(el, 'name') === null) {
      problems.push(revisionFinding(
        'unnamed-move-range', partPath,
        `<w:${el.localName}${id === null ? '' : ` w:id="${id}"`}> has no w:name. On CT_MoveBookmark the name is required, and it is the only thing that ties the two halves of a move together — without it the move reads as an unrelated deletion and an unrelated insertion.`,
        'Add a w:name matching the other half of the move.',
        subject
      ));
    }
    if (attr(el, 'date') === null) {
      problems.push(revisionFinding(
        'missing-required-date', partPath,
        `<w:${el.localName}${id === null ? '' : ` w:id="${id}"`}> has no w:date. Unlike w:ins and w:del, CT_MoveBookmark requires it, so this one is a schema violation rather than a missing convenience.`,
        'Add w:date to the move range marker.',
        subject
      ));
    }
    if (attr(el, 'author') === null) {
      problems.push(revisionFinding(
        'missing-author', partPath,
        `<w:${el.localName}${id === null ? '' : ` w:id="${id}"`}> has no w:author, which CT_MoveBookmark requires.`,
        'Add w:author to the move range marker.',
        subject
      ));
    }
  }

  // --- moves, paired by name --------------------------------------------------
  const moves = pairMoves(ranges);
  const namesSeen = new Map<string, number>();
  for (const range of ranges) {
    if (range.name === '') continue;
    const key = `${MOVE_RANGE_STARTS[range.kind as keyof typeof MOVE_RANGE_STARTS]}:${range.name}`;
    namesSeen.set(key, (namesSeen.get(key) ?? 0) + 1);
  }
  for (const [key, count] of namesSeen) {
    if (count < 2) continue;
    const [half, name] = [key.slice(0, key.indexOf(':')), key.slice(key.indexOf(':') + 1)];
    problems.push(revisionFinding(
      'duplicate-move-name', partPath,
      `${count} <w:${half}RangeStart> markers are named "${name}". The name is what pairs the halves of a move, so with more than one candidate nothing in the markup decides which half goes with which.`,
      'Give each move a name used exactly once for its source and once for its destination.',
      { name, half }
    ));
  }
  for (const move of moves) {
    if (move.paired) continue;
    const present = move.from ? 'source' : 'destination';
    const absent = move.from ? 'destination' : 'source';
    problems.push(revisionFinding(
      'unpaired-move', partPath,
      `Move "${move.name}" has a ${present} but no ${absent} in this part. Word can no longer show it as a move: the half that is here degrades into a plain ${move.from ? 'deletion' : 'insertion'}, so a reviewer sees content vanishing or appearing with no indication it came from anywhere.`,
      `Restore the <w:move${move.from ? 'To' : 'From'}RangeStart w:name="${move.name}"> half, or convert the remaining half to a plain w:${move.from ? 'del' : 'ins'}.`,
      { name: move.name }
    ));
  }

  // --- text that is in the file but not in the document ------------------------
  const deletedText = revisions
    .filter(r => r.kind === 'del' || r.kind === 'moveFrom')
    .reduce((total, r) => total + r.text.length, 0);
  if (deletedText > 0) {
    problems.push(revisionFinding(
      'deleted-text-retained', partPath,
      `${deletedText} character(s) of deleted text are still stored in ${partPath || 'this part'} inside w:delText. Deleting text with change tracking on does not remove it from the file — anyone who unzips the package can read it, so this is not redaction. In the other direction, an extractor that reads text nodes without checking their ancestors emits this content as though the document said it.`,
      'Accept the deletions before distributing the file if the removed wording is not meant to travel with it.',
      { characters: String(deletedText) }
    ));
  }

  // Which spelling a run uses is decided by the INNERMOST enclosing revision: a w:ins
  // nested inside a w:del legitimately holds w:t.
  let liveInDeletion = 0;
  let orphanDeleted = 0;
  for (const el of elements) {
    if (!isW(el, 't') && !isW(el, 'delText')) continue;
    const innermost = enclosingRevisions(el)[0] ?? null;
    const inDeletion = innermost !== null && REMOVED_BY_ACCEPT.has(innermost);
    if (isW(el, 't') && inDeletion) liveInDeletion += 1;
    if (isW(el, 'delText') && !inDeletion) orphanDeleted += 1;
  }
  if (liveInDeletion > 0) {
    problems.push(revisionFinding(
      'live-text-in-deletion', partPath,
      `${liveInDeletion} <w:t> element(s) sit inside a deletion, where deleted content is spelled <w:delText>. The schema permits it — CT_R admits either regardless of what encloses the run — so no validator objects, but any consumer that recognises deletions by looking for w:delText will treat this text as live and emit words the document does not say.`,
      'Rewrite the deleted runs to use <w:delText>, keeping the text unchanged.',
      { count: String(liveInDeletion) }
    ));
  }
  if (orphanDeleted > 0) {
    problems.push(revisionFinding(
      'orphan-deleted-text', partPath,
      `${orphanDeleted} <w:delText> element(s) are not inside any w:del or w:moveFrom, so the markup says the text is deleted and says nothing deleted it. Whether Word displays such a run was not confirmed, so this reports the inconsistency rather than predicting the rendering.`,
      'Wrap the runs in a w:del, or change <w:delText> to <w:t> if the text is meant to be live.',
      { count: String(orphanDeleted) }
    ));
  }

  // --- id collisions, reported from the revision side -------------------------
  const revisionElements = new Set(revisions.map(r => r.kind));
  for (const collision of findMarkupIdCollisions(doc)) {
    if (!collision.elements.some(name => revisionElements.has(name))) continue;
    problems.push(revisionFinding(
      'id-collision', partPath,
      `w:id "${collision.id}" is used by ${collision.elements.length} elements (${collision.elements.join(', ')}), at least one of them a tracked change. w:id is one space shared by revisions, bookmarks and permissions; Word rejects a file that reuses an id while lenient readers open it without complaint, so this reaches users having passed every test that opens the document.`,
      `Renumber the colliding markers from ${nextSafeMarkupId(doc)} upward, which is one past the highest id of any type in the part.`,
      { id: collision.id }
    ));
  }

  // --- the headline: this part has two readings -------------------------------
  const outcome = compareRevisionOutcomes(doc);
  if (revisions.length > 0) {
    const paragraphMarks = revisions.filter(r => r.category === 'paragraph-mark').length;
    const formatOnly = revisions.filter(r => r.category === 'format').length;

    problems.push(revisionFinding(
      'unaccepted-revisions', partPath,
      `${revisions.length} unaccepted tracked change(s) by ${byAuthor.size} author(s). ${
        outcome.differs
          ? 'The text with every change accepted is NOT the text with every change rejected, so this part says two different things and anything extracting text from it silently picks one.'
          : 'Accepting or rejecting every change produces the same text, so the changes cancel out — the revisions are formatting or structure, not wording.'
      }`,
      'Accept or reject the changes before treating the extracted text as the document’s content, or state which reading was taken.',
      { revisions: String(revisions.length), authors: String(byAuthor.size) }
    ));

    if (paragraphMarks > 0) {
      problems.push(revisionFinding(
        'paragraph-mark-revision', partPath,
        `${paragraphMarks} paragraph mark(s) are themselves inserted or deleted (w:pPr/w:rPr/w:ins or w:del). These touch no run, so nothing that walks runs can see them, and resolving them merges or splits paragraphs — which moves every heading level, list number and cross-reference after the change.`,
        'Treat paragraph-mark revisions as structural: resolve them before comparing paragraph counts or numbering between two versions of the document.',
        { count: String(paragraphMarks) }
      ));
    }

    if (formatOnly > 0) {
      problems.push(revisionFinding(
        'format-only-revision', partPath,
        `${formatOnly} format-only revision(s) (w:pPrChange, w:rPrChange, w:tblPrChange, w:sectPrChange and friends). Each records the PREVIOUS formatting rather than the new formatting, and changes no character, so no text comparison of any kind can detect them and no text extraction is affected by accepting them.`,
        'Compare formatting rather than text if these matter; the previous properties are inside each *Change element.',
        { count: String(formatOnly) }
      ));
    }
  }

  return { revisions, byAuthor, moves, problems };
}

/**
 * Pairs move halves by `@w:name` — the only thing that relates them.
 *
 * The range start and its end pair by `@w:id`, which is the bookmark reader's job. This
 * is the other pairing: source to destination, and the one that decides whether a move
 * is a move at all.
 */
export function pairMoves(ranges: readonly Bookmark[]): Move[] {
  const byName = new Map<string, Move>();
  for (const range of ranges) {
    const half = MOVE_RANGE_STARTS[range.kind as keyof typeof MOVE_RANGE_STARTS];
    if (half === undefined) continue;
    const move = byName.get(range.name) ?? { name: range.name, from: null, to: null, paired: false };
    if (half === 'moveFrom') move.from = move.from ?? range;
    else move.to = move.to ?? range;
    move.paired = move.from !== null && move.to !== null;
    byName.set(range.name, move);
  }
  return [...byName.values()];
}

export interface RevisionOutcome {
  /** What the document says with every change accepted. */
  accepted: string;
  /** What it said before any of them — every change rejected. */
  rejected: string;
  /**
   * What a tag-blind extractor produces: every `w:t` and `w:delText` concatenated.
   * When a part has both insertions and deletions this matches neither reading, which is
   * the failure the whole module exists to make visible.
   */
  naive: string;
  /** False when the two readings coincide — revisions that change format, not wording. */
  differs: boolean;
}

/**
 * The two texts: what this document says if every change is accepted, versus rejected.
 *
 * Paragraph marks are part of the answer, not a detail. A deleted paragraph mark means
 * the accepted reading has one fewer paragraph break than the rejected one, so the two
 * strings differ in structure even when every word survives.
 *
 * ⚠️ SCOPE. Only `w:t` and `w:delText` are collected, joined with `\n` at paragraph
 * boundaries. Tabs, breaks, symbols, field results and `w:delInstrText` are not
 * included, so the two strings are exactly comparable *to each other* but are not a full
 * text extraction — do not present either as the document's text.
 *
 * ⚠️ ASSUMPTION. Each paragraph mark is resolved independently. Whether Word does
 * anything different with two adjacent mark revisions by different authors was not
 * confirmed.
 */
export function compareRevisionOutcomes(doc: Document | Element): RevisionOutcome {
  const root = rootOf(doc);
  const elements = allElements(root);
  const paragraphs = elements.filter(el => isW(el, 'p'));

  // Text is bucketed by its NEAREST ancestor paragraph, so a paragraph inside a text box
  // inside a run does not have its content counted twice.
  const textByParagraph = new Map<Element, Element[]>();
  for (const el of elements) {
    if (!isW(el, 't') && !isW(el, 'delText')) continue;
    const paragraph = nearest(el, 'p');
    if (!paragraph) continue;
    const bucket = textByParagraph.get(paragraph);
    if (bucket) bucket.push(el);
    else textByParagraph.set(paragraph, [el]);
  }

  const accepted: string[] = [];
  const rejected: string[] = [];
  const naive: string[] = [];

  paragraphs.forEach((paragraph, index) => {
    for (const el of textByParagraph.get(paragraph) ?? []) {
      const text = el.textContent ?? '';
      const enclosing = enclosingRevisions(el);
      naive.push(text);
      if (!enclosing.some(kind => REMOVED_BY_ACCEPT.has(kind))) accepted.push(text);
      if (!enclosing.some(kind => REMOVED_BY_REJECT.has(kind))) rejected.push(text);
    }

    if (index === paragraphs.length - 1) return;
    const mark = paragraphMarkRevision(paragraph);
    naive.push('\n');
    // A deleted mark disappears when the deletion is accepted, so the paragraph merges
    // with the next; an inserted mark disappears when the insertion is rejected.
    if (!(mark !== null && REMOVED_BY_ACCEPT.has(mark))) accepted.push('\n');
    if (!(mark !== null && REMOVED_BY_REJECT.has(mark))) rejected.push('\n');
  });

  const acceptedText = accepted.join('');
  const rejectedText = rejected.join('');
  return {
    accepted: acceptedText,
    rejected: rejectedText,
    naive: naive.join(''),
    differs: acceptedText !== rejectedText
  };
}

export interface RevisionSettings {
  /** Every further edit to this document will be recorded as a revision. */
  trackRevisions: boolean;
  /** Moves will be recorded as an unrelated deletion plus insertion rather than as moves. */
  doNotTrackMoves: boolean;
  doNotTrackFormatting: boolean;
  /** Annotation types `w:revisionView` suppresses from display, in plain words. */
  hiddenFromView: string[];
}

/** `w:revisionView` attribute → what turning it off hides. Names taken from the schema. */
const REVISION_VIEW_LABELS: Array<[string, string]> = [
  ['markup', 'the visual indicator of the markup area'],
  ['comments', 'comments'],
  ['insDel', 'content revisions — insertions and deletions'],
  ['formatting', 'formatting revisions'],
  ['inkAnnotations', 'ink annotations']
];

/**
 * Reads the revision-related settings from `word/settings.xml`.
 *
 * `w:revisionView` is the interesting one: it can suppress the *display* of revisions
 * while leaving every one of them in the file. A document opened that way looks final and
 * is not, which is the same two-texts problem wearing a disguise.
 */
export function readRevisionSettings(doc: Document | Element | null): RevisionSettings {
  if (!doc) {
    return { trackRevisions: false, doNotTrackMoves: false, doNotTrackFormatting: false, hiddenFromView: [] };
  }
  const root = rootOf(doc);
  const elements = allElements(root);
  const first = (local: string) => elements.find(el => isW(el, local)) ?? null;

  const view = first('revisionView');
  const hiddenFromView: string[] = [];
  if (view) {
    for (const [name, label] of REVISION_VIEW_LABELS) {
      // OnOffValue: only an explicit off suppresses. An absent attribute leaves the
      // default alone, and reading absence as "hidden" would flag every document.
      const raw = attr(view, name);
      if (raw === '0' || raw === 'false' || raw === 'off') hiddenFromView.push(label);
    }
  }

  return {
    trackRevisions: onOff(first('trackRevisions')),
    doNotTrackMoves: onOff(first('doNotTrackMoves')),
    doNotTrackFormatting: onOff(first('doNotTrackFormatting')),
    hiddenFromView
  };
}

/**
 * The finding for a document that hides its own revisions from view.
 *
 * Kept separate from `readRevisions` because it needs two parts — the body to know there
 * are revisions, and `word/settings.xml` to know they are suppressed — and the analyzer
 * registry hands those to different calls.
 */
export function checkRevisionVisibility(
  settings: RevisionSettings,
  revisionCount: number,
  part = 'word/settings.xml'
): Finding[] {
  if (revisionCount === 0 || settings.hiddenFromView.length === 0) return [];
  return [
    revisionFinding(
      'revisions-hidden-in-view', part,
      `${part} suppresses the display of ${settings.hiddenFromView.join(' and ')} through w:revisionView, while the document still holds ${revisionCount} unaccepted change(s). A reader opens the file, sees no markup, and takes the text on screen for the final text — the changes are still there and still unresolved.`,
      'Remove the w:revisionView suppression so reviewers can see what is unresolved, or accept the changes.',
      { revisions: String(revisionCount) }
    )
  ];
}

/** Word body parts, each with its own revisions and its own id space. */
const WORD_BODY = /^word\/(?:document\d*|header[^/]*|footer[^/]*|footnotes\d*|endnotes\d*)\.xml$/;

const SETTINGS_PART = 'word/settings.xml';

const excerpt = (text: string, limit = 160): string =>
  text.length > limit ? `${text.slice(0, limit)}…` : text;

/**
 * Evidence lines for the AI panel.
 *
 * Leads with the two texts, because every other fact about tracked changes is a detail of
 * that one. Everything here is computed from the markup; nothing is asserted by a model.
 */
export function computeRevisionEvidenceForMarkup(
  parts: Record<string, string>,
  rawXml: string
): { lines: string[]; unresolved: string[] } | null {
  const entry = Object.entries(parts).find(([path]) => WORD_BODY.test(path));
  if (!entry) return null;

  const [path, xml] = entry;
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) return null;

  const index = readRevisions(doc, path);
  const outcome = compareRevisionOutcomes(doc);

  const settingsXml = parts[SETTINGS_PART];
  let settings: RevisionSettings = {
    trackRevisions: false,
    doNotTrackMoves: false,
    doNotTrackFormatting: false,
    hiddenFromView: []
  };
  if (settingsXml !== undefined) {
    const settingsDoc = parser.parseFromString(settingsXml, 'application/xml');
    if (settingsDoc.getElementsByTagName('parsererror').length === 0) {
      settings = readRevisionSettings(settingsDoc);
    }
  }

  if (index.revisions.length === 0 && index.problems.length === 0 && !settings.trackRevisions) return null;

  const lines: string[] = [];
  const unresolved: string[] = [];

  if (index.revisions.length === 0) {
    lines.push(
      `${path} contains no tracked changes, so the text it holds is the text it says — there is no second reading.`
    );
  } else {
    const authors = [...index.byAuthor.entries()]
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
      .map(([author, count]) => `${author} (${count})`);
    lines.push(
      `${path} holds ${index.revisions.length} unaccepted tracked change(s) by ${index.byAuthor.size} author(s): ${authors.join(', ')}.`
    );

    lines.push(
      outcome.differs
        ? `Accepting every change gives: "${excerpt(outcome.accepted)}". Rejecting every change gives: "${excerpt(outcome.rejected)}". These are two different documents and the file is both.`
        : `Accepting every change and rejecting every change both give: "${excerpt(outcome.accepted)}". The revisions here do not alter the wording.`
    );

    if (outcome.naive !== outcome.accepted) {
      lines.push(
        `A text extractor that ignores revision markup concatenates ${outcome.naive.length} characters against ${outcome.accepted.length} in the accepted reading, because w:delText is still a text node. Its output matches neither reading.`
      );
    }

    const counts = new Map<RevisionCategory, number>();
    for (const revision of index.revisions) counts.set(revision.category, (counts.get(revision.category) ?? 0) + 1);
    lines.push(
      `By kind: ${[...counts.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([c, n]) => `${c} ${n}`).join(', ')}.`
    );

    if (index.moves.length > 0) {
      const paired = index.moves.filter(m => m.paired).length;
      lines.push(
        `${index.moves.length} move range name(s), ${paired} with both halves present. A move whose halves do not pair reviews as an unrelated deletion plus an unrelated insertion.`
      );
    }
  }

  if (settings.trackRevisions) {
    lines.push(
      `${SETTINGS_PART} sets w:trackRevisions, so every further edit to this document will be recorded as a revision rather than applied to the text.`
    );
  }
  if (settings.doNotTrackMoves) {
    lines.push(
      `${SETTINGS_PART} sets w:doNotTrackMoves, so moved content is recorded as a plain deletion and a plain insertion — an absence of w:moveFrom/w:moveTo in this file does not mean nothing was moved.`
    );
  }
  if (settings.doNotTrackFormatting) {
    lines.push(
      `${SETTINGS_PART} sets w:doNotTrackFormatting, so formatting changes are applied without a *Change record — an absence of w:pPrChange does not mean the formatting was never edited.`
    );
  }

  // The selected element in the editor, when it is a revision.
  const selected = /<w:(ins|del|moveFrom|moveTo|pPrChange|rPrChange|sectPrChange|tbl(?:Pr|PrEx|Grid)Change|tcPrChange|trPrChange)\b[^>]*?w:id="(\d+)"/.exec(rawXml);
  if (selected) {
    const match = index.revisions.find(r => r.kind === selected[1] && r.id === selected[2]);
    if (match) {
      lines.push(
        `The selected <w:${match.kind} w:id="${match.id}"> is a ${match.category} revision by ${match.author ?? UNATTRIBUTED_AUTHOR}${match.date === null ? ' with no date' : ` dated ${match.date}`}${
          match.text === '' ? ' and encloses no text' : `, covering "${excerpt(match.text, 120)}"`
        }.`
      );
    }
  }

  lines.push(...renderFindings([...index.problems, ...checkRevisionVisibility(settings, index.revisions.length, SETTINGS_PART)]));

  unresolved.push(
    `Only ${path} was read. Headers, footers, footnotes and endnotes carry their own revisions in their own id space, so these counts are for one story, and a move with one half in another part cannot be paired from here.`
  );
  unresolved.push(
    'The accepted and rejected readings are built from w:t and w:delText joined at paragraph boundaries. Tabs, breaks, symbols, field results and w:delInstrText are excluded, so the two are comparable to each other but neither is a complete text extraction.'
  );
  if (index.revisions.some(r => r.date === null)) {
    unresolved.push(
      'Some revisions carry no w:date. The schema does not require one except on move range markers, so they cannot be ordered in time and no reading of the markup can recover when they were made.'
    );
  }
  if (index.revisions.some(r => r.category === 'format')) {
    unresolved.push(
      'What a format-only revision would look like once resolved is not determined here: the *Change element stores the previous properties, and deciding what the reader currently sees needs the style hierarchy resolved, not just the revision read.'
    );
  }

  return { lines, unresolved };
}
