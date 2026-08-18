/**
 * WordprocessingML comments — four parts, and why the answer is never where you look.
 *
 * A comment is not one thing in one place. It is a *range* in the body, a *body* in a
 * second part, and its *thread position and resolved state* in a third part that is not
 * part of ECMA-376 at all. Every question a reader actually asks is answered by a
 * different file, and two of them are answered by a file that is allowed to be missing.
 *
 *   word/document.xml       <w:commentRangeStart w:id="3"/> …content…
 *                           <w:commentRangeEnd w:id="3"/>
 *                           <w:r><w:commentReference w:id="3"/></w:r>
 *   word/comments.xml       <w:comment w:id="3" w:author="…" w:date="…">…paragraphs…</w:comment>
 *   word/commentsExtended.xml   <w15:commentEx w15:paraId="…" w15:paraIdParent="…" w15:done="1"/>
 *   word/commentsIds.xml    <w16cid:commentId w16cid:paraId="…" w16cid:durableId="…"/>
 *
 * Four consequences follow, and each one is a class of bug.
 *
 * 1. THE RANGE IS MATCHED BY @w:id, AND THE MARKERS ARE NOT INSIDE RUNS.
 *    commentRangeStart and commentRangeEnd are the same CT_MarkupRange type bookmarks
 *    use, sitting as siblings of runs rather than inside them. A start with no end, or
 *    an end with no start, does not produce a malformed comment — it produces a comment
 *    whose highlight silently covers the wrong span, or nothing. The document opens and
 *    nothing looks wrong.
 *
 * 2. "WHAT TEXT DOES COMMENT 3 COVER?" IS NOT STORED ANYWHERE.
 *    Start and end are independent markers, so a range can begin mid-paragraph and end
 *    inside a different table cell; nesting and overlap are both legal and both routine
 *    (Word produces nesting whenever someone comments on a phrase inside a commented
 *    sentence). The covered text therefore has to be *walked* between two points in
 *    document order — see `commentRangeText`, which is the same walk `bookmarkText`
 *    performs, including its distinction between `null` (no answer possible) and `''`
 *    (a real answer: the comment is anchored to an insertion point covering no text).
 *
 * 3. THE BODY AND THE ANCHOR CAN EXIST WITHOUT EACH OTHER.
 *    A commentReference naming an id with no w:comment renders as a comment with no
 *    text. A w:comment nothing in the body references never displays at all — it is
 *    present in the package, invisible in Word, and fully visible to anyone who unzips
 *    the file, which is how "deleted" review comments keep shipping to counterparties.
 *
 * 4. THREADING AND RESOLVED-STATE ARE NOT IN comments.xml, AND MAY NOT EXIST.
 *    Replies and the "Resolved" flag live only in word/commentsExtended.xml, an Office
 *    2013 side-car in the w15 namespace. It keys on the **w14:paraId of the comment's
 *    last paragraph** — not on the comment id — so the join runs through an attribute on
 *    a paragraph, and a comment whose last paragraph has no w14:paraId cannot be keyed
 *    at all. When the part is absent, "is this a reply?" and "is this resolved?" are
 *    *unknowable*. This module reports them as unknown (`null`) and never as `false`:
 *    answering "not resolved" for a thread that was resolved, or "not a reply" for a
 *    reply, is the failure this module exists to prevent.
 *
 * VERIFIED against the Open XML SDK schema data (data/schemas, data/namespaces.json):
 *   - w15 = `http://schemas.microsoft.com/office/word/2012/wordml` and
 *     w14 = `http://schemas.microsoft.com/office/word/2010/wordml`. Note there is **no**
 *     trailing `/main` on either, unlike the ECMA namespaces.
 *   - `w15:commentEx` carries `w15:paraId` (required, 4-byte hex), `w15:paraIdParent`
 *     (optional, 4-byte hex) and `w15:done` (OnOffValue). Root element `w15:commentsEx`,
 *     part class WordprocessingCommentsExPart.
 *   - `w14:paraId` is an attribute of `w:p` (CT_P), 4-byte hex, Office 2010+.
 *   - `w:comment` (CT_Comment): `w:id` **required**, `w:author` **required** (max 255),
 *     `w:initials` optional (max 9), `w:date` optional.
 *   - `w:commentRangeStart`/`w:commentRangeEnd` are CT_MarkupRange and
 *     `w:commentReference` is CT_Markup — so all three take the same `@w:id` union as
 *     bookmarks: a non-negative number, or a signed number ≤ -2. -1 is excluded.
 *   - w16cid = `http://schemas.microsoft.com/office/word/2016/wordml/cid`, root
 *     `w16cid:commentsIds`, entries `w16cid:commentId` with `w16cid:paraId` and
 *     `w16cid:durableId`, both required.
 *
 * NOT VERIFIED — stated as behaviour, not as a citation:
 *   - That Word *requires* a `w:commentReference` for a comment to display, and that a
 *     range alone will not anchor it. This is consistently observed but the schema makes
 *     every one of the three markers optional, so `missing-reference` is reported as a
 *     warning about likely non-display rather than as a schema violation.
 *   - That `w15:paraIdParent` always names the *root* of a thread rather than the
 *     immediate predecessor. Word's UI presents threads flat; this module resolves the
 *     parent link transitively so both readings give the same thread membership.
 *   - Whether Word regenerates a missing `commentsExtended.xml` on save, or discards the
 *     thread structure. Untested; the module reports unknown either way.
 *   - `w16cex` (`.../word/2018/wordml/cex`, `commentsExtensible.xml`) carries a UTC date
 *     and a "not yet resolved" placeholder keyed on durableId. Its element and attribute
 *     names were read from the SDK schema, but this module does not parse it, so nothing
 *     here depends on that reading.
 */

import { W_NAMESPACE } from './wordStyleResolver';
import { finding, renderFindings, type Finding, type Severity } from './findings';

/** Office 2013 side-car namespace: threading and resolved-state. No trailing `/main`. */
export const W15_NAMESPACE = 'http://schemas.microsoft.com/office/word/2012/wordml';
/** Office 2010 namespace: `w14:paraId`, the key the side-car joins on. */
export const W14_NAMESPACE = 'http://schemas.microsoft.com/office/word/2010/wordml';
/** Office 2016 namespace: `commentsIds.xml`, paraId → durable id. */
export const W16CID_NAMESPACE = 'http://schemas.microsoft.com/office/word/2016/wordml/cid';

/** `CT_Comment/@w:author` carries a StringValidator with MaxLength 255. */
export const MAX_COMMENT_AUTHOR_LENGTH = 255;
/** `CT_Comment/@w:initials` carries a StringValidator with MaxLength 9. */
export const MAX_COMMENT_INITIALS_LENGTH = 9;

/**
 * Severity, silence, and WHICH PART each kind is really about.
 *
 * The last column matters: comments live across three parts, and a finding that says
 * `word/document.xml` when the fault is in `commentsExtended.xml` sends the reader to
 * the wrong file.
 *
 * On silence: a comment that displays with no text is VISIBLE — someone sees the empty
 * bubble. A comment that never displays at all is SILENT, and so is every threading
 * fault, because the margin looks perfectly normal when replies have quietly flattened
 * into top-level comments.
 */
const COMMENT_RULES = {
  'unmatched-range-start': { severity: 'error',   silent: true,  where: 'body' },
  'unmatched-range-end':   { severity: 'warning', silent: true,  where: 'body' },
  'missing-reference':     { severity: 'warning', silent: true,  where: 'body' },
  'missing-body':          { severity: 'error',   silent: false, where: 'comments' },
  'comments-part-missing': { severity: 'error',   silent: false, where: 'comments' },
  'orphan-comment':        { severity: 'warning', silent: true,  where: 'comments' },
  'duplicate-id':          { severity: 'error',   silent: true,  where: 'comments' },
  'id-out-of-range':       { severity: 'error',   silent: false, where: 'body' },
  'reversed-range':        { severity: 'warning', silent: true,  where: 'body' },
  'overlapping-range':     { severity: 'note',    silent: true,  where: 'body' },
  'nested-range':          { severity: 'note',    silent: true,  where: 'body' },
  'missing-author':        { severity: 'warning', silent: false, where: 'comments' },
  'author-too-long':       { severity: 'error',   silent: false, where: 'comments' },
  'initials-too-long':     { severity: 'error',   silent: false, where: 'comments' },
  'missing-para-id':       { severity: 'warning', silent: true,  where: 'comments' },
  'threading-unknown':     { severity: 'note',    silent: true,  where: 'extended' },
  'dangling-parent':       { severity: 'warning', silent: true,  where: 'extended' },
  'orphan-comment-ex':     { severity: 'note',    silent: true,  where: 'extended' }
} as const satisfies Record<string, { severity: Severity; silent: boolean; where: CommentPartRole }>;

type CommentPartRole = 'body' | 'comments' | 'extended';

export type CommentProblemKind = keyof typeof COMMENT_RULES;

/** Where each role lives in a conventional package. */
export const COMMENT_PART_PATHS: Record<CommentPartRole, string> = {
  body: 'word/document.xml',
  comments: 'word/comments.xml',
  extended: 'word/commentsExtended.xml'
};

/**
 * Turns a problem description into a Finding, applying the table above.
 *
 * Takes an object rather than positional arguments so the call sites keep naming their
 * fields — with eighteen kinds and two optional identifiers, positional arguments would
 * be unreadable and easy to transpose.
 */
const commentFinding = (input: {
  kind: CommentProblemKind;
  message: string;
  remediation: string;
  id?: string;
  paraId?: string;
}): Finding => {
  const rule = COMMENT_RULES[input.kind];
  const subject: Record<string, string> = {};
  if (input.id !== undefined) subject.id = input.id;
  if (input.paraId !== undefined) subject.paraId = input.paraId;
  return finding(
    `comment/${input.kind}`,
    COMMENT_PART_PATHS[rule.where],
    input.message,
    input.remediation,
    {
      severity: rule.severity,
      silent: rule.silent,
      ...(Object.keys(subject).length > 0 ? { subject } : {})
    }
  );
};

/** Where a comment attaches in the body. All three markers are independent. */
export interface CommentAnchor {
  id: string;
  /** null when only an end or only a reference carries this id. */
  rangeStart: Element | null;
  /** null when the range was opened and never closed. */
  rangeEnd: Element | null;
  /** The `w:commentReference` run marking the anchor point; null when absent. */
  reference: Element | null;
}

/**
 * Thread position and resolved state, which live in a part that is allowed to be absent.
 *
 * `known` is false whenever `commentsExtended.xml` was not supplied, or supplied without
 * an entry for this comment. When it is false every other field is null, and `null` must
 * be rendered as "unknown" — never coerced to false.
 */
export interface CommentThreadInfo {
  known: boolean;
  /** null when unknowable. */
  isReply: boolean | null;
  /** `w15:paraIdParent` verbatim; null when this is a root comment or unknowable. */
  parentParaId: string | null;
  /** The parent's `w:comment/@w:id`, when the paraId resolves to a comment. */
  parentId: string | null;
  /** `w15:done`. null when unknowable. */
  resolved: boolean | null;
}

export interface Comment {
  id: string;
  /** @w:author is required; null means it was missing, not that it was empty. */
  author: string | null;
  date: string | null;
  initials: string | null;
  /** The `w:comment` element itself. */
  element: Element;
  /** The comment's own text, runs joined; excludes deleted text and field codes. */
  text: string;
  /** `w14:paraId` of the **last** paragraph — the key commentsExtended joins on. */
  paraId: string | null;
  /** `w16cid:durableId` for this paraId, when commentsIds.xml was supplied. */
  durableId: string | null;
  /** Where it attaches in the body; null when nothing in the body references it. */
  anchor: CommentAnchor | null;
  thread: CommentThreadInfo;
}

export interface CommentIndex {
  comments: Comment[];
  byId: Map<string, Comment>;
  /** Every id the body refers to, including ids with no `w:comment` behind them. */
  anchors: CommentAnchor[];
  anchorsById: Map<string, CommentAnchor>;
  /** False when threading and resolved-state could not be determined for any comment. */
  threadingKnown: boolean;
  problems: Finding[];
}

/** A root comment and every reply beneath it, flattened in document order. */
export interface CommentThread {
  root: Comment;
  replies: Comment[];
  /** The root's `w15:done`. null only if the root itself could not be keyed. */
  resolved: boolean | null;
}

export interface CommentParts {
  /** Parsed `word/document.xml`, or any body part carrying the range markers. */
  document: Document | Element;
  /** Parsed `word/comments.xml`. Omit or pass null when the part is absent. */
  comments?: Document | Element | null;
  /** Parsed `word/commentsExtended.xml`. Absent means threading is unknowable. */
  commentsExtended?: Document | Element | null;
  /** Parsed `word/commentsIds.xml`. Absent only costs the durable id. */
  commentsIds?: Document | Element | null;
}

const rootOf = (node: Document | Element): ParentNode =>
  'documentElement' in node && node.documentElement ? node.documentElement : (node as Element);

const attr = (el: Element, name: string) => el.getAttributeNS(W_NAMESPACE, name);

/** All descendants of `root` in `ns` whose local name is in `names`, in document order. */
const collect = (root: ParentNode, ns: string, names: string[]): Element[] => {
  const wanted = new Set(names);
  return Array.from(root.querySelectorAll('*')).filter(
    el => el.namespaceURI === ns && wanted.has(el.localName)
  );
};

/**
 * `w:id` is `ST_NonNegativeDecimalNumber` unioned with signed numbers ≤ -2, so -1 is the
 * one integer the union excludes. Identical to the bookmark id rule — same base types.
 */
const idIsInRange = (raw: string): boolean => {
  if (!/^-?\d+$/.test(raw)) return false;
  const n = Number(raw);
  return n >= 0 || n <= -2;
};

/** ST_OnOff: 1/true/on are true, 0/false/off are false. Anything else is not a value. */
const onOff = (raw: string | null): boolean | null => {
  if (raw === null) return null;
  if (raw === '1' || raw === 'true' || raw === 'on') return true;
  if (raw === '0' || raw === 'false' || raw === 'off') return false;
  return null;
};

/**
 * Text of a run-bearing element: `w:t` only. `w:delText` is excluded because it is not
 * part of the document as read, and `w:instrText` because it is code, not content.
 */
const textOf = (el: Element): string =>
  collect(el, W_NAMESPACE, ['t'])
    .map(t => t.textContent ?? '')
    .join('');

/**
 * Read every comment in a package, joining the four parts that each hold one third of
 * the answer, and reporting what could not be joined rather than throwing.
 *
 * Pass whichever parts exist. A missing `comments.xml` is reported, not assumed empty; a
 * missing `commentsExtended.xml` makes threading unknown rather than absent.
 */
export function readComments(parts: CommentParts): CommentIndex {
  const problems: Finding[] = [];
  const docRoot = rootOf(parts.document);

  const { anchors, anchorsById } = readAnchors(docRoot, problems);
  const { comments, byId } = readBodies(parts.comments ?? null, problems);

  if (!parts.comments && anchors.length > 0) {
    problems.push(commentFinding({
      kind: 'comments-part-missing',
      message: `The body refers to ${anchors.length} comment${anchors.length === 1 ? '' : 's'} but word/comments.xml was not supplied. Every comment shows in the margin with no text.`,
      remediation: 'Restore word/comments.xml, or delete the range markers and commentReference runs that point into it.'
    }));
  }

  // An id the body names with nothing behind it: the comment displays, empty.
  for (const anchor of anchors) {
    if (byId.has(anchor.id)) continue;
    if (!parts.comments) continue; // already reported once, at the part level
    problems.push(commentFinding({
      kind: 'missing-body',
      id: anchor.id,
      message: `The body references comment ${anchor.id} but word/comments.xml has no <w:comment w:id="${anchor.id}">. The comment is anchored and highlighted but has no text.`,
      remediation: `Add the missing <w:comment w:id="${anchor.id}">, or remove the range markers and commentReference for id ${anchor.id}.`
    }));
  }

  attachAnchors(comments, anchorsById, problems);
  const threadingKnown = resolveThreading(comments, parts.commentsExtended ?? null, problems);
  readDurableIds(comments, parts.commentsIds ?? null);

  return { comments, byId, anchors, anchorsById, threadingKnown, problems };
}

/** Pair range markers by id and collect reference runs. */
function readAnchors(docRoot: ParentNode, problems: Finding[]) {
  const starts = collect(docRoot, W_NAMESPACE, ['commentRangeStart']);
  const ends = collect(docRoot, W_NAMESPACE, ['commentRangeEnd']);
  const references = collect(docRoot, W_NAMESPACE, ['commentReference']);

  // Document-order position of every marker, for the overlap/nesting comparison below.
  const order = new Map<Element, number>();
  for (const el of collect(docRoot, W_NAMESPACE, [
    'commentRangeStart',
    'commentRangeEnd',
    'commentReference'
  ])) {
    order.set(el, order.size);
  }

  const endsById = new Map<string, Element[]>();
  for (const end of ends) {
    const id = attr(end, 'id');
    if (id === null) continue;
    endsById.set(id, [...(endsById.get(id) ?? []), end]);
  }

  const anchors: CommentAnchor[] = [];
  const anchorsById = new Map<string, CommentAnchor>();
  const claimedEnds = new Set<Element>();

  const anchorFor = (id: string): CommentAnchor => {
    const existing = anchorsById.get(id);
    if (existing) return existing;
    const created: CommentAnchor = { id, rangeStart: null, rangeEnd: null, reference: null };
    anchorsById.set(id, created);
    anchors.push(created);
    return created;
  };

  for (const start of starts) {
    const id = attr(start, 'id');
    if (id === null) {
      problems.push(commentFinding({
        kind: 'unmatched-range-start',
        message: 'A <w:commentRangeStart> has no w:id, so nothing can close it and no comment can claim it. @w:id is required.',
        remediation: 'Give the marker the id of the comment it opens, or delete it.'
      }));
      continue;
    }
    if (!idIsInRange(id)) {
      problems.push(commentFinding({
        kind: 'id-out-of-range',
        id,
        message: `Comment range id "${id}" is outside the permitted range. The type admits any non-negative number, or a negative number of -2 or below; -1 and non-integers are excluded.`,
        remediation: 'Renumber the comment and all three of its markers to a non-negative integer.'
      }));
    }

    const anchor = anchorFor(id);
    if (anchor.rangeStart !== null) {
      problems.push(commentFinding({
        kind: 'duplicate-id',
        id,
        message: `Two <w:commentRangeStart w:id="${id}"/> markers open the same comment. Ends match by id, so the second start competes with the first for the same end and one of the two highlights silently covers the wrong span.`,
        remediation: 'Give each comment its own id across document.xml and comments.xml.'
      }));
      continue;
    }
    anchor.rangeStart = start;

    const end = (endsById.get(id) ?? []).find(candidate => !claimedEnds.has(candidate)) ?? null;
    if (end) {
      claimedEnds.add(end);
      anchor.rangeEnd = end;
      const startPos = order.get(start);
      const endPos = order.get(end);
      if (startPos !== undefined && endPos !== undefined && endPos < startPos) {
        problems.push(commentFinding({
          kind: 'reversed-range',
          id,
          message: `Comment ${id} has its range end before its range start, so it highlights no text and the anchor point is ambiguous.`,
          remediation: 'Move the end marker after the start, or swap the two.'
        }));
      }
    } else {
      problems.push(commentFinding({
        kind: 'unmatched-range-start',
        id,
        message: `Comment ${id} opens a range that never closes — no <w:commentRangeEnd w:id="${id}"/> exists. The highlight has no defined extent, so what the comment appears to be about depends on the reader's application rather than on the document.`,
        remediation: `Insert <w:commentRangeEnd w:id="${id}"/> at the point the commented text should end.`
      }));
    }
  }

  for (const end of ends) {
    if (claimedEnds.has(end)) continue;
    const id = attr(end, 'id');
    problems.push(commentFinding({
      kind: 'unmatched-range-end',
      id: id ?? undefined,
      message: `A <w:commentRangeEnd${id === null ? '' : ` w:id="${id}"`}/> closes a comment range that was never opened. The start carried no name and neither does the end, so there is no way to recover where the highlight was meant to begin.`,
      remediation: 'Delete the stray end, or restore the matching <w:commentRangeStart>.'
    }));
  }

  for (const reference of references) {
    const id = attr(reference, 'id');
    if (id === null) continue;
    const anchor = anchorFor(id);
    if (anchor.reference === null) anchor.reference = reference;
  }

  for (const anchor of anchors) {
    if (anchor.reference !== null) continue;
    problems.push(commentFinding({
      kind: 'missing-reference',
      id: anchor.id,
      message: `Comment ${anchor.id} has range markers but no <w:commentReference w:id="${anchor.id}"/> run. The reference is the anchor point Word draws the balloon from; without it the comment is generally not displayed at all, even though its text is in the package.`,
      remediation: `Add <w:r><w:commentReference w:id="${anchor.id}"/></w:r> at the end of the commented range.`
    }));
  }

  reportRangeOverlaps(anchors, order, problems);
  return { anchors, anchorsById };
}

/**
 * Overlapping and nested ranges are both legal, and nesting is routine. They are
 * surfaced because they change what "the text of comment N" means: a nested comment's
 * text is a subset of its parent's, and overlapping highlights are what produce the
 * "why is this sentence coloured twice" question no single element answers.
 */
function reportRangeOverlaps(
  anchors: CommentAnchor[],
  order: Map<Element, number>,
  problems: Finding[]
) {
  const spans = anchors
    .filter(a => a.rangeStart !== null && a.rangeEnd !== null)
    .map(a => ({
      id: a.id,
      start: order.get(a.rangeStart!) ?? -1,
      end: order.get(a.rangeEnd!) ?? -1
    }))
    .filter(s => s.start >= 0 && s.end > s.start)
    .sort((a, b) => a.start - b.start);

  for (let i = 0; i < spans.length; i += 1) {
    for (let j = i + 1; j < spans.length; j += 1) {
      const outer = spans[i];
      const inner = spans[j];
      if (inner.start > outer.end) break; // sorted by start: nothing later can touch it
      if (inner.end <= outer.end) {
        problems.push(commentFinding({
          kind: 'nested-range',
          id: inner.id,
          message: `Comment ${inner.id}'s range sits entirely inside comment ${outer.id}'s. This is legal and common, but the two comments cover overlapping text, so "the text comment ${outer.id} refers to" includes everything comment ${inner.id} refers to.`,
          remediation: 'No action needed. Surfaced so that tools reporting commented text do not present the two ranges as independent.'
        }));
      } else {
        problems.push(commentFinding({
          kind: 'overlapping-range',
          id: inner.id,
          message: `Comment ${inner.id}'s range starts inside comment ${outer.id}'s and ends outside it. Partial overlap is legal but neither comment contains the other, so the highlights interleave and no nesting order exists.`,
          remediation: 'No action needed, unless the overlap was accidental — in which case move one end marker so the ranges either nest or separate.'
        }));
      }
    }
  }
}

/** Read `word/comments.xml` into Comment records, validating what the schema requires. */
function readBodies(commentsPart: Document | Element | null, problems: Finding[]) {
  const comments: Comment[] = [];
  const byId = new Map<string, Comment>();
  if (!commentsPart) return { comments, byId };

  for (const el of collect(rootOf(commentsPart), W_NAMESPACE, ['comment'])) {
    const id = attr(el, 'id');
    if (id === null) {
      problems.push(commentFinding({
        kind: 'duplicate-id',
        message: 'A <w:comment> has no w:id, so nothing in the body can reference it and it can never display. @w:id is required.',
        remediation: 'Give the comment the id used by its commentReference, or delete it.'
      }));
      continue;
    }
    if (!idIsInRange(id)) {
      problems.push(commentFinding({
        kind: 'id-out-of-range',
        id,
        message: `Comment id "${id}" is outside the permitted range. The type admits any non-negative number, or a negative number of -2 or below; -1 and non-integers are excluded.`,
        remediation: 'Renumber the comment and all three of its body markers to a non-negative integer.'
      }));
    }

    const author = attr(el, 'author');
    const initials = attr(el, 'initials');
    if (author === null) {
      problems.push(commentFinding({
        kind: 'missing-author',
        id,
        message: `Comment ${id} has no w:author. The attribute is required, and it is the only thing that identifies who said this — readers see an unattributed comment.`,
        remediation: 'Add a w:author. Use a deliberate placeholder rather than an empty string if the real name must not ship.'
      }));
    } else if (author.length > MAX_COMMENT_AUTHOR_LENGTH) {
      problems.push(commentFinding({
        kind: 'author-too-long',
        id,
        message: `Comment ${id} has an author of ${author.length} characters; the maximum is ${MAX_COMMENT_AUTHOR_LENGTH}.`,
        remediation: `Shorten the author name to ${MAX_COMMENT_AUTHOR_LENGTH} characters or fewer.`
      }));
    }
    if (initials !== null && initials.length > MAX_COMMENT_INITIALS_LENGTH) {
      problems.push(commentFinding({
        kind: 'initials-too-long',
        id,
        message: `Comment ${id} has initials of ${initials.length} characters; the maximum is ${MAX_COMMENT_INITIALS_LENGTH}.`,
        remediation: `Shorten the initials to ${MAX_COMMENT_INITIALS_LENGTH} characters or fewer.`
      }));
    }

    // commentsExtended keys on the LAST paragraph's w14:paraId, not the first and not
    // the comment id. Getting this wrong joins a reply to the wrong thread.
    const paragraphs = collect(el, W_NAMESPACE, ['p']);
    const lastParagraph = paragraphs[paragraphs.length - 1] ?? null;
    const paraId = lastParagraph?.getAttributeNS(W14_NAMESPACE, 'paraId') ?? null;

    const comment: Comment = {
      id,
      author,
      date: attr(el, 'date'),
      initials,
      element: el,
      text: textOf(el),
      paraId,
      durableId: null,
      anchor: null,
      thread: { known: false, isReply: null, parentParaId: null, parentId: null, resolved: null }
    };

    if (byId.has(id)) {
      problems.push(commentFinding({
        kind: 'duplicate-id',
        id,
        message: `Two <w:comment> elements share w:id "${id}". A commentReference names one id, so one of the two bodies is unreachable and which one displays is not something the markup determines.`,
        remediation: 'Renumber one of them, and update its range markers and commentReference to match.'
      }));
    } else {
      byId.set(id, comment);
    }
    comments.push(comment);
  }

  return { comments, byId };
}

/** Link bodies to anchors, and report the bodies nothing in the document points at. */
function attachAnchors(
  comments: Comment[],
  anchorsById: Map<string, CommentAnchor>,
  problems: Finding[]
) {
  for (const comment of comments) {
    const anchor = anchorsById.get(comment.id) ?? null;
    comment.anchor = anchor;
    if (anchor !== null) continue;
    problems.push(commentFinding({
      kind: 'orphan-comment',
      id: comment.id,
      message: `Comment ${comment.id}${comment.author === null ? '' : ` by ${comment.author}`} is in word/comments.xml but nothing in the document references it, so Word never displays it. Its text still ships inside the file and is readable by anyone who unzips it.`,
      remediation: 'Delete the <w:comment>, or add the range markers and commentReference that were meant to anchor it.'
    }));
  }
}

/**
 * Join `commentsExtended.xml` on paraId and fill in `thread`.
 *
 * Returns whether thread position and resolved state are known for *every* comment. The
 * two states this function exists to keep apart are "no reply, not resolved" and "we
 * cannot tell". A document with no comments is vacuously known.
 */
function resolveThreading(
  comments: Comment[],
  extendedPart: Document | Element | null,
  problems: Finding[]
): boolean {
  if (!extendedPart) {
    if (comments.length === 0) return true;
    problems.push(commentFinding({
      kind: 'threading-unknown',
      message: `word/commentsExtended.xml was not supplied, so for all ${comments.length} comment${comments.length === 1 ? '' : 's'} it cannot be determined which are replies or which threads are resolved. Both are stored only in that part.`,
      remediation: 'Supply word/commentsExtended.xml. Until then, present reply and resolved state as unknown rather than as "no" — a resolved thread shown as open is the failure this reports.'
    }));
    return false;
  }

  const byParaId = new Map<string, Comment>();
  for (const comment of comments) {
    if (comment.paraId !== null && !byParaId.has(comment.paraId)) {
      byParaId.set(comment.paraId, comment);
    }
  }

  const extras = collect(rootOf(extendedPart), W15_NAMESPACE, ['commentEx']);
  const matched = new Set<Comment>();

  for (const ex of extras) {
    const paraId = ex.getAttributeNS(W15_NAMESPACE, 'paraId');
    if (paraId === null) continue;
    const comment = byParaId.get(paraId);
    if (!comment) {
      problems.push(commentFinding({
        kind: 'orphan-comment-ex',
        paraId,
        message: `A <w15:commentEx w15:paraId="${paraId}"/> names a paragraph that no comment's last paragraph carries. Whatever thread position or resolved state it recorded applies to nothing.`,
        remediation: 'Delete the stray commentEx, or restore the w14:paraId on the comment paragraph it was written for.'
      }));
      continue;
    }
    matched.add(comment);

    const parentParaId = ex.getAttributeNS(W15_NAMESPACE, 'paraIdParent');
    const parent = parentParaId === null ? null : (byParaId.get(parentParaId) ?? null);
    if (parentParaId !== null && parent === null) {
      problems.push(commentFinding({
        kind: 'dangling-parent',
        id: comment.id,
        paraId: parentParaId,
        message: `Comment ${comment.id} is a reply to paragraph ${parentParaId}, but no comment carries that w14:paraId. The comment it answers is gone, so the reply displays detached from the question it was written against.`,
        remediation: 'Restore the parent comment, or drop w15:paraIdParent so the reply becomes a top-level comment instead of a broken one.'
      }));
    }

    comment.thread = {
      known: true,
      isReply: parentParaId !== null,
      parentParaId,
      parentId: parent?.id ?? null,
      resolved: onOff(ex.getAttributeNS(W15_NAMESPACE, 'done')) ?? false
    };
  }

  let complete = true;
  for (const comment of comments) {
    if (matched.has(comment)) continue;
    complete = false;
    if (comment.paraId === null) {
      problems.push(commentFinding({
        kind: 'missing-para-id',
        id: comment.id,
        message: `Comment ${comment.id}'s last paragraph has no w14:paraId, which is the only key commentsExtended.xml joins on. Whether it is a reply, and whether its thread is resolved, cannot be determined for this comment even though the part is present.`,
        remediation: 'Add a w14:paraId to the comment\'s last w:p and a matching w15:commentEx entry. Word writes both on save.'
      }));
    } else {
      problems.push(commentFinding({
        kind: 'threading-unknown',
        id: comment.id,
        paraId: comment.paraId,
        message: `commentsExtended.xml has no <w15:commentEx> for paraId ${comment.paraId}, so comment ${comment.id}'s reply and resolved state are unknown. An absent entry is missing information, not a record that the comment is a resolved-free root.`,
        remediation: `Add <w15:commentEx w15:paraId="${comment.paraId}"/> with the intended w15:paraIdParent and w15:done.`
      }));
    }
  }

  return complete;
}

/** `commentsIds.xml` maps the same paraId to an id that survives editing. */
function readDurableIds(comments: Comment[], idsPart: Document | Element | null) {
  if (!idsPart) return;
  const durableByParaId = new Map<string, string>();
  for (const el of collect(rootOf(idsPart), W16CID_NAMESPACE, ['commentId'])) {
    const paraId = el.getAttributeNS(W16CID_NAMESPACE, 'paraId');
    const durableId = el.getAttributeNS(W16CID_NAMESPACE, 'durableId');
    if (paraId !== null && durableId !== null) durableByParaId.set(paraId, durableId);
  }
  for (const comment of comments) {
    if (comment.paraId === null) continue;
    comment.durableId = durableByParaId.get(comment.paraId) ?? null;
  }
}

/**
 * The text a comment covers, which is the question people actually ask and the one no
 * single element answers.
 *
 * Returns null when there is no answer to give — the range never opened, or opened and
 * never closed. That is deliberately distinct from `''`, which is a real answer: the
 * comment is anchored to a point rather than to a span, which is what a comment on an
 * empty paragraph or on a picture looks like.
 *
 * Deleted text (`w:delText`) is excluded because it is not part of the document as read,
 * and field instructions (`w:instrText`) because they are code, not content.
 */
export function commentRangeText(anchor: CommentAnchor | null): string | null {
  if (!anchor || !anchor.rangeStart || !anchor.rangeEnd) return null;
  const root = anchor.rangeStart.ownerDocument?.documentElement;
  if (!root) return null;

  const parts: string[] = [];
  for (const t of collect(root, W_NAMESPACE, ['t'])) {
    const afterStart = anchor.rangeStart.compareDocumentPosition(t) & Node.DOCUMENT_POSITION_FOLLOWING;
    const beforeEnd = anchor.rangeEnd.compareDocumentPosition(t) & Node.DOCUMENT_POSITION_PRECEDING;
    if (afterStart && beforeEnd) parts.push(t.textContent ?? '');
  }
  return parts.join('');
}

/**
 * Assemble threads: each root comment with every reply beneath it, transitively.
 *
 * Returns **null** when threading is unknown, rather than a list in which every comment
 * looks like a root. A flat list of roots is what a document with no commentsExtended.xml
 * would produce, and it is indistinguishable from a document that genuinely has no
 * replies — which is exactly the confusion this module exists to prevent.
 */
export function commentThreads(index: CommentIndex): CommentThread[] | null {
  if (!index.threadingKnown) return null;

  const threads: CommentThread[] = [];
  const threadByRootId = new Map<string, CommentThread>();

  for (const comment of index.comments) {
    if (comment.thread.isReply) continue;
    const thread: CommentThread = { root: comment, replies: [], resolved: comment.thread.resolved };
    threads.push(thread);
    threadByRootId.set(comment.id, thread);
  }

  const rootOfComment = (comment: Comment): Comment | null => {
    let current = comment;
    // paraIdParent may name the immediate predecessor rather than the thread root, so
    // walk up. The guard bounds a cycle in malformed input.
    for (let hops = 0; hops <= index.comments.length; hops += 1) {
      if (!current.thread.isReply) return current;
      const parentId = current.thread.parentId;
      if (parentId === null) return null;
      const parent = index.byId.get(parentId);
      if (!parent) return null;
      current = parent;
    }
    return null;
  };

  for (const comment of index.comments) {
    if (!comment.thread.isReply) continue;
    const root = rootOfComment(comment);
    if (!root) continue; // dangling parent: already reported, not silently re-rooted
    threadByRootId.get(root.id)?.replies.push(comment);
  }

  return threads;
}

/**
 * Evidence lines for the AI panel.
 *
 * Needs the three side-car parts alongside the body, and says so when they are absent
 * rather than reporting a threaded discussion as a flat list of unrelated comments —
 * which is exactly what a missing `commentsExtended.xml` looks like from `comments.xml`.
 */
export function computeCommentEvidenceForMarkup(
  parts: Record<string, string>,
  rawXml: string
): { lines: string[]; unresolved: string[] } | null {
  const parse = (xml: string | undefined): Document | null => {
    if (xml === undefined) return null;
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
  };

  const bodyPath = Object.keys(parts).find(p => /^word\/(?:document\d*|header[^/]*|footer[^/]*)\.xml$/.test(p));
  if (bodyPath === undefined) return null;
  const document = parse(parts[bodyPath]);
  if (!document) return null;

  const index = readComments({
    document,
    comments: parse(parts[COMMENT_PART_PATHS.comments]),
    commentsExtended: parse(parts[COMMENT_PART_PATHS.extended]),
    commentsIds: parse(parts['word/commentsIds.xml'])
  });

  if (index.comments.length === 0 && index.anchors.length === 0 && index.problems.length === 0) return null;

  const lines: string[] = [];
  const unresolved: string[] = [];

  lines.push(
    `${bodyPath} anchors ${index.anchors.length} comment reference(s); word/comments.xml supplies ${index.comments.length} comment body/bodies.`
  );

  if (index.threadingKnown) {
    const threads = commentThreads(index);
    const replies = index.comments.filter(c => c.thread.isReply).length;
    const resolved = index.comments.filter(c => c.thread.resolved === true).length;
    lines.push(
      `${threads?.length ?? 0} thread(s): ${replies} of the comments are replies, and ${resolved} are marked resolved.`
    );
  } else {
    // The distinction the module exists to protect: unknown is not "no".
    unresolved.push(
      `${COMMENT_PART_PATHS.extended} is absent, so whether any comment is a reply, and whether any is resolved, cannot be determined. They are not known to be top-level or unresolved — they are unknown.`
    );
  }

  // The selected comment, when the user has one open.
  const selectedId = /<w:comment(?:RangeStart|RangeEnd|Reference)?\b[^>]*w:id="([^"]*)"/.exec(rawXml)?.[1];
  const anchor = selectedId ? index.anchorsById.get(selectedId) : undefined;
  if (anchor) {
    const covered = commentRangeText(anchor);
    const comment = index.byId.get(selectedId!);
    lines.push(
      covered === null
        ? `Comment ${selectedId} has no resolvable range, so what it is attached to cannot be determined from the markup.`
        : covered === ''
          ? `Comment ${selectedId} is anchored to a point rather than a span of text.`
          : `Comment ${selectedId} covers: "${covered.slice(0, 200)}${covered.length > 200 ? '…' : ''}".`
    );
    if (comment) {
      lines.push(`Comment ${selectedId} was written by ${comment.author || 'an unnamed author'}.`);
    }
  }

  lines.push(...renderFindings(index.problems));
  return { lines, unresolved };
}
