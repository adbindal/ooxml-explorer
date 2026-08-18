import { describe, it, expect } from 'vitest';
import {
  readComments,
  commentRangeText,
  commentThreads,
  W14_NAMESPACE,
  W15_NAMESPACE,
  W16CID_NAMESPACE,
  MAX_COMMENT_AUTHOR_LENGTH,
  MAX_COMMENT_INITIALS_LENGTH
} from '../services/wordComments';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W = `xmlns:w="${W_NS}"`;
const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const W15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
const W16CID = 'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"';

const parse = (xml: string) => new DOMParser().parseFromString(xml, 'application/xml');

const doc = (body: string) => parse(`<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`);

const commentsPart = (inner: string) =>
  parse(`<?xml version="1.0"?><w:comments ${W} ${W14}>${inner}</w:comments>`);

const extendedPart = (inner: string) =>
  parse(`<?xml version="1.0"?><w15:commentsEx ${W15}>${inner}</w15:commentsEx>`);

const idsPart = (inner: string) =>
  parse(`<?xml version="1.0"?><w16cid:commentsIds ${W16CID}>${inner}</w16cid:commentsIds>`);

const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const ref = (id: string) => `<w:r><w:commentReference w:id="${id}"/></w:r>`;

/** A body-side comment: range around `text`, closed, with a reference run. */
const anchored = (id: string, text: string) =>
  `<w:commentRangeStart w:id="${id}"/>${run(text)}<w:commentRangeEnd w:id="${id}"/>${ref(id)}`;

/** A comments.xml entry whose single (and therefore last) paragraph carries a paraId. */
const body = (id: string, text: string, paraId: string, author = 'Ada') =>
  `<w:comment w:id="${id}" w:author="${author}" w:initials="A" w:date="2026-01-01T00:00:00Z">` +
  `<w:p w14:paraId="${paraId}">${run(text)}</w:p></w:comment>`;

const kinds = (index: { problems: { kind: string }[] }) => index.problems.map(p => p.kind);

describe('namespace constants match the published schema data', () => {
  it('uses the w14/w15/w16cid URIs the Open XML SDK declares, without a trailing /main', () => {
    // Verified against dotnet/Open-XML-SDK data/namespaces.json. The ECMA namespaces end
    // in /main and these do not, which is the easiest way to write a resolver that finds
    // nothing and reports a document with no threading as a document with no replies.
    expect(W14_NAMESPACE).toBe('http://schemas.microsoft.com/office/word/2010/wordml');
    expect(W15_NAMESPACE).toBe('http://schemas.microsoft.com/office/word/2012/wordml');
    expect(W16CID_NAMESPACE).toBe('http://schemas.microsoft.com/office/word/2016/wordml/cid');
  });
});

describe('readComments — joining the parts', () => {
  it('joins body range, comment text and threading into one record', () => {
    const index = readComments({
      document: doc(`<w:p>${anchored('1', 'commented')}</w:p>`),
      comments: commentsPart(body('1', 'Please check this', '11111111')),
      commentsExtended: extendedPart(`<w15:commentEx w15:paraId="11111111" w15:done="0"/>`)
    });

    expect(index.problems).toEqual([]);
    expect(index.comments).toHaveLength(1);
    const [comment] = index.comments;
    expect(comment.author).toBe('Ada');
    expect(comment.initials).toBe('A');
    expect(comment.date).toBe('2026-01-01T00:00:00Z');
    expect(comment.text).toBe('Please check this');
    expect(comment.paraId).toBe('11111111');
    expect(comment.anchor?.rangeEnd).not.toBeNull();
    expect(comment.thread).toEqual({
      known: true,
      isReply: false,
      parentParaId: null,
      parentId: null,
      resolved: false
    });
  });

  it('finds range markers that sit outside runs', () => {
    // commentRangeStart/End are siblings of w:r, not children. Walking runs alone finds
    // the commentReference and misses both ends of the highlight.
    const index = readComments({
      document: doc(`<w:p>${run('before')}<w:commentRangeStart w:id="4"/>${run('inside')}<w:commentRangeEnd w:id="4"/>${ref('4')}</w:p>`),
      comments: commentsPart(body('4', 'note', 'aaaa0001')),
      commentsExtended: extendedPart(`<w15:commentEx w15:paraId="aaaa0001"/>`)
    });

    expect(index.anchors).toHaveLength(1);
    expect(index.anchors[0].rangeStart).not.toBeNull();
    expect(index.anchors[0].rangeEnd).not.toBeNull();
    expect(index.anchors[0].reference).not.toBeNull();
  });

  it('keys threading on the LAST paragraph of a multi-paragraph comment', () => {
    // The join runs through the paraId of the final paragraph. Taking the first one
    // silently fails to match and reports a threaded comment as unthreaded.
    const multi =
      `<w:comment w:id="1" w:author="Ada"><w:p w14:paraId="AAAA0001">${run('first')}</w:p>` +
      `<w:p w14:paraId="BBBB0002">${run('last')}</w:p></w:comment>`;
    const index = readComments({
      document: doc(`<w:p>${anchored('1', 'x')}</w:p>`),
      comments: commentsPart(multi),
      commentsExtended: extendedPart(`<w15:commentEx w15:paraId="BBBB0002" w15:done="1"/>`)
    });

    expect(index.comments[0].paraId).toBe('BBBB0002');
    expect(index.comments[0].thread.resolved).toBe(true);
    expect(index.problems).toEqual([]);
  });

  it('reads the durable id from commentsIds.xml, keyed on the same paraId', () => {
    const index = readComments({
      document: doc(`<w:p>${anchored('1', 'x')}</w:p>`),
      comments: commentsPart(body('1', 'note', 'CAFE0001')),
      commentsExtended: extendedPart(`<w15:commentEx w15:paraId="CAFE0001"/>`),
      commentsIds: idsPart(`<w16cid:commentId w16cid:paraId="CAFE0001" w16cid:durableId="7A3F0011"/>`)
    });

    expect(index.comments[0].durableId).toBe('7A3F0011');
  });

  it('leaves the durable id null when commentsIds.xml is absent', () => {
    const index = readComments({
      document: doc(`<w:p>${anchored('1', 'x')}</w:p>`),
      comments: commentsPart(body('1', 'note', 'CAFE0001')),
      commentsExtended: extendedPart(`<w15:commentEx w15:paraId="CAFE0001"/>`)
    });

    expect(index.comments[0].durableId).toBeNull();
  });
});

describe('readComments — anchors that break silently', () => {
  it('reports a range start with no end', () => {
    const index = readComments({
      document: doc(`<w:p><w:commentRangeStart w:id="7"/>${run('text')}${ref('7')}</w:p>`),
      comments: commentsPart(body('7', 'note', 'dead0001')),
      commentsExtended: extendedPart(`<w15:commentEx w15:paraId="dead0001"/>`)
    });

    const problem = index.problems.find(p => p.kind === 'unmatched-range-start');
    expect(problem?.id).toBe('7');
    expect(problem?.message).toContain('never closes');
    expect(index.anchors[0].rangeEnd).toBeNull();
  });

  it('reports a stray range end, and admits what cannot be recovered', () => {
    const index = readComments({
      document: doc(`<w:p>${run('text')}<w:commentRangeEnd w:id="9"/></w:p>`)
    });

    const problem = index.problems.find(p => p.kind === 'unmatched-range-end');
    expect(problem?.id).toBe('9');
    expect(problem?.message).toContain('never opened');
  });

  it('does not let one end close two starts sharing an id', () => {
    const index = readComments({
      document: doc(
        `<w:p><w:commentRangeStart w:id="5"/><w:commentRangeStart w:id="5"/>${run('x')}` +
          `<w:commentRangeEnd w:id="5"/>${ref('5')}</w:p>`
      )
    });

    expect(kinds(index)).toContain('duplicate-id');
    expect(index.anchors).toHaveLength(1);
  });

  it('consumes an end once, so a second end with the same id is the stray one', () => {
    // Two ends and one start: exactly one end did a job. Re-offering an already-claimed
    // end reports the *used* end as stray too, which sends an editor to delete the
    // marker that was holding the highlight together.
    const parsed = doc(
      `<w:p><w:commentRangeStart w:id="1"/>${run('a')}<w:commentRangeEnd w:id="1"/>` +
        `${run('b')}<w:commentRangeEnd w:id="1"/>${ref('1')}</w:p>`
    );
    const index = readComments({ document: parsed });

    expect(index.problems.filter(p => p.kind === 'unmatched-range-end')).toHaveLength(1);
    const ends = Array.from(parsed.getElementsByTagNameNS(W_NS, 'commentRangeEnd'));
    expect(index.anchorsById.get('1')!.rangeEnd).toBe(ends[0]);
  });

  it('reports a range whose end precedes its start', () => {
    const index = readComments({
      document: doc(`<w:p><w:commentRangeEnd w:id="1"/>${run('x')}<w:commentRangeStart w:id="1"/>${ref('1')}</w:p>`)
    });

    expect(kinds(index)).toContain('reversed-range');
  });

  it('reports a range with no commentReference run to anchor it', () => {
    const index = readComments({
      document: doc(`<w:p><w:commentRangeStart w:id="2"/>${run('x')}<w:commentRangeEnd w:id="2"/></w:p>`)
    });

    const problem = index.problems.find(p => p.kind === 'missing-reference');
    expect(problem?.id).toBe('2');
    expect(problem?.message).toContain('not displayed');
  });

  it('rejects -1 but accepts -2 on a range marker, the quirk of the id union type', () => {
    // commentRangeStart is CT_MarkupRange and commentReference is CT_Markup, so both
    // take the same @w:id union as bookmarks: non-negative, or signed and ≤ -2.
    const minusOne = readComments({
      document: doc(`<w:p><w:commentRangeStart w:id="-1"/><w:commentRangeEnd w:id="-1"/>${ref('-1')}</w:p>`)
    });
    const minusTwo = readComments({
      document: doc(`<w:p><w:commentRangeStart w:id="-2"/><w:commentRangeEnd w:id="-2"/>${ref('-2')}</w:p>`)
    });

    expect(kinds(minusOne)).toContain('id-out-of-range');
    expect(kinds(minusTwo)).not.toContain('id-out-of-range');
  });
});

describe('readComments — bodies and anchors existing without each other', () => {
  it('reports an id the body references with no w:comment behind it', () => {
    const index = readComments({
      document: doc(`<w:p>${anchored('3', 'x')}</w:p>`),
      comments: commentsPart(''),
      commentsExtended: extendedPart('')
    });

    const problem = index.problems.find(p => p.kind === 'missing-body');
    expect(problem?.id).toBe('3');
    expect(problem?.message).toContain('no text');
  });

  it('reports an orphan comment that never displays but still ships', () => {
    const index = readComments({
      document: doc(`<w:p>${run('no comments here')}</w:p>`),
      comments: commentsPart(body('1', 'internal: lowball them', 'FEED0001')),
      commentsExtended: extendedPart(`<w15:commentEx w15:paraId="FEED0001"/>`)
    });

    const problem = index.problems.find(p => p.kind === 'orphan-comment');
    expect(problem?.id).toBe('1');
    expect(problem?.message).toContain('unzips');
    expect(index.comments[0].anchor).toBeNull();
  });

  it('reports a missing comments.xml once at the part level, not per id', () => {
    const index = readComments({
      document: doc(`<w:p>${anchored('1', 'a')}${anchored('2', 'b')}</w:p>`)
    });

    expect(index.problems.filter(p => p.kind === 'comments-part-missing')).toHaveLength(1);
    expect(kinds(index)).not.toContain('missing-body');
  });

  it('flags two w:comment elements sharing an id', () => {
    const index = readComments({
      document: doc(`<w:p>${anchored('1', 'x')}</w:p>`),
      comments: commentsPart(body('1', 'first', 'AAAA0001') + body('1', 'second', 'BBBB0002')),
      commentsExtended: extendedPart(
        `<w15:commentEx w15:paraId="AAAA0001"/><w15:commentEx w15:paraId="BBBB0002"/>`
      )
    });

    expect(kinds(index)).toContain('duplicate-id');
  });

  it('flags a comment with no author, and one over the length caps', () => {
    const noAuthor = readComments({
      document: doc(`<w:p>${anchored('1', 'x')}</w:p>`),
      comments: commentsPart(`<w:comment w:id="1"><w:p w14:paraId="AAAA0001">${run('t')}</w:p></w:comment>`),
      commentsExtended: extendedPart(`<w15:commentEx w15:paraId="AAAA0001"/>`)
    });
    const tooLong = readComments({
      document: doc(`<w:p>${anchored('1', 'x')}</w:p>`),
      comments: commentsPart(
        `<w:comment w:id="1" w:author="${'A'.repeat(MAX_COMMENT_AUTHOR_LENGTH + 1)}" ` +
          `w:initials="${'I'.repeat(MAX_COMMENT_INITIALS_LENGTH + 1)}">` +
          `<w:p w14:paraId="AAAA0001">${run('t')}</w:p></w:comment>`
      ),
      commentsExtended: extendedPart(`<w15:commentEx w15:paraId="AAAA0001"/>`)
    });

    expect(noAuthor.comments[0].author).toBeNull();
    expect(kinds(noAuthor)).toContain('missing-author');
    expect(kinds(tooLong)).toContain('author-too-long');
    expect(kinds(tooLong)).toContain('initials-too-long');
  });

  it('excludes deleted text and field instructions from the comment body text', () => {
    // Comment bodies carry tracked changes and fields like anything else. A reviewer
    // reading "Approved gone" where the author wrote "Approved" and someone else struck
    // out "gone" is being shown a retracted sentence as if it still stood.
    const withRevisions =
      `<w:comment w:id="1" w:author="Ada"><w:p w14:paraId="AAAA0001">` +
      `${run('Approved')}<w:del><w:r><w:delText> pending legal</w:delText></w:r></w:del>` +
      `<w:r><w:instrText> DATE </w:instrText></w:r></w:p></w:comment>`;
    const index = readComments({
      document: doc(`<w:p>${anchored('1', 'x')}</w:p>`),
      comments: commentsPart(withRevisions),
      commentsExtended: extendedPart(`<w15:commentEx w15:paraId="AAAA0001"/>`)
    });

    expect(index.comments[0].text).toBe('Approved');
  });

  it('accepts an author and initials exactly at the caps', () => {
    const index = readComments({
      document: doc(`<w:p>${anchored('1', 'x')}</w:p>`),
      comments: commentsPart(
        `<w:comment w:id="1" w:author="${'A'.repeat(MAX_COMMENT_AUTHOR_LENGTH)}" ` +
          `w:initials="${'I'.repeat(MAX_COMMENT_INITIALS_LENGTH)}">` +
          `<w:p w14:paraId="AAAA0001">${run('t')}</w:p></w:comment>`
      ),
      commentsExtended: extendedPart(`<w15:commentEx w15:paraId="AAAA0001"/>`)
    });

    expect(index.problems).toEqual([]);
  });
});

describe('commentRangeText — what does comment N cover?', () => {
  it('collects text across runs between the markers', () => {
    const index = readComments({
      document: doc(
        `<w:p>${run('skip')}<w:commentRangeStart w:id="1"/>${run('one ')}${run('two')}` +
          `<w:commentRangeEnd w:id="1"/>${ref('1')}${run('skip')}</w:p>`
      )
    });

    expect(commentRangeText(index.anchorsById.get('1')!)).toBe('one two');
  });

  it('collects text across paragraph and table-cell boundaries', () => {
    // The range markers are independent, so a comment can start mid-paragraph and end
    // inside a table cell. Anything that stops at the parent paragraph truncates.
    const index = readComments({
      document: doc(
        `<w:p><w:commentRangeStart w:id="1"/>${run('para ')}</w:p>` +
          `<w:tbl><w:tr><w:tc><w:p>${run('cell')}<w:commentRangeEnd w:id="1"/>${ref('1')}</w:p></w:tc></w:tr></w:tbl>`
      )
    });

    expect(commentRangeText(index.anchorsById.get('1')!)).toBe('para cell');
  });

  it("distinguishes a range that never closed from one covering nothing", () => {
    // '' is a real answer: a comment on an empty paragraph or on a picture is anchored
    // to a point. null means no answer exists. Collapsing the two turns "deliberately
    // anchored to a point" into "broken", and vice versa.
    const point = readComments({
      document: doc(`<w:p>${run('x')}<w:commentRangeStart w:id="1"/><w:commentRangeEnd w:id="1"/>${ref('1')}</w:p>`)
    });
    const broken = readComments({
      document: doc(`<w:p><w:commentRangeStart w:id="2"/>${run('x')}${ref('2')}</w:p>`)
    });

    expect(commentRangeText(point.anchorsById.get('1')!)).toBe('');
    expect(commentRangeText(broken.anchorsById.get('2')!)).toBeNull();
  });

  it('returns null for a reference with no range at all, and for no anchor', () => {
    const index = readComments({ document: doc(`<w:p>${run('x')}${ref('1')}</w:p>`) });

    expect(commentRangeText(index.anchorsById.get('1')!)).toBeNull();
    expect(commentRangeText(null)).toBeNull();
  });

  it('excludes deleted text and field instructions', () => {
    const index = readComments({
      document: doc(
        `<w:p><w:commentRangeStart w:id="1"/>` +
          `<w:r><w:instrText> PAGE </w:instrText></w:r>` +
          `<w:del><w:r><w:delText>gone</w:delText></w:r></w:del>` +
          `${run('kept')}<w:commentRangeEnd w:id="1"/>${ref('1')}</w:p>`
      )
    });

    expect(commentRangeText(index.anchorsById.get('1')!)).toBe('kept');
  });

  it('surfaces a nested range without calling it an error', () => {
    const index = readComments({
      document: doc(
        `<w:p><w:commentRangeStart w:id="1"/>${run('outer ')}<w:commentRangeStart w:id="2"/>${run('inner')}` +
          `<w:commentRangeEnd w:id="2"/>${ref('2')}<w:commentRangeEnd w:id="1"/>${ref('1')}</w:p>`
      )
    });

    const nested = index.problems.find(p => p.kind === 'nested-range');
    expect(nested?.id).toBe('2');
    expect(nested?.remediation).toContain('No action needed');
    expect(commentRangeText(index.anchorsById.get('1')!)).toBe('outer inner');
    expect(commentRangeText(index.anchorsById.get('2')!)).toBe('inner');
  });

  it('distinguishes partial overlap from nesting', () => {
    const index = readComments({
      document: doc(
        `<w:p><w:commentRangeStart w:id="1"/>${run('a')}<w:commentRangeStart w:id="2"/>${run('b')}` +
          `<w:commentRangeEnd w:id="1"/>${ref('1')}${run('c')}<w:commentRangeEnd w:id="2"/>${ref('2')}</w:p>`
      )
    });

    expect(kinds(index)).toContain('overlapping-range');
    expect(kinds(index)).not.toContain('nested-range');
  });

  it('does not report separate ranges as overlapping', () => {
    const index = readComments({
      document: doc(`<w:p>${anchored('1', 'a')}${run(' ')}${anchored('2', 'b')}</w:p>`)
    });

    expect(kinds(index)).not.toContain('overlapping-range');
    expect(kinds(index)).not.toContain('nested-range');
  });
});

describe('threading — unknown is not the same as false', () => {
  it('reports reply and resolved state as unknown when commentsExtended.xml is absent', () => {
    // The whole point of the module. Defaulting to "not a reply, not resolved" tells a
    // reader that a resolved thread is still open and that a reply is a new topic.
    const index = readComments({
      document: doc(`<w:p>${anchored('1', 'x')}</w:p>`),
      comments: commentsPart(body('1', 'note', 'AAAA0001'))
    });

    expect(index.threadingKnown).toBe(false);
    expect(index.comments[0].thread.known).toBe(false);
    expect(index.comments[0].thread.isReply).toBeNull();
    expect(index.comments[0].thread.resolved).toBeNull();
    const problem = index.problems.find(p => p.kind === 'threading-unknown');
    expect(problem?.message).toContain('commentsExtended.xml was not supplied');
  });

  it('reports unknown for a comment the side-car has no entry for', () => {
    const index = readComments({
      document: doc(`<w:p>${anchored('1', 'a')}${anchored('2', 'b')}</w:p>`),
      comments: commentsPart(body('1', 'first', 'AAAA0001') + body('2', 'second', 'BBBB0002')),
      commentsExtended: extendedPart(`<w15:commentEx w15:paraId="AAAA0001"/>`)
    });

    expect(index.byId.get('1')!.thread.known).toBe(true);
    expect(index.byId.get('2')!.thread.known).toBe(false);
    expect(index.byId.get('2')!.thread.resolved).toBeNull();
    expect(index.threadingKnown).toBe(false);
    expect(index.problems.find(p => p.kind === 'threading-unknown')?.id).toBe('2');
  });

  it('reports unknown when the last paragraph carries no w14:paraId to join on', () => {
    const index = readComments({
      document: doc(`<w:p>${anchored('1', 'x')}</w:p>`),
      comments: commentsPart(`<w:comment w:id="1" w:author="Ada"><w:p>${run('note')}</w:p></w:comment>`),
      commentsExtended: extendedPart(`<w15:commentEx w15:paraId="AAAA0001"/>`)
    });

    expect(index.comments[0].paraId).toBeNull();
    expect(index.comments[0].thread.known).toBe(false);
    expect(kinds(index)).toContain('missing-para-id');
    // The side-car entry now matches nothing either, and that is worth saying too.
    expect(kinds(index)).toContain('orphan-comment-ex');
  });

  it('resolves a reply to the comment it answers, by paraId not by comment id', () => {
    const index = readComments({
      document: doc(`<w:p>${anchored('1', 'x')}${ref('2')}</w:p>`),
      comments: commentsPart(body('1', 'question', 'AAAA0001') + body('2', 'answer', 'BBBB0002', 'Grace')),
      commentsExtended: extendedPart(
        `<w15:commentEx w15:paraId="AAAA0001" w15:done="1"/>` +
          `<w15:commentEx w15:paraId="BBBB0002" w15:paraIdParent="AAAA0001"/>`
      )
    });

    expect(index.byId.get('2')!.thread.isReply).toBe(true);
    expect(index.byId.get('2')!.thread.parentParaId).toBe('AAAA0001');
    expect(index.byId.get('2')!.thread.parentId).toBe('1');
    expect(index.byId.get('1')!.thread.resolved).toBe(true);
    expect(index.threadingKnown).toBe(true);
  });

  it('reads w15:done as an ST_OnOff value, not as mere presence', () => {
    const parsed = (done: string) =>
      readComments({
        document: doc(`<w:p>${anchored('1', 'x')}</w:p>`),
        comments: commentsPart(body('1', 'note', 'AAAA0001')),
        commentsExtended: extendedPart(`<w15:commentEx w15:paraId="AAAA0001" w15:done="${done}"/>`)
      }).comments[0].thread.resolved;

    expect(parsed('1')).toBe(true);
    expect(parsed('true')).toBe(true);
    expect(parsed('0')).toBe(false);
    expect(parsed('false')).toBe(false);
  });

  it('treats an absent w15:done as not resolved, because the entry itself exists', () => {
    const index = readComments({
      document: doc(`<w:p>${anchored('1', 'x')}</w:p>`),
      comments: commentsPart(body('1', 'note', 'AAAA0001')),
      commentsExtended: extendedPart(`<w15:commentEx w15:paraId="AAAA0001"/>`)
    });

    expect(index.comments[0].thread.resolved).toBe(false);
    expect(index.comments[0].thread.known).toBe(true);
  });

  it('reports a reply whose parent paraId names no comment', () => {
    const index = readComments({
      document: doc(`<w:p>${anchored('2', 'x')}</w:p>`),
      comments: commentsPart(body('2', 'answer', 'BBBB0002')),
      commentsExtended: extendedPart(`<w15:commentEx w15:paraId="BBBB0002" w15:paraIdParent="DEAD0000"/>`)
    });

    const problem = index.problems.find(p => p.kind === 'dangling-parent');
    expect(problem?.id).toBe('2');
    expect(problem?.paraId).toBe('DEAD0000');
    // Still known to be a reply — we just cannot say to what.
    expect(index.comments[0].thread.isReply).toBe(true);
    expect(index.comments[0].thread.parentId).toBeNull();
  });

  it('reports a commentEx whose paraId matches no comment', () => {
    const index = readComments({
      document: doc(`<w:p>${anchored('1', 'x')}</w:p>`),
      comments: commentsPart(body('1', 'note', 'AAAA0001')),
      commentsExtended: extendedPart(
        `<w15:commentEx w15:paraId="AAAA0001"/><w15:commentEx w15:paraId="9999FFFF"/>`
      )
    });

    expect(index.problems.find(p => p.kind === 'orphan-comment-ex')?.paraId).toBe('9999FFFF');
  });
});

describe('commentThreads', () => {
  it('returns null rather than a flat list of roots when threading is unknown', () => {
    // A list in which every comment is a root is exactly what a threaded document
    // without its side-car produces, and it is indistinguishable from a document with
    // no replies. Returning null forces the caller to say "unknown".
    const index = readComments({
      document: doc(`<w:p>${anchored('1', 'x')}${ref('2')}</w:p>`),
      comments: commentsPart(body('1', 'question', 'AAAA0001') + body('2', 'answer', 'BBBB0002'))
    });

    expect(commentThreads(index)).toBeNull();
  });

  it('groups replies under their root with the root resolved state', () => {
    const index = readComments({
      document: doc(`<w:p>${anchored('1', 'x')}${ref('2')}${ref('3')}</w:p>`),
      comments: commentsPart(
        body('1', 'question', 'AAAA0001') +
          body('2', 'answer', 'BBBB0002', 'Grace') +
          body('3', 'follow-up', 'CCCC0003', 'Ada')
      ),
      commentsExtended: extendedPart(
        `<w15:commentEx w15:paraId="AAAA0001" w15:done="1"/>` +
          `<w15:commentEx w15:paraId="BBBB0002" w15:paraIdParent="AAAA0001"/>` +
          // Points at the previous *reply*, not at the root: both readings of
          // paraIdParent must land the comment in the same thread.
          `<w15:commentEx w15:paraId="CCCC0003" w15:paraIdParent="BBBB0002"/>`
      )
    });

    const threads = commentThreads(index)!;
    expect(threads).toHaveLength(1);
    expect(threads[0].root.id).toBe('1');
    expect(threads[0].replies.map(r => r.id)).toEqual(['2', '3']);
    expect(threads[0].resolved).toBe(true);
  });

  it('drops a reply whose parent is gone instead of promoting it to a root', () => {
    const index = readComments({
      document: doc(`<w:p>${anchored('2', 'x')}</w:p>`),
      comments: commentsPart(body('2', 'answer', 'BBBB0002')),
      commentsExtended: extendedPart(`<w15:commentEx w15:paraId="BBBB0002" w15:paraIdParent="DEAD0000"/>`)
    });

    // Threading is known for this comment, so threads are computable — but a reply with
    // no parent is not a top-level comment, and presenting it as one invents a topic.
    expect(commentThreads(index)).toEqual([]);
    expect(kinds(index)).toContain('dangling-parent');
  });

  it('does not hang on a parent cycle', () => {
    const index = readComments({
      document: doc(`<w:p>${ref('1')}${ref('2')}</w:p>`),
      comments: commentsPart(body('1', 'a', 'AAAA0001') + body('2', 'b', 'BBBB0002')),
      commentsExtended: extendedPart(
        `<w15:commentEx w15:paraId="AAAA0001" w15:paraIdParent="BBBB0002"/>` +
          `<w15:commentEx w15:paraId="BBBB0002" w15:paraIdParent="AAAA0001"/>`
      )
    });

    expect(commentThreads(index)).toEqual([]);
  });
});

describe('malformed input is tolerated', () => {
  it('returns an empty index for a document with no comments', () => {
    const index = readComments({ document: doc(`<w:p>${run('nothing here')}</w:p>`) });

    expect(index.comments).toEqual([]);
    expect(index.anchors).toEqual([]);
    expect(index.problems).toEqual([]);
    expect(index.threadingKnown).toBe(true);
    expect(commentThreads(index)).toEqual([]);
  });

  it('does not throw on markers with no id at all', () => {
    const index = readComments({
      document: doc(`<w:p><w:commentRangeStart/><w:commentRangeEnd/><w:r><w:commentReference/></w:r></w:p>`),
      comments: commentsPart(`<w:comment w:author="Ada"><w:p>${run('t')}</w:p></w:comment>`),
      commentsExtended: extendedPart(`<w15:commentEx/>`)
    });

    expect(index.comments).toEqual([]);
    expect(index.problems.length).toBeGreaterThan(0);
  });

  it('accepts a bare element rather than a Document for every part', () => {
    const documentEl = doc(`<w:p>${anchored('1', 'x')}</w:p>`).documentElement;
    const commentsEl = commentsPart(body('1', 'note', 'AAAA0001')).documentElement;
    const extendedEl = extendedPart(`<w15:commentEx w15:paraId="AAAA0001"/>`).documentElement;

    const index = readComments({
      document: documentEl,
      comments: commentsEl,
      commentsExtended: extendedEl
    });

    expect(index.problems).toEqual([]);
    expect(index.comments[0].thread.known).toBe(true);
  });
});
