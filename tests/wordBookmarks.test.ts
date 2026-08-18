import { describe, it, expect } from 'vitest';
import {
  readBookmarks,
  bookmarkText,
  findMarkupIdCollisions,
  nextSafeMarkupId,
  findBookmarkReferences,
  MAX_BOOKMARK_NAME_LENGTH,
  computeBookmarkEvidenceForMarkup
} from '../services/wordBookmarks';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const doc = (body: string) =>
  new DOMParser().parseFromString(
    `<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`,
    'application/xml'
  );

const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;

describe('readBookmarks — the range model', () => {
  it('pairs a start to an end by id, not by name', () => {
    // Only the start carries a name. If a resolver matched on name it would find
    // nothing to match against, because the end has no name to match.
    const index = readBookmarks(
      doc(`<w:p><w:bookmarkStart w:id="1" w:name="Intro"/>${run('hello')}<w:bookmarkEnd w:id="1"/></w:p>`)
    );

    expect(index.problems).toEqual([]);
    expect(index.bookmarks).toHaveLength(1);
    expect(index.bookmarks[0].name).toBe('Intro');
    expect(index.bookmarks[0].end).not.toBeNull();
    expect(index.byName.get('Intro')).toHaveLength(1);
  });

  it('finds bookmarks that sit outside runs', () => {
    // Bookmark markers are direct children of w:p. Walking runs alone misses them
    // entirely — the bug that made this project's diff report "no changes" for a
    // document that had lost every cross-reference target.
    const index = readBookmarks(
      doc(`<w:p>${run('before')}<w:bookmarkStart w:id="4" w:name="Mid"/>${run('after')}<w:bookmarkEnd w:id="4"/></w:p>`)
    );

    expect(index.bookmarks.map(b => b.name)).toEqual(['Mid']);
  });

  it('reports a start with no end as a bookmark that does not exist', () => {
    const index = readBookmarks(
      doc(`<w:p><w:bookmarkStart w:id="7" w:name="Dangling"/>${run('text')}</w:p>`)
    );

    expect(index.bookmarks[0].end).toBeNull();
    const problem = index.problems.find(p => p.kind === 'unmatched-start');
    expect(problem?.name).toBe('Dangling');
    expect(problem?.message).toContain('resolve to nothing');
    expect(problem?.remediation).toContain('w:bookmarkEnd');
  });

  it('reports a stray end, and admits the lost name is unrecoverable', () => {
    const index = readBookmarks(doc(`<w:p>${run('text')}<w:bookmarkEnd w:id="9"/></w:p>`));

    const problem = index.problems.find(p => p.kind === 'unmatched-end');
    expect(problem?.id).toBe('9');
    // The end carries no name, so we can say a bookmark was lost but not which.
    expect(problem?.message).toContain('no way to tell which');
  });

  it('does not close a bookmarkStart with a moveToRangeEnd that shares its id', () => {
    // Move bookmarks are a separate range kind. Matching on id alone, ignoring kind,
    // would pair these and report a healthy document.
    const index = readBookmarks(
      doc(`<w:p><w:bookmarkStart w:id="2" w:name="Real"/>${run('x')}<w:moveToRangeEnd w:id="2"/></w:p>`)
    );

    expect(index.bookmarks[0].end).toBeNull();
    expect(index.problems.map(p => p.kind).sort()).toEqual(['unmatched-end', 'unmatched-start']);
  });

  it('keeps the kinds separate when a document contains both', () => {
    // A tracked move alongside ordinary bookmarks is routine, and it is the only
    // arrangement where cross-kind matching is observable: here the unclosed
    // bookmark 5 must NOT be closed by the moveToRangeEnd that shares its id.
    const index = readBookmarks(
      doc(
        `<w:p><w:bookmarkStart w:id="1" w:name="Fine"/>${run('a')}<w:bookmarkEnd w:id="1"/>` +
          `<w:bookmarkStart w:id="5" w:name="Broken"/>${run('b')}<w:moveToRangeEnd w:id="5"/></w:p>`
      )
    );

    expect(index.byName.get('Fine')![0].end).not.toBeNull();
    expect(index.byName.get('Broken')![0].end).toBeNull();
    expect(index.problems.find(p => p.kind === 'unmatched-start')?.name).toBe('Broken');
    expect(index.problems.find(p => p.kind === 'unmatched-end')?.id).toBe('5');
  });

  it('treats move bookmarks as their own kind of range', () => {
    const index = readBookmarks(
      doc(
        `<w:p><w:moveToRangeStart w:id="3" w:name="move1"/>${run('moved')}<w:moveToRangeEnd w:id="3"/></w:p>`
      )
    );

    expect(index.problems).toEqual([]);
    expect(index.bookmarks[0].kind).toBe('moveToRangeStart');
  });
});

describe('readBookmarks — what silently loses a range', () => {
  it('flags two starts sharing one id', () => {
    const index = readBookmarks(
      doc(
        `<w:p><w:bookmarkStart w:id="1" w:name="A"/>${run('a')}<w:bookmarkEnd w:id="1"/>` +
          `<w:bookmarkStart w:id="1" w:name="B"/>${run('b')}</w:p>`
      )
    );

    const problem = index.problems.find(p => p.kind === 'duplicate-id');
    expect(problem?.message).toContain('"A"');
    expect(problem?.message).toContain('"B"');
  });

  it('does not let one end close two starts', () => {
    // With a single end available, the second start must be left open rather than
    // both claiming the same element and both looking healthy.
    const index = readBookmarks(
      doc(
        `<w:p><w:bookmarkStart w:id="5" w:name="A"/><w:bookmarkStart w:id="5" w:name="B"/>` +
          `${run('x')}<w:bookmarkEnd w:id="5"/></w:p>`
      )
    );

    const closed = index.bookmarks.filter(b => b.end !== null);
    expect(closed).toHaveLength(1);
  });

  it('flags two bookmarks sharing a name, because references cannot choose', () => {
    const index = readBookmarks(
      doc(
        `<w:p><w:bookmarkStart w:id="1" w:name="Dup"/>${run('a')}<w:bookmarkEnd w:id="1"/>` +
          `<w:bookmarkStart w:id="2" w:name="Dup"/>${run('b')}<w:bookmarkEnd w:id="2"/></w:p>`
      )
    );

    expect(index.problems.filter(p => p.kind === 'duplicate-name')).toHaveLength(1);
    expect(index.byName.get('Dup')).toHaveLength(2);
  });

  it('flags a range whose end precedes its start', () => {
    const index = readBookmarks(
      doc(`<w:p><w:bookmarkEnd w:id="1"/>${run('x')}<w:bookmarkStart w:id="1" w:name="Backwards"/></w:p>`)
    );

    expect(index.problems.find(p => p.kind === 'reversed-range')?.name).toBe('Backwards');
  });

  it('flags a name over the 40-character cap', () => {
    const long = 'B'.repeat(MAX_BOOKMARK_NAME_LENGTH + 1);
    const index = readBookmarks(
      doc(`<w:p><w:bookmarkStart w:id="1" w:name="${long}"/><w:bookmarkEnd w:id="1"/></w:p>`)
    );

    expect(index.problems.find(p => p.kind === 'name-too-long')?.message).toContain('41 characters');
  });

  it('accepts a name exactly at the cap', () => {
    const exact = 'B'.repeat(MAX_BOOKMARK_NAME_LENGTH);
    const index = readBookmarks(
      doc(`<w:p><w:bookmarkStart w:id="1" w:name="${exact}"/><w:bookmarkEnd w:id="1"/></w:p>`)
    );

    expect(index.problems).toEqual([]);
  });

  it('rejects -1 but accepts -2, the quirk of the id union type', () => {
    // @w:id is a non-negative number OR a signed number of -2 or below, so -1 is the
    // single integer the union excludes.
    const minusOne = readBookmarks(
      doc(`<w:p><w:bookmarkStart w:id="-1" w:name="A"/><w:bookmarkEnd w:id="-1"/></w:p>`)
    );
    const minusTwo = readBookmarks(
      doc(`<w:p><w:bookmarkStart w:id="-2" w:name="A"/><w:bookmarkEnd w:id="-2"/></w:p>`)
    );

    expect(minusOne.problems.map(p => p.kind)).toContain('id-out-of-range');
    expect(minusTwo.problems).toEqual([]);
  });
});

describe('readBookmarks — classification', () => {
  it('marks underscore-prefixed bookmarks hidden and names what generated them', () => {
    const index = readBookmarks(
      doc(
        `<w:p><w:bookmarkStart w:id="1" w:name="_Toc12345"/><w:bookmarkEnd w:id="1"/>` +
          `<w:bookmarkStart w:id="2" w:name="_GoBack"/><w:bookmarkEnd w:id="2"/>` +
          `<w:bookmarkStart w:id="3" w:name="Visible"/><w:bookmarkEnd w:id="3"/></w:p>`
      )
    );

    const [toc, goBack, visible] = index.bookmarks;
    expect(toc.hidden).toBe(true);
    expect(toc.generatedBy).toContain('table-of-contents');
    expect(goBack.generatedBy).toContain('last edit position');
    expect(visible.hidden).toBe(false);
    expect(visible.generatedBy).toBeNull();
  });

  it('records column ranges, which are how a bookmark covers table columns', () => {
    const index = readBookmarks(
      doc(
        `<w:tbl><w:tr><w:tc><w:p><w:bookmarkStart w:id="1" w:name="Cols" w:colFirst="0" w:colLast="2"/>` +
          `<w:bookmarkEnd w:id="1"/></w:p></w:tc></w:tr></w:tbl>`
      )
    );

    expect(index.bookmarks[0].isColumnRange).toBe(true);
    expect(index.bookmarks[0].colFirst).toBe('0');
    expect(index.bookmarks[0].colLast).toBe('2');
  });
});

describe('bookmarkText', () => {
  it('collects text across runs between the markers', () => {
    const index = readBookmarks(
      doc(
        `<w:p>${run('skip')}<w:bookmarkStart w:id="1" w:name="Span"/>${run('one ')}${run('two')}` +
          `<w:bookmarkEnd w:id="1"/>${run('skip')}</w:p>`
      )
    );

    expect(bookmarkText(index.bookmarks[0])).toBe('one two');
  });

  it('collects text across paragraph and table boundaries', () => {
    // Start and end are independent markers, so a range may begin in a paragraph and
    // end inside a table cell. Anything that stops at the parent paragraph truncates.
    const index = readBookmarks(
      doc(
        `<w:p><w:bookmarkStart w:id="1" w:name="Wide"/>${run('para ')}</w:p>` +
          `<w:tbl><w:tr><w:tc><w:p>${run('cell')}<w:bookmarkEnd w:id="1"/></w:p></w:tc></w:tr></w:tbl>`
      )
    );

    expect(bookmarkText(index.bookmarks[0])).toBe('para cell');
  });

  it('distinguishes an unclosed bookmark from one covering nothing', () => {
    // '' is a real answer: a cross-reference target is normally an insertion point
    // with no content. null means the range never closed. Collapsing the two loses
    // the difference between "empty on purpose" and "broken".
    const empty = readBookmarks(
      doc(`<w:p>${run('x')}<w:bookmarkStart w:id="1" w:name="Point"/><w:bookmarkEnd w:id="1"/></w:p>`)
    );
    const broken = readBookmarks(doc(`<w:p><w:bookmarkStart w:id="2" w:name="Open"/>${run('x')}</w:p>`));

    expect(bookmarkText(empty.bookmarks[0])).toBe('');
    expect(bookmarkText(broken.bookmarks[0])).toBeNull();
  });

  it('excludes deleted text and field instructions', () => {
    const index = readBookmarks(
      doc(
        `<w:p><w:bookmarkStart w:id="1" w:name="B"/>` +
          `<w:r><w:instrText> PAGE </w:instrText></w:r>` +
          `<w:del><w:r><w:delText>gone</w:delText></w:r></w:del>` +
          `${run('kept')}<w:bookmarkEnd w:id="1"/></w:p>`
      )
    );

    expect(bookmarkText(index.bookmarks[0])).toBe('kept');
  });
});

describe('findMarkupIdCollisions — the corruption class', () => {
  it('catches a tracked change reusing a bookmark id', () => {
    // The failure this exists for: a generator numbers its revisions from 1, the
    // document already has bookmark 1, Word refuses to open the file, and every
    // lenient reader opens it fine — so it ships.
    const collisions = findMarkupIdCollisions(
      doc(
        `<w:p><w:bookmarkStart w:id="1" w:name="Intro"/>` +
          `<w:ins w:id="1" w:author="gen" w:date="2026-01-01T00:00:00Z">${run('added')}</w:ins>` +
          `<w:bookmarkEnd w:id="1"/></w:p>`
      )
    );

    expect(collisions).toHaveLength(1);
    expect(collisions[0].id).toBe('1');
    expect(collisions[0].elements).toContain('ins');
    expect(collisions[0].elements).toContain('bookmarkStart');
  });

  it('does not flag a document whose ids are disjoint', () => {
    const collisions = findMarkupIdCollisions(
      doc(
        `<w:p><w:bookmarkStart w:id="1" w:name="Intro"/>` +
          `<w:ins w:id="2" w:author="gen" w:date="2026-01-01T00:00:00Z">${run('added')}</w:ins>` +
          `<w:bookmarkEnd w:id="3"/></w:p>`
      )
    );

    expect(collisions).toEqual([]);
  });

  it('spans the whole markup id space, not just bookmarks and revisions', () => {
    const collisions = findMarkupIdCollisions(
      doc(`<w:p><w:permStart w:id="8"/><w:bookmarkStart w:id="8" w:name="X"/></w:p>`)
    );

    expect(collisions[0].elements.sort()).toEqual(['bookmarkStart', 'permStart']);
  });
});

describe('nextSafeMarkupId', () => {
  it('clears every element type, not just bookmarks', () => {
    // Taking max+1 over bookmarks alone returns 6 here and collides with the revision.
    const parsed = doc(
      `<w:p><w:bookmarkStart w:id="5" w:name="A"/>` +
        `<w:ins w:id="90" w:author="g" w:date="2026-01-01T00:00:00Z">${run('x')}</w:ins>` +
        `<w:bookmarkEnd w:id="5"/></w:p>`
    );

    expect(nextSafeMarkupId(parsed)).toBe(91);
  });

  it('starts at 0 for a document with no markup ids', () => {
    expect(nextSafeMarkupId(doc(`<w:p>${run('plain')}</w:p>`))).toBe(0);
  });
});

describe('findBookmarkReferences', () => {
  it('finds hyperlink anchors and field instructions', () => {
    const parsed = doc(
      `<w:p><w:bookmarkStart w:id="1" w:name="Target"/><w:bookmarkEnd w:id="1"/></w:p>` +
        `<w:p><w:hyperlink w:anchor="Target">${run('go')}</w:hyperlink></w:p>` +
        `<w:p><w:r><w:instrText> REF Target \\h </w:instrText></w:r></w:p>` +
        `<w:p><w:fldSimple w:instr=" PAGEREF Target \\h ">${run('3')}</w:fldSimple></w:p>`
    );

    expect(findBookmarkReferences(parsed, 'Target')).toHaveLength(3);
  });

  it('does not match a bookmark name that is a prefix of another', () => {
    // Substring matching reports "Ch1" as referenced by a REF to Ch10, which turns a
    // broken cross-reference into a healthy-looking one.
    const parsed = doc(`<w:p><w:r><w:instrText> REF Ch10 \\h </w:instrText></w:r></w:p>`);

    expect(findBookmarkReferences(parsed, 'Ch1')).toEqual([]);
    expect(findBookmarkReferences(parsed, 'Ch10')).toHaveLength(1);
  });

  it('returns nothing for a bookmark no one points at', () => {
    const parsed = doc(`<w:p><w:bookmarkStart w:id="1" w:name="Lonely"/><w:bookmarkEnd w:id="1"/></w:p>`);

    expect(findBookmarkReferences(parsed, 'Lonely')).toEqual([]);
  });
});

describe('malformed input is tolerated', () => {
  it('does not throw on a start with no id or name', () => {
    const index = readBookmarks(doc(`<w:p><w:bookmarkStart/><w:bookmarkEnd/></w:p>`));

    expect(index.bookmarks).toEqual([]);
    expect(index.problems.length).toBeGreaterThan(0);
  });

  it('reports a missing name against the id that has it', () => {
    const index = readBookmarks(doc(`<w:p><w:bookmarkStart w:id="1"/><w:bookmarkEnd w:id="1"/></w:p>`));

    expect(index.problems.find(p => p.kind === 'missing-name')?.id).toBe('1');
  });

  it('returns an empty index for a document with no bookmarks', () => {
    const index = readBookmarks(doc(`<w:p>${run('nothing here')}</w:p>`));

    expect(index.bookmarks).toEqual([]);
    expect(index.problems).toEqual([]);
  });
});

describe('computeBookmarkEvidenceForMarkup — panel wiring', () => {
  const part = (body: string) => ({
    'word/document.xml': `<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`
  });

  it('returns null when the part has no bookmarks, so the panel degrades quietly', () => {
    expect(computeBookmarkEvidenceForMarkup(part(`<w:p>${run('plain')}</w:p>`), '')).toBeNull();
  });

  it('returns null for a part that carries no bookmarks at all', () => {
    expect(computeBookmarkEvidenceForMarkup({ 'word/styles.xml': '<w:styles/>' }, '')).toBeNull();
  });

  it('counts visible and hidden bookmarks separately', () => {
    const evidence = computeBookmarkEvidenceForMarkup(
      part(
        `<w:p><w:bookmarkStart w:id="1" w:name="Chapter"/><w:bookmarkEnd w:id="1"/>` +
          `<w:bookmarkStart w:id="2" w:name="_Toc99"/><w:bookmarkEnd w:id="2"/></w:p>`
      ),
      ''
    );

    expect(evidence!.lines[0]).toContain('2 bookmark range(s): 1 user-visible, 1 hidden');
  });

  it('reports the covered text of the selected bookmark', () => {
    const evidence = computeBookmarkEvidenceForMarkup(
      part(`<w:p><w:bookmarkStart w:id="1" w:name="Sel"/>${run('covered text')}<w:bookmarkEnd w:id="1"/></w:p>`),
      '<w:bookmarkStart w:id="1" w:name="Sel"/>'
    );

    expect(evidence!.lines.some(l => l.includes('covers: "covered text"'))).toBe(true);
  });

  it('says a selected bookmark is unreferenced without calling it an error', () => {
    const evidence = computeBookmarkEvidenceForMarkup(
      part(`<w:p><w:bookmarkStart w:id="1" w:name="Sel"/>${run('x')}<w:bookmarkEnd w:id="1"/></w:p>`),
      '<w:bookmarkStart w:id="1" w:name="Sel"/>'
    );

    const line = evidence!.lines.find(l => l.includes('Nothing in word/document.xml references'));
    expect(line).toContain('not itself an error');
  });

  it('caps what it claims: cross-part references go to unresolved', () => {
    // Only the open part is read, so a reference from a header cannot be seen. The
    // tier must be capped rather than the absence reported as fact.
    const evidence = computeBookmarkEvidenceForMarkup(
      part(`<w:p><w:bookmarkStart w:id="1" w:name="Sel"/><w:bookmarkEnd w:id="1"/></w:p>`),
      '<w:bookmarkStart w:id="1" w:name="Sel"/>'
    );

    expect(evidence!.unresolved.some(u => u.includes('other parts of the package'))).toBe(true);
  });

  it('surfaces an id collision with the id to renumber from', () => {
    const evidence = computeBookmarkEvidenceForMarkup(
      part(
        `<w:p><w:bookmarkStart w:id="1" w:name="A"/>` +
          `<w:ins w:id="1" w:author="g" w:date="2026-01-01T00:00:00Z">${run('x')}</w:ins>` +
          `<w:bookmarkEnd w:id="1"/></w:p>`
      ),
      ''
    );

    const line = evidence!.lines.find(l => l.includes('is used by'));
    expect(line).toContain('Word rejects');
    expect(line).toContain('Renumber from 2');
  });

  it('reports problems even when nothing is selected', () => {
    const evidence = computeBookmarkEvidenceForMarkup(
      part(`<w:p><w:bookmarkStart w:id="1" w:name="Open"/>${run('x')}</w:p>`),
      ''
    );

    expect(evidence!.lines.some(l => l.includes('opens but never closes'))).toBe(true);
  });

  it('returns null rather than throwing on malformed XML', () => {
    expect(computeBookmarkEvidenceForMarkup({ 'word/document.xml': '<w:document><oops>' }, '')).toBeNull();
  });
});
