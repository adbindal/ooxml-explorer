import { describe, it, expect } from 'vitest';
import {
  readRevisions,
  compareRevisionOutcomes,
  paragraphMarkRevision,
  pairMoves,
  readRevisionSettings,
  checkRevisionVisibility,
  computeRevisionEvidenceForMarkup,
  UNATTRIBUTED_AUTHOR,
  MAX_AUTHOR_LENGTH
} from '../services/wordRevisions';
import { readBookmarks } from '../services/wordBookmarks';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const doc = (body: string) =>
  new DOMParser().parseFromString(
    `<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`,
    'application/xml'
  );

const settingsDoc = (body: string) =>
  new DOMParser().parseFromString(
    `<?xml version="1.0"?><w:settings ${W}>${body}</w:settings>`,
    'application/xml'
  );

/** A tracked-change attribute set that is complete, so tests isolate one fault at a time. */
const meta = (id: string, author = 'Ann') => `w:id="${id}" w:author="${author}" w:date="2024-01-01T00:00:00Z"`;

const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const delRun = (text: string) => `<w:r><w:delText>${text}</w:delText></w:r>`;

const ins = (text: string, id = '1', author = 'Ann') => `<w:ins ${meta(id, author)}>${run(text)}</w:ins>`;
const del = (text: string, id = '2', author = 'Ann') => `<w:del ${meta(id, author)}>${delRun(text)}</w:del>`;

const para = (inner: string, pPr = '') => `<w:p>${pPr}${inner}</w:p>`;

/** `w:pPr/w:rPr/w:ins|w:del` — the paragraph mark itself, not any run in the paragraph. */
const markRevision = (kind: string, id = '9') => `<w:pPr><w:rPr><w:${kind} ${meta(id)}/></w:rPr></w:pPr>`;

describe('readRevisions — what is in the file', () => {
  it('reads an insertion and a deletion, keeping the deleted text', () => {
    const index = readRevisions(doc(para(`${ins('new wording')}${del('old wording')}`)), 'word/document.xml');

    expect(index.revisions.map(r => r.kind)).toEqual(['ins', 'del']);
    expect(index.revisions[0].text).toBe('new wording');
    // The deleted words are still in the file — that is the whole point.
    expect(index.revisions[1].text).toBe('old wording');
    expect(index.revisions.every(r => r.category === 'content')).toBe(true);
  });

  it('counts revisions per author', () => {
    const index = readRevisions(
      doc(para(`${ins('a', '1', 'Ann')}${ins('b', '2', 'Bo')}${del('c', '3', 'Ann')}`))
    );

    expect(index.byAuthor.get('Ann')).toBe(2);
    expect(index.byAuthor.get('Bo')).toBe(1);
  });

  it('attributes a revision with no author rather than dropping it from the counts', () => {
    const index = readRevisions(doc(para(`<w:ins w:id="1" w:date="2024-01-01T00:00:00Z">${run('x')}</w:ins>`)));

    expect(index.byAuthor.get(UNATTRIBUTED_AUTHOR)).toBe(1);
    expect(index.problems.map(p => p.code)).toContain('revision/missing-author');
  });

  it('finds a paragraph-mark revision, which touches no run', () => {
    // Nothing inside any w:r is revised here. Anything that walks runs sees a clean
    // paragraph, and accepting the change merges it with the next one.
    const index = readRevisions(doc(para(run('kept'), markRevision('del'))));

    const mark = index.revisions.find(r => r.category === 'paragraph-mark');
    expect(mark?.kind).toBe('del');
    expect(mark?.text).toBe('');
  });

  it('does not mistake a run revision for a paragraph-mark revision', () => {
    const index = readRevisions(doc(para(del('gone'))));

    expect(index.revisions.map(r => r.category)).toEqual(['content']);
    expect(paragraphMarkRevision(doc(para(del('gone'))).getElementsByTagName('w:p')[0])).toBeNull();
  });

  it('looks only at w:pPr/w:rPr’s own children for the mark, not anywhere in the paragraph', () => {
    // The paragraph HAS a pPr/rPr, so an early "no rPr, no mark" exit cannot decide this
    // one — the revision has to be rejected because of where it sits. Without that, every
    // paragraph containing a deleted run would report its mark as deleted and the
    // accepted reading would lose a paragraph break that nothing deleted.
    const paragraph = doc(para(del('gone'), '<w:pPr><w:rPr><w:b/></w:rPr></w:pPr>')).getElementsByTagName('w:p')[0];

    expect(paragraphMarkRevision(paragraph)).toBeNull();

    const body = para(del('gone'), '<w:pPr><w:rPr><w:b/></w:rPr></w:pPr>') + para(run('next'));
    expect(compareRevisionOutcomes(doc(body)).accepted).toBe('\nnext');
  });

  it('ignores a run wrongly nested inside a paragraph-mark revision', () => {
    // Out of schema: CT_TrackChange is a leaf, so w:pPr/w:rPr/w:del cannot contain runs.
    // Text found there is not paragraph content and must not be scored as deleted.
    const body = para(
      run('body'),
      `<w:pPr><w:rPr><w:del ${meta('9')}>${delRun('stray')}</w:del></w:rPr></w:pPr>`
    );
    const outcome = compareRevisionOutcomes(doc(body));

    expect(outcome.accepted).toBe('body');
    expect(outcome.rejected).toBe('body');
  });

  it('does not judge the spelling of text inside a paragraph mark either', () => {
    // A w:t inside a paragraph-mark deletion is not a deleted run spelled wrongly; it is
    // a run somewhere runs cannot go. Reporting it as live-text-in-deletion would hand
    // the reader a fix ("use w:delText") that leaves the real fault in place.
    const body = para(run('body'), `<w:pPr><w:rPr><w:del ${meta('9')}>${run('stray')}</w:del></w:rPr></w:pPr>`);

    expect(readRevisions(doc(body)).problems.map(p => p.code)).not.toContain('revision/live-text-in-deletion');
  });

  it('treats run properties as properties too, not only paragraph properties', () => {
    // Out of schema in the other direction: CT_RPr admits no revision elements at all.
    // Without an w:rPr guard this text has no w:pPr above it and would be scored as
    // inserted content, putting "stray" into the accepted reading.
    const body = para(
      `<w:r><w:rPr><w:ins ${meta('9')}>${run('stray')}</w:ins></w:rPr><w:t>real</w:t></w:r>`
    );

    expect(compareRevisionOutcomes(doc(body)).accepted).toBe('real');
  });

  it('classifies format-only revisions, which change no character', () => {
    const index = readRevisions(
      doc(
        para(run('text'), `<w:pPr><w:pPrChange ${meta('4')}><w:pPr/></w:pPrChange></w:pPr>`) +
          `<w:tbl><w:tblPr><w:tblPrChange ${meta('5')}><w:tblPr/></w:tblPrChange></w:tblPr></w:tbl>`
      )
    );

    expect(index.revisions.filter(r => r.category === 'format').map(r => r.kind).sort()).toEqual([
      'pPrChange',
      'tblPrChange'
    ]);
  });

  it('classifies row, cell and numbering revisions by their parent, not their name', () => {
    // w:ins means four different things depending on where it sits.
    const index = readRevisions(
      doc(
        `<w:tbl><w:tr><w:trPr><w:ins ${meta('1')}/></w:trPr>` +
          `<w:tc><w:tcPr><w:cellIns ${meta('2')}/></w:tcPr></w:tc></w:tr></w:tbl>` +
          para(run('x'), `<w:pPr><w:numPr><w:ins ${meta('3')}/></w:numPr></w:pPr>`)
      )
    );

    const byCategory = Object.fromEntries(index.revisions.map(r => [r.category, r.kind]));
    expect(byCategory.row).toBe('ins');
    expect(byCategory.cell).toBe('cellIns');
    expect(byCategory.numbering).toBe('ins');
  });

  it('finds nested revisions — text inserted by one author and deleted by another', () => {
    const index = readRevisions(
      doc(para(`<w:del ${meta('2', 'Bo')}><w:ins ${meta('1', 'Ann')}>${delRun('draft')}</w:ins></w:del>`))
    );

    expect(index.revisions.map(r => r.kind).sort()).toEqual(['del', 'ins']);
  });

  it('reports nothing on a document with no revisions at all', () => {
    const index = readRevisions(doc(para(run('plain text'))));

    expect(index.revisions).toEqual([]);
    expect(index.problems).toEqual([]);
  });
});

describe('compareRevisionOutcomes — the two texts', () => {
  it('gives a different text for accept and for reject', () => {
    const outcome = compareRevisionOutcomes(doc(para(`${run('The ')}${ins('new')}${del('old')}${run(' way')}`)));

    expect(outcome.accepted).toBe('The new way');
    expect(outcome.rejected).toBe('The old way');
    expect(outcome.differs).toBe(true);
  });

  it('produces a naive reading that matches neither, which is what extractors emit', () => {
    const outcome = compareRevisionOutcomes(doc(para(`${run('The ')}${ins('new')}${del('old')}${run(' way')}`)));

    expect(outcome.naive).toBe('The newold way');
    expect(outcome.naive).not.toBe(outcome.accepted);
    expect(outcome.naive).not.toBe(outcome.rejected);
  });

  it('treats a move as removing content from one place and adding it in another', () => {
    const body =
      para(`${run('A ')}<w:moveFrom ${meta('1')}>${delRun('B')}</w:moveFrom>${run(' C')}`) +
      para(`<w:moveTo ${meta('2')}>${run('B')}</w:moveTo>`);
    const outcome = compareRevisionOutcomes(doc(body));

    expect(outcome.accepted).toBe('A  C\nB');
    expect(outcome.rejected).toBe('A B C\n');
  });

  it('merges paragraphs when a deleted paragraph mark is accepted', () => {
    // No run is revised. The only change is to the mark, and it changes the structure.
    const body = para(run('first'), markRevision('del')) + para(run('second'));
    const outcome = compareRevisionOutcomes(doc(body));

    expect(outcome.accepted).toBe('firstsecond');
    expect(outcome.rejected).toBe('first\nsecond');
    expect(outcome.differs).toBe(true);
  });

  it('merges paragraphs the other way round when an inserted mark is rejected', () => {
    const body = para(run('first'), markRevision('ins')) + para(run('second'));
    const outcome = compareRevisionOutcomes(doc(body));

    expect(outcome.accepted).toBe('first\nsecond');
    expect(outcome.rejected).toBe('firstsecond');
  });

  it('drops text that was inserted and then deleted from both readings', () => {
    // It is in neither document: the original never had it, the final one does not either.
    const body = para(
      `${run('keep ')}<w:del ${meta('2', 'Bo')}><w:ins ${meta('1', 'Ann')}>${delRun('churn')}</w:ins></w:del>${run(' end')}`
    );
    const outcome = compareRevisionOutcomes(doc(body));

    expect(outcome.accepted).toBe('keep  end');
    expect(outcome.rejected).toBe('keep  end');
    expect(outcome.differs).toBe(false);
  });

  it('reports no difference when every revision is format-only', () => {
    const body = para(run('unchanged'), `<w:pPr><w:pPrChange ${meta('4')}><w:pPr/></w:pPrChange></w:pPr>`);
    const outcome = compareRevisionOutcomes(doc(body));

    expect(outcome.accepted).toBe('unchanged');
    expect(outcome.differs).toBe(false);
  });

  it('does not count text in a nested paragraph twice', () => {
    // A text box holds its own w:p inside a run of the outer paragraph. Bucketing text by
    // its nearest ancestor paragraph is what keeps "inner" out of the outer paragraph.
    const body = para(
      `${run('outer')}<w:r><w:pict><w:txbxContent>${para(run('inner'))}</w:txbxContent></w:pict></w:r>`
    );
    const outcome = compareRevisionOutcomes(doc(body));

    expect(outcome.accepted).toBe('outer\ninner');
  });
});

describe('moves — paired by name, not by id', () => {
  /** Ids are distinct per call so two moves in one part do not collide. */
  const moveBody = (fromName: string, toName: string, base = 1) =>
    para(
      `<w:moveFromRangeStart w:id="${base}" w:name="${fromName}" w:author="Ann" w:date="2024-01-01T00:00:00Z"/>` +
        `<w:moveFrom ${meta(String(base + 1))}>${delRun('clause')}</w:moveFrom>` +
        `<w:moveFromRangeEnd w:id="${base}"/>`
    ) +
    para(
      `<w:moveToRangeStart w:id="${base + 2}" w:name="${toName}" w:author="Ann" w:date="2024-01-01T00:00:00Z"/>` +
        `<w:moveTo ${meta(String(base + 3))}>${run('clause')}</w:moveTo>` +
        `<w:moveToRangeEnd w:id="${base + 2}"/>`
    );

  it('pairs the two halves and names each one', () => {
    const index = readRevisions(doc(moveBody('move1', 'move1')), 'word/document.xml');

    expect(index.moves).toHaveLength(1);
    expect(index.moves[0].paired).toBe(true);
    expect(index.revisions.filter(r => r.category === 'move').map(r => r.moveName)).toEqual(['move1', 'move1']);
    expect(index.problems.map(p => p.code)).not.toContain('revision/unpaired-move');
  });

  it('reports a move whose halves do not pair, because it degrades into a delete plus insert', () => {
    const index = readRevisions(doc(moveBody('move1', 'move2')), 'word/document.xml');

    const unpaired = index.problems.filter(p => p.code === 'revision/unpaired-move');
    expect(unpaired).toHaveLength(2);
    expect(unpaired.map(p => p.subject?.name).sort()).toEqual(['move1', 'move2']);
    expect(unpaired.find(p => p.subject?.name === 'move1')?.message).toContain('deletion');
    expect(unpaired.find(p => p.subject?.name === 'move2')?.message).toContain('insertion');
  });

  it('names each move half from the range it is actually inside', () => {
    // Two moves in one part. Taking "the first range of the right half" would label both
    // sources "alpha" and report a pairing that the markup does not support.
    const index = readRevisions(doc(moveBody('alpha', 'alpha') + moveBody('beta', 'beta', 10)));

    expect(index.revisions.filter(r => r.kind === 'moveFrom').map(r => r.moveName)).toEqual(['alpha', 'beta']);
    expect(index.revisions.filter(r => r.kind === 'moveTo').map(r => r.moveName)).toEqual(['alpha', 'beta']);
  });

  it('leaves a move half outside every range unnamed rather than guessing', () => {
    const body = moveBody('alpha', 'alpha') + para(`<w:moveFrom ${meta('20')}>${delRun('loose')}</w:moveFrom>`);
    const index = readRevisions(doc(body));

    expect(index.revisions.filter(r => r.kind === 'moveFrom').map(r => r.moveName)).toEqual(['alpha', null]);
  });

  it('does not name a source half from a destination range it happens to sit in', () => {
    // Ranges nest and overlap freely, so a w:moveFrom can land inside a moveTo range.
    // Only a moveFromRangeStart names a source; taking whichever range encloses it would
    // report a pairing that does not exist.
    const body = para(
      `<w:moveToRangeStart w:id="1" w:name="alpha" w:author="Ann" w:date="2024-01-01T00:00:00Z"/>` +
        `<w:moveFrom ${meta('2')}>${delRun('clause')}</w:moveFrom>` +
        `<w:moveToRangeEnd w:id="1"/>`
    );
    const index = readRevisions(doc(body));

    expect(index.revisions.find(r => r.kind === 'moveFrom')?.moveName).toBeNull();
  });

  it('reports two source ranges sharing a name as ambiguous', () => {
    const index = readRevisions(doc(moveBody('move1', 'move1') + moveBody('move1', 'move9', 10)));

    const duplicate = index.problems.find(p => p.code === 'revision/duplicate-move-name');
    expect(duplicate?.subject?.name).toBe('move1');
    expect(duplicate?.subject?.half).toBe('moveFrom');
  });

  it('requires w:name, w:author and w:date on a move range marker', () => {
    // CT_MoveBookmark is the one track-change type where w:date IS required.
    const index = readRevisions(doc(para('<w:moveFromRangeStart w:id="1"/><w:moveFromRangeEnd w:id="1"/>')));

    const codes = index.problems.map(p => p.code);
    expect(codes).toContain('revision/unnamed-move-range');
    expect(codes).toContain('revision/missing-required-date');
    expect(codes).toContain('revision/missing-author');
  });

  it('pairMoves works off the bookmark reader’s ranges, not its own scan', () => {
    const parsed = doc(moveBody('move1', 'move1'));
    const ranges = readBookmarks(parsed).bookmarks.filter(b => b.kind !== 'bookmarkStart');

    expect(pairMoves(ranges).map(m => ({ name: m.name, paired: m.paired }))).toEqual([
      { name: 'move1', paired: true }
    ]);
  });
});

describe('required attributes — what the schema actually says', () => {
  it('treats a missing w:author as an error and a missing w:date as a note', () => {
    // The SDK schema puts a RequiredValidator on @w:author but NOT on @w:date for
    // CT_TrackChange and CT_RunTrackChange. Reporting both as errors would be wrong.
    const index = readRevisions(doc(para(`<w:ins w:id="1">${run('x')}</w:ins>`)), 'word/document.xml');

    const author = index.problems.find(p => p.code === 'revision/missing-author');
    const date = index.problems.find(p => p.code === 'revision/missing-date');
    expect(author?.severity).toBe('error');
    expect(date?.severity).toBe('note');
    expect(date?.message).toContain('does not require');
  });

  it('does not demand an author from w:tblGridChange, which has none in the schema', () => {
    // CT_TblGridChange declares only @w:id. Every other *Change type declares author too.
    const index = readRevisions(
      doc(`<w:tbl><w:tblGrid><w:tblGridChange w:id="7"><w:tblGrid/></w:tblGridChange></w:tblGrid></w:tbl>`)
    );

    expect(index.revisions.map(r => r.kind)).toEqual(['tblGridChange']);
    expect(index.problems.map(p => p.code)).not.toContain('revision/missing-author');
    expect(index.problems.map(p => p.code)).not.toContain('revision/missing-date');
  });

  it('reports a missing w:id', () => {
    const index = readRevisions(doc(para(`<w:ins w:author="Ann" w:date="2024-01-01T00:00:00Z">${run('x')}</w:ins>`)));

    expect(index.problems.map(p => p.code)).toContain('revision/missing-id');
  });

  it('rejects w:id="-1", the one integer the union type excludes', () => {
    const index = readRevisions(doc(para(`<w:ins ${meta('-1')}>${run('x')}</w:ins>`)));
    const clean = readRevisions(doc(para(`<w:ins ${meta('-2')}>${run('x')}</w:ins>`)));

    expect(index.problems.map(p => p.code)).toContain('revision/id-out-of-range');
    expect(clean.problems.map(p => p.code)).not.toContain('revision/id-out-of-range');
  });

  it('flags an author longer than the schema maximum without claiming Word refuses it', () => {
    const index = readRevisions(doc(para(ins('x', '1', 'A'.repeat(MAX_AUTHOR_LENGTH + 1)))));

    const problem = index.problems.find(p => p.code === 'revision/author-too-long');
    expect(problem?.severity).toBe('warning');
    expect(problem?.message).toContain('not confirmed');
  });
});

describe('text in the file that is not text in the document', () => {
  it('reports how much deleted text is still readable in the package', () => {
    const index = readRevisions(doc(para(del('confidential'))), 'word/document.xml');

    const problem = index.problems.find(p => p.code === 'revision/deleted-text-retained');
    expect(problem?.subject?.characters).toBe('12');
    expect(problem?.message).toContain('not redaction');
    expect(problem?.silent).toBe(true);
  });

  it('reports a w:t inside a deletion, which no validator objects to', () => {
    const index = readRevisions(doc(para(`<w:del ${meta('2')}>${run('deleted but spelled live')}</w:del>`)));

    const problem = index.problems.find(p => p.code === 'revision/live-text-in-deletion');
    expect(problem?.subject?.count).toBe('1');
    expect(problem?.severity).toBe('error');
  });

  it('does not flag a w:t inside an insertion nested in a deletion', () => {
    // The innermost revision decides the spelling, and w:ins legitimately holds w:t.
    const index = readRevisions(
      doc(para(`<w:del ${meta('2', 'Bo')}><w:ins ${meta('1')}>${run('draft')}</w:ins></w:del>`))
    );

    expect(index.problems.map(p => p.code)).not.toContain('revision/live-text-in-deletion');
  });

  it('reports a w:delText that no deletion encloses', () => {
    const index = readRevisions(doc(para(delRun('orphan'))));

    const problem = index.problems.find(p => p.code === 'revision/orphan-deleted-text');
    expect(problem?.message).toContain('not confirmed');
  });
});

describe('the shared markup id space', () => {
  it('reports a revision colliding with a bookmark, and says where to renumber from', () => {
    // The corruption class: a generator numbers its revisions from 1 and lands on an id
    // a bookmark already owns. Word refuses the file; every lenient reader opens it.
    const body = para(
      `<w:bookmarkStart w:id="1" w:name="Intro"/>${ins('added', '1')}<w:bookmarkEnd w:id="1"/>`
    );
    const index = readRevisions(doc(body), 'word/document.xml');

    const collision = index.problems.find(p => p.code === 'revision/id-collision');
    expect(collision?.subject?.id).toBe('1');
    expect(collision?.message).toContain('bookmarkStart');
    expect(collision?.remediation).toContain('2');
  });

  it('stays quiet about a collision no revision took part in', () => {
    // Two bookmarks sharing an id is the bookmark analyzer's finding, not this one's.
    const body = para(
      `<w:bookmarkStart w:id="1" w:name="A"/><w:bookmarkStart w:id="1" w:name="B"/>${ins('x', '5')}`
    );
    const index = readRevisions(doc(body));

    expect(index.problems.map(p => p.code)).not.toContain('revision/id-collision');
  });
});

describe('the headline finding', () => {
  it('says the part has two readings when accept and reject differ', () => {
    const index = readRevisions(doc(para(`${ins('new')}${del('old')}`)), 'word/document.xml');

    const headline = index.problems.find(p => p.code === 'revision/unaccepted-revisions');
    expect(headline?.message).toContain('two different things');
    expect(headline?.subject?.revisions).toBe('2');
    expect(headline?.severity).toBe('note');
  });

  it('says the changes cancel out when they do', () => {
    const index = readRevisions(doc(para(run('same'), `<w:pPr><w:pPrChange ${meta('4')}><w:pPr/></w:pPrChange></w:pPr>`)));

    expect(index.problems.find(p => p.code === 'revision/unaccepted-revisions')?.message).toContain('same text');
  });

  it('calls out paragraph-mark and format-only revisions separately', () => {
    const body =
      para(run('a'), markRevision('del')) +
      para(run('b'), `<w:pPr><w:pPrChange ${meta('4')}><w:pPr/></w:pPrChange></w:pPr>`);
    const index = readRevisions(doc(body));

    expect(index.problems.find(p => p.code === 'revision/paragraph-mark-revision')?.subject?.count).toBe('1');
    expect(index.problems.find(p => p.code === 'revision/format-only-revision')?.subject?.count).toBe('1');
  });

  it('marks every finding silent except the one Word refuses to open', () => {
    const index = readRevisions(doc(para(`${ins('new', '-1')}${del('old')}`)));

    for (const problem of index.problems) {
      expect(problem.silent).toBe(problem.code !== 'revision/id-out-of-range');
    }
  });
});

describe('readRevisionSettings', () => {
  it('reads tracking, move and formatting switches', () => {
    const settings = readRevisionSettings(
      settingsDoc('<w:trackRevisions/><w:doNotTrackMoves w:val="1"/><w:doNotTrackFormatting w:val="0"/>')
    );

    expect(settings.trackRevisions).toBe(true);
    expect(settings.doNotTrackMoves).toBe(true);
    // CT_OnOff with an explicit off is off, even though the element is present.
    expect(settings.doNotTrackFormatting).toBe(false);
  });

  it('only counts an explicit off in w:revisionView as hidden', () => {
    const hidden = readRevisionSettings(settingsDoc('<w:revisionView w:insDel="0"/>'));
    const shown = readRevisionSettings(settingsDoc('<w:revisionView w:formatting="1"/>'));

    expect(hidden.hiddenFromView).toEqual(['content revisions — insertions and deletions']);
    expect(shown.hiddenFromView).toEqual([]);
  });

  it('returns a quiet default when there is no settings part', () => {
    expect(readRevisionSettings(null).trackRevisions).toBe(false);
    expect(readRevisionSettings(null).hiddenFromView).toEqual([]);
  });

  it('flags revisions hidden from view only when there are revisions to hide', () => {
    const settings = readRevisionSettings(settingsDoc('<w:revisionView w:insDel="0"/>'));

    expect(checkRevisionVisibility(settings, 3)).toHaveLength(1);
    expect(checkRevisionVisibility(settings, 3)[0].code).toBe('revision/revisions-hidden-in-view');
    expect(checkRevisionVisibility(settings, 0)).toEqual([]);
    expect(checkRevisionVisibility(readRevisionSettings(null), 3)).toEqual([]);
  });
});

describe('computeRevisionEvidenceForMarkup — panel wiring', () => {
  const partsWith = (body: string, settings?: string) => ({
    'word/document.xml': `<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`,
    ...(settings === undefined ? {} : { 'word/settings.xml': `<?xml version="1.0"?><w:settings ${W}>${settings}</w:settings>` })
  });

  it('leads with both readings', () => {
    const evidence = computeRevisionEvidenceForMarkup(partsWith(para(`${run('The ')}${ins('new')}${del('old')}`)), '');

    expect(evidence).not.toBeNull();
    expect(evidence!.lines[1]).toContain('The new');
    expect(evidence!.lines[1]).toContain('The old');
    expect(evidence!.lines[0]).toContain('2 unaccepted tracked change(s)');
  });

  it('warns that a tag-blind extractor matches neither reading', () => {
    const evidence = computeRevisionEvidenceForMarkup(partsWith(para(`${ins('new')}${del('old')}`)), '');

    expect(evidence!.lines.join('\n')).toContain('matches neither reading');
  });

  it('describes the selected revision when the editor has one open', () => {
    const body = para(`${ins('inserted here', '11', 'Bo')}${del('gone', '12')}`);
    const raw = `<w:ins w:id="11" w:author="Bo" w:date="2024-01-01T00:00:00Z"><w:r><w:t>inserted here</w:t></w:r></w:ins>`;
    const evidence = computeRevisionEvidenceForMarkup(partsWith(body), raw);

    const line = evidence!.lines.find(l => l.startsWith('The selected'));
    expect(line).toContain('w:id="11"');
    expect(line).toContain('Bo');
    expect(line).toContain('inserted here');
  });

  it('reports the settings that change what an absence of markup means', () => {
    const evidence = computeRevisionEvidenceForMarkup(
      partsWith(para(ins('x')), '<w:trackRevisions/><w:doNotTrackMoves/><w:doNotTrackFormatting/>'),
      ''
    );
    const prose = evidence!.lines.join('\n');

    expect(prose).toContain('w:trackRevisions');
    expect(prose).toContain('does not mean nothing was moved');
    expect(prose).toContain('does not mean the formatting was never edited');
  });

  it('surfaces revisions suppressed from view, which readRevisions alone cannot see', () => {
    const evidence = computeRevisionEvidenceForMarkup(
      partsWith(para(ins('x')), '<w:revisionView w:insDel="0"/>'),
      ''
    );

    expect(evidence!.lines.join('\n')).toContain('suppresses the display');
  });

  it('admits what it did not read', () => {
    const evidence = computeRevisionEvidenceForMarkup(partsWith(para(ins('x'))), '');
    const unresolved = evidence!.unresolved.join('\n');

    expect(unresolved).toContain('Headers, footers');
    expect(unresolved).toContain('neither is a complete text extraction');
  });

  it('admits that undated revisions cannot be ordered, only when some are undated', () => {
    const undated = computeRevisionEvidenceForMarkup(
      partsWith(para(`<w:ins w:id="1" w:author="Ann">${run('x')}</w:ins>`)),
      ''
    );
    const dated = computeRevisionEvidenceForMarkup(partsWith(para(ins('x'))), '');

    expect(undated!.unresolved.some(u => u.includes('cannot be ordered in time'))).toBe(true);
    expect(dated!.unresolved.some(u => u.includes('cannot be ordered in time'))).toBe(false);
  });

  it('returns null when there is no Word body part', () => {
    expect(computeRevisionEvidenceForMarkup({ 'xl/workbook.xml': '<x/>' }, '')).toBeNull();
  });

  it('returns null for a clean document rather than padding the prompt', () => {
    expect(computeRevisionEvidenceForMarkup(partsWith(para(run('plain'))), '')).toBeNull();
  });

  it('still speaks up for a clean document that has tracking switched on', () => {
    const evidence = computeRevisionEvidenceForMarkup(partsWith(para(run('plain')), '<w:trackRevisions/>'), '');

    expect(evidence!.lines[0]).toContain('no tracked changes');
    expect(evidence!.lines.join('\n')).toContain('w:trackRevisions');
  });
});

describe('malformed input is tolerated', () => {
  it('returns null rather than throwing on unparseable markup', () => {
    expect(computeRevisionEvidenceForMarkup({ 'word/document.xml': '<w:document' }, '')).toBeNull();
  });

  it('ignores a broken settings part and still reads the body', () => {
    const evidence = computeRevisionEvidenceForMarkup(
      {
        'word/document.xml': `<?xml version="1.0"?><w:document ${W}><w:body>${para(ins('x'))}</w:body></w:document>`,
        'word/settings.xml': '<w:settings'
      },
      ''
    );

    expect(evidence!.lines[0]).toContain('1 unaccepted tracked change(s)');
  });

  it('ignores elements in another namespace that share these local names', () => {
    const other = new DOMParser().parseFromString(
      `<?xml version="1.0"?><w:document ${W}><w:body><w:p><ins xmlns="urn:other" id="1"/></w:p></w:body></w:document>`,
      'application/xml'
    );

    expect(readRevisions(other).revisions).toEqual([]);
  });

  it('handles an empty document', () => {
    const index = readRevisions(doc(''));

    expect(index.revisions).toEqual([]);
    expect(compareRevisionOutcomes(doc('')).accepted).toBe('');
  });
});
