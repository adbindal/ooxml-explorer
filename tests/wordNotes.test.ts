import { describe, it, expect } from 'vitest';
import { readNotes, readNoteReferences, noteFindings, computeNoteEvidenceForMarkup, hasNotes } from '../services/wordNotes';
import type { PackageParts } from '../services/packageIntegrity';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const body = (inner: string) => `<?xml version="1.0"?><w:document ${W}><w:body>${inner}</w:body></w:document>`;
const ref = (id: string) => `<w:p><w:r><w:footnoteReference w:id="${id}"/></w:r></w:p>`;
const notes = (inner: string) => `<?xml version="1.0"?><w:footnotes ${W}>${inner}</w:footnotes>`;
const note = (id: string, text = 'note text', type = '') =>
  `<w:footnote w:id="${id}"${type ? ` w:type="${type}"` : ''}><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:footnote>`;

/** The separator pair Word writes into every document that has ever had a footnote. */
const separators = note('-1', 'sep', 'separator') + note('0', 'cont', 'continuationSeparator');

const parse = (xml: string) => new DOMParser().parseFromString(xml, 'application/xml');

describe('reading notes', () => {
  it('treats an absent w:type as normal, which is the schema default', () => {
    expect(readNotes(parse(notes(note('1'))), 'footnote')[0]).toMatchObject({ type: 'normal', structural: false });
  });

  it('marks the separator graphics as structural', () => {
    const read = readNotes(parse(notes(separators)), 'footnote');

    expect(read.every(n => n.structural)).toBe(true);
  });

  it('reads reference ids from the body', () => {
    expect(readNoteReferences(parse(body(ref('1') + ref('2'))), 'footnote')).toEqual(['1', '2']);
  });
});

describe('the separator trap', () => {
  it('does NOT report the separator notes as orphans', () => {
    // These are referenced from sectPr, never from the text, and Word writes them into
    // every document with a footnote. Reporting them puts two false positives on
    // essentially every real file - the fastest way to make a report unreadable.
    const parts: PackageParts = {
      'word/document.xml': body(ref('1')),
      'word/footnotes.xml': notes(separators + note('1'))
    };

    expect(noteFindings(parts).map(p => p.code)).not.toContain('note/orphan-note');
  });

  it('still reports a genuine orphan alongside the separators', () => {
    // The discriminating case: separators AND a real orphan in one file. A check that
    // exempted everything unreferenced would report nothing here.
    const parts: PackageParts = {
      'word/document.xml': body(ref('1')),
      'word/footnotes.xml': notes(separators + note('1') + note('2', 'never shown'))
    };

    const problem = noteFindings(parts).find(p => p.code === 'note/orphan-note');
    expect(problem?.subject?.id).toBe('2');
    expect(problem?.message).toContain('never shown');
  });
});

describe('broken pairing', () => {
  it('reports a reference with no note behind it', () => {
    const parts: PackageParts = {
      'word/document.xml': body(ref('7')),
      'word/footnotes.xml': notes(separators)
    };

    const problem = noteFindings(parts).find(p => p.code === 'note/missing-note');
    expect(problem?.silent).toBe(true);
    expect(problem?.message).toContain('no note at the bottom of the page');
  });

  it('reports two notes sharing an id', () => {
    const parts: PackageParts = {
      'word/document.xml': body(ref('1')),
      'word/footnotes.xml': notes(separators + note('1', 'first') + note('1', 'second'))
    };

    expect(noteFindings(parts).map(p => p.code)).toContain('note/duplicate-note-id');
  });

  it('reports a missing notes part only when something references it', () => {
    const withRefs: PackageParts = { 'word/document.xml': body(ref('1')), 'word/endnotes.xml': notes('') };
    const without: PackageParts = { 'word/document.xml': body(''), 'word/endnotes.xml': notes('') };

    expect(noteFindings(withRefs).map(p => p.code)).toContain('note/notes-part-missing');
    expect(noteFindings(without).map(p => p.code)).not.toContain('note/notes-part-missing');
  });

  it('finds references in headers and footers, not just document.xml', () => {
    const parts: PackageParts = {
      'word/header1.xml': body(ref('9')),
      'word/footnotes.xml': notes(separators)
    };

    expect(noteFindings(parts).map(p => p.subject?.id)).toContain('9');
  });

  it('does not report a note as orphaned when a header references it', () => {
    // The body-part sweep has to cover every story, or every note used only in a
    // header reports as unreferenced.
    const parts: PackageParts = {
      'word/document.xml': body(''),
      'word/header1.xml': body(ref('1')),
      'word/footnotes.xml': notes(separators + note('1'))
    };

    expect(noteFindings(parts).map(p => p.code)).not.toContain('note/orphan-note');
  });

  it('says nothing about a healthy document', () => {
    const parts: PackageParts = {
      'word/document.xml': body(ref('1')),
      'word/footnotes.xml': notes(separators + note('1'))
    };

    expect(noteFindings(parts)).toEqual([]);
  });
});

describe('separators', () => {
  it('notes a document that uses footnotes but declares no separator', () => {
    const parts: PackageParts = {
      'word/document.xml': body(ref('1')),
      'word/footnotes.xml': notes(note('1'))
    };

    expect(noteFindings(parts).map(p => p.code)).toContain('note/missing-separator');
  });

  it('does not ask for a separator in a document with no notes', () => {
    const parts: PackageParts = { 'word/document.xml': body(''), 'word/footnotes.xml': notes('') };

    expect(noteFindings(parts).map(p => p.code)).not.toContain('note/missing-separator');
  });
});

describe('evidence', () => {
  it('returns null for a package with no notes parts', () => {
    expect(hasNotes({ 'word/document.xml': body('') })).toBe(false);
    expect(computeNoteEvidenceForMarkup({ 'word/document.xml': body('') })).toBeNull();
  });

  it('counts content notes separately from separators, and explains why', () => {
    const evidence = computeNoteEvidenceForMarkup({
      'word/document.xml': body(ref('1')),
      'word/footnotes.xml': notes(separators + note('1'))
    });

    expect(evidence!.lines[0]).toContain('1 footnote(s) plus 2 separator note(s)');
    expect(evidence!.lines[0]).toContain('section properties');
  });

  it('caps the claim to the body parts it was given', () => {
    const evidence = computeNoteEvidenceForMarkup({
      'word/document.xml': body(''),
      'word/footnotes.xml': notes(separators + note('1'))
    });

    expect(evidence!.unresolved.some(u => u.includes('not loaded here'))).toBe(true);
  });
});
