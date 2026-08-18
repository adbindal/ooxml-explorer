import { describe, it, expect } from 'vitest';
import {
  finding,
  compareFindings,
  analyzerOf,
  findingsFrom,
  silentFindings,
  errorsOnly,
  renderFinding,
  renderFindings,
  type Finding
} from '../services/findings';

const f = (over: Partial<Finding> = {}): Finding => ({
  code: 'demo/thing',
  severity: 'error',
  part: 'word/document.xml',
  message: 'Something is wrong.',
  remediation: 'Fix it.',
  silent: false,
  ...over
});

describe('finding()', () => {
  it('defaults to error and not-silent, which under-claims rather than over-claims', () => {
    // An analyzer that forgets to mark something silent reports a defect the reader
    // can see. The opposite mistake hides one they cannot.
    const result = finding('demo/x', 'word/document.xml', 'msg', 'fix');

    expect(result.severity).toBe('error');
    expect(result.silent).toBe(false);
  });

  it('omits subject entirely rather than emitting an empty object', () => {
    // A consumer checking `if (f.subject)` should not get a truthy empty object.
    expect(finding('demo/x', 'p', 'm', 'r')).not.toHaveProperty('subject');
    expect(finding('demo/x', 'p', 'm', 'r', { subject: { id: '1' } }).subject).toEqual({ id: '1' });
  });
});

describe('ordering', () => {
  it('puts errors before warnings before notes', () => {
    const sorted = [f({ severity: 'note' }), f({ severity: 'error' }), f({ severity: 'warning' })].sort(
      compareFindings
    );

    expect(sorted.map(x => x.severity)).toEqual(['error', 'warning', 'note']);
  });

  it('puts silent findings ahead of visible ones at equal severity', () => {
    // The visible ones are already discoverable by looking at the document. The list
    // should lead with what looking cannot tell you.
    //
    // The codes are chosen so alphabetical order CONTRADICTS the expected order: an
    // earlier version used 'a/silent' and 'a/visible', which sort correctly by name
    // alone, so the test passed even with the silence tiebreak deleted.
    const sorted = [f({ code: 'a/aaa', silent: false }), f({ code: 'a/zzz', silent: true })].sort(
      compareFindings
    );

    expect(sorted.map(x => x.code)).toEqual(['a/zzz', 'a/aaa']);
  });

  it('is deterministic for otherwise-equal findings', () => {
    const sorted = [f({ code: 'z/b' }), f({ code: 'a/a' })].sort(compareFindings);

    expect(sorted.map(x => x.code)).toEqual(['a/a', 'z/b']);
  });

  it('decides what survives truncation, so renderFindings sorts before rendering', () => {
    const lines = renderFindings([
      f({ severity: 'note', message: 'Least important.' }),
      f({ severity: 'error', message: 'Most important.' })
    ]);

    expect(lines[0]).toContain('Most important');
  });
});

describe('grouping and filtering', () => {
  it('reads the analyzer out of the namespaced code', () => {
    expect(analyzerOf('ole/data-part-missing')).toBe('ole');
    expect(findingsFrom([f({ code: 'ole/x' }), f({ code: 'bookmark/y' })], 'ole')).toHaveLength(1);
  });

  it('separates silent from visible, and errors from the rest', () => {
    const all = [f({ silent: true }), f({ silent: false, severity: 'note' })];

    expect(silentFindings(all)).toHaveLength(1);
    expect(errorsOnly(all)).toHaveLength(1);
  });
});

describe('renderFinding', () => {
  it('joins the consequence to the fix', () => {
    expect(renderFinding(f())).toBe('Something is wrong. Fix it.');
  });

  it('marks a note, because an unmarked note reads as a defect', () => {
    expect(renderFinding(f({ severity: 'note' }))).toMatch(/^Note: /);
  });

  it('says when something renders correctly and is broken anyway', () => {
    expect(renderFinding(f({ silent: true }))).toContain('renders correctly and is broken anyway');
  });

  it('does not say that about a visible problem', () => {
    expect(renderFinding(f({ silent: false }))).not.toContain('broken anyway');
  });
});

describe('the contract other agents depend on', () => {
  it('survives a JSON round-trip unchanged', () => {
    // The stated goal is feeding this to another agent. Anything that does not
    // serialise cleanly - a Map, an Element, a Date - breaks that quietly.
    const original = f({ subject: { id: '3', name: 'Chapter' } });

    expect(JSON.parse(JSON.stringify(original))).toEqual(original);
  });

  it('namespaces every code, so two analyzers cannot collide on a generic kind', () => {
    // 'duplicate-id' is used by bookmarks AND comments. Unnamespaced they would be
    // indistinguishable to any consumer grouping or suppressing by code.
    expect(analyzerOf('bookmark/duplicate-id')).not.toBe(analyzerOf('comment/duplicate-id'));
  });
});

describe('every analyzer obeys the contract', () => {
  // A shared type is only worth having if nothing quietly opts out of it. These run the
  // real analyzers over deliberately broken input and check the records they emit.
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const O = 'xmlns:o="urn:schemas-microsoft-com:office:office"';
  const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

  const collectAll = async (): Promise<Finding[]> => {
    const { readBookmarks } = await import('../services/wordBookmarks');
    const { readOleObjects } = await import('../services/oleObjects');
    const { readComments } = await import('../services/wordComments');
    const { readPivotTables } = await import('../services/excelPivotTables');
    const { checkPackageIntegrity } = await import('../services/packageIntegrity');
    const parse = (xml: string) => new DOMParser().parseFromString(xml, 'application/xml');

    const out: Finding[] = [];

    // a bookmark that opens and never closes, plus a colliding tracked-change id
    out.push(
      ...readBookmarks(
        parse(`<?xml version="1.0"?><w:document ${W}><w:body><w:p>
          <w:bookmarkStart w:id="1" w:name="Open"/>
          <w:bookmarkEnd w:id="9"/></w:p></w:body></w:document>`),
        'word/document.xml'
      ).problems
    );

    // an OLE object whose embedding is gone but whose preview is not
    out.push(
      ...readOleObjects(
        {
          'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${O} ${R}><w:body><w:p><w:r><w:object>
            <o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" r:id="rId4"/></w:object></w:r></w:p></w:body></w:document>`,
          'word/_rels/document.xml.rels': `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/></Relationships>`
        },
        'word/document.xml'
      ).flatMap(o => o.problems)
    );

    // a comment anchored in the body with no comments.xml behind it
    out.push(
      ...readComments({
        document: parse(`<?xml version="1.0"?><w:document ${W}><w:body><w:p>
          <w:commentRangeStart w:id="1"/><w:r><w:t>x</w:t></w:r><w:commentRangeEnd w:id="1"/>
          </w:p></w:body></w:document>`)
      }).problems
    );

    // a pivot table part nothing relates to
    out.push(
      ...readPivotTables({
        'xl/pivotTables/pivotTable1.xml': `<?xml version="1.0"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P" cacheId="4" dataCaption="V"/>`
      }).flatMap(t => [...t.chain.problems, ...t.problems])
    );

    // a package with no [Content_Types].xml
    out.push(...checkPackageIntegrity({ 'word/document.xml': '<w:document/>' }));

    return out;
  };

  it('produces findings from all five analyzers', async () => {
    const analyzers = new Set((await collectAll()).map(f => analyzerOf(f.code)));

    expect([...analyzers].sort()).toEqual(['bookmark', 'comment', 'ole', 'package', 'pivot']);
  });

  it('namespaces every code as analyzer/kind', async () => {
    for (const found of await collectAll()) {
      expect(found.code).toMatch(/^[a-z]+\/[a-z0-9-]+$/);
    }
  });

  it('never emits an empty message or remediation', async () => {
    // The gap this refactor closed: integrity findings used to carry no remediation.
    for (const found of await collectAll()) {
      expect(found.message.length, found.code).toBeGreaterThan(0);
      expect(found.remediation.length, found.code).toBeGreaterThan(0);
    }
  });

  it('always names the part the fault is in', async () => {
    for (const found of await collectAll()) {
      expect(found.part.length, found.code).toBeGreaterThan(0);
    }
  });

  it('uses only the three declared severities', async () => {
    for (const found of await collectAll()) {
      expect(['error', 'warning', 'note']).toContain(found.severity);
    }
  });

  it('serialises every real finding as JSON', async () => {
    const all = await collectAll();

    expect(JSON.parse(JSON.stringify(all))).toEqual(all);
  });
});
