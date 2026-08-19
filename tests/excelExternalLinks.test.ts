import { describe, it, expect } from 'vitest';
import {
  readExternalLinks,
  externalLinkFindings,
  externalIndexesIn,
  externalSourceIsPresent,
  computeExternalLinkEvidenceForMarkup,
  EXTERNAL_LINK_PART,
  EXTERNAL_LINK_HOST_PART,
  EXTERNAL_LINK_FORMULA_PART,
  EXTERNAL_LINK_RELATIONSHIP_TYPE
} from '../services/excelExternalLinks';
import type { PackageParts } from '../services/packageIntegrity';

const S = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const rels = (body: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;

const rel = (id: string, target: string, external = false, type = 'externalLink') =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"${
    external ? ' TargetMode="External"' : ''
  }/>`;

/** `xl/workbook.xml` with an arbitrary body between the sheet list and the end. */
const workbook = (body: string) =>
  `<?xml version="1.0"?><workbook ${S} ${R}><sheets><sheet name="S1" sheetId="1" r:id="rIdSheet"/></sheets>${body}</workbook>`;

const references = (...ids: (string | null)[]) =>
  workbook(
    `<externalReferences>${ids
      .map(id => (id === null ? '<externalReference/>' : `<externalReference r:id="${id}"/>`))
      .join('')}</externalReferences>`
  );

const link = (body: string) => `<?xml version="1.0"?><externalLink ${S} ${R}>${body}</externalLink>`;

const book = (body: string, id: string | null = 'rId1') =>
  link(`<externalBook${id === null ? '' : ` r:id="${id}"`}>${body}</externalBook>`);

/** sheetNames + a one-cell cached sheet, the shape Excel writes for a live link. */
const CACHE =
  '<sheetNames><sheetName val="Budget"/><sheetName val="Actuals"/></sheetNames>' +
  '<sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>42</v></cell></row></sheetData></sheetDataSet>';

/** The healthy package: one reference, one link part, one external target, cached values. */
const healthy = (overrides: Partial<PackageParts> = {}): PackageParts => ({
  'xl/workbook.xml': references('rId5'),
  'xl/_rels/workbook.xml.rels': rels(rel('rId5', 'externalLinks/externalLink1.xml')),
  'xl/externalLinks/externalLink1.xml': book(CACHE),
  'xl/externalLinks/_rels/externalLink1.xml.rels': rels(
    rel('rId1', 'file:///Z:/finance/Source.xlsx', true, 'externalLinkPath')
  ),
  ...overrides
});

const codes = (parts: PackageParts) => externalLinkFindings(parts).map(f => f.code);

describe('the chain from workbook.xml outwards', () => {
  it('resolves a healthy link end to end and reports nothing about it', () => {
    const set = readExternalLinks(healthy());

    expect(set.workbookRead).toBe(true);
    expect(set.references).toHaveLength(1);
    const [reference] = set.references;
    expect(reference.index).toBe(1);
    expect(reference.relationshipId).toBe('rId5');
    expect(reference.partPath).toBe('xl/externalLinks/externalLink1.xml');
    expect(reference.kind).toBe('externalBook');
    expect(reference.problems).toEqual([]);
    expect(set.problems).toEqual([]);
    expect(externalLinkFindings(healthy())).toEqual([]);
  });

  it('reads the source location as the relationship states it, and never resolves it as a path', () => {
    const [reference] = readExternalLinks(healthy()).references;

    expect(reference.book?.target).toBe('file:///Z:/finance/Source.xlsx');
    expect(reference.book?.targetIsExternal).toBe(true);
  });

  it('reads sheet names, defined names and the cached cells', () => {
    const parts = healthy({
      'xl/externalLinks/externalLink1.xml': book(
        '<sheetNames><sheetName val="Budget"/><sheetName val="Actuals"/></sheetNames>' +
          '<definedNames><definedName name="Total" refersTo="[1]Budget!$A$1"/></definedNames>' +
          '<sheetDataSet><sheetData sheetId="0">' +
          '<row r="1"><cell r="A1"><v>1</v></cell><cell r="B1"><v>2</v></cell></row>' +
          '<row r="2"><cell r="A2"><v>3</v></cell></row>' +
          '</sheetData></sheetDataSet>'
      )
    });
    const bookRead = readExternalLinks(parts).references[0].book;

    expect(bookRead?.sheetNames).toEqual(['Budget', 'Actuals']);
    expect(bookRead?.definedNames).toEqual(['Total']);
    expect(bookRead?.cachedSheets).toEqual([
      { sheetId: '0', sheetName: 'Budget', refreshError: false, rowCount: 2, cellCount: 3 }
    ]);
  });

  it('numbers references by position, not by relationship id', () => {
    // The whole point of the module: [N] counts entries in document order. A package
    // whose ids are out of order and whose part names disagree with the position is the
    // only fixture where the two readings can be told apart.
    const parts: PackageParts = {
      'xl/workbook.xml': references('rId9', 'rId3'),
      'xl/_rels/workbook.xml.rels': rels(
        rel('rId3', 'externalLinks/externalLink1.xml') + rel('rId9', 'externalLinks/externalLink2.xml')
      ),
      'xl/externalLinks/externalLink1.xml': book(CACHE),
      'xl/externalLinks/externalLink2.xml': book(CACHE),
      'xl/externalLinks/_rels/externalLink1.xml.rels': rels(rel('rId1', 'a.xlsx', true)),
      'xl/externalLinks/_rels/externalLink2.xml.rels': rels(rel('rId1', 'b.xlsx', true))
    };

    const set = readExternalLinks(parts);
    expect(set.references.map(r => [r.index, r.relationshipId, r.partPath])).toEqual([
      [1, 'rId9', 'xl/externalLinks/externalLink2.xml'],
      [2, 'rId3', 'xl/externalLinks/externalLink1.xml']
    ]);
    // [1] is externalLink2.xml, whose source is b.xlsx. Reading the digit in the part
    // name, or the digits in the r:id, would both give the other answer.
    expect(set.references[0].book?.target).toBe('b.xlsx');
  });

  it('reports no references, rather than "not checked", for a workbook with no external links', () => {
    const set = readExternalLinks({ 'xl/workbook.xml': workbook('') });

    expect(set.workbookRead).toBe(true);
    expect(set.references).toEqual([]);
  });

  it('says the workbook was not read when it is absent, unparseable, or not a workbook', () => {
    for (const parts of [
      {},
      { 'xl/workbook.xml': '<workbook><unclosed>' },
      { 'xl/workbook.xml': `<?xml version="1.0"?><worksheet ${S}/>` },
      // Right local name, wrong namespace: conformance.ts normalises Strict to
      // Transitional before any analyzer runs, so anything else is genuinely foreign.
      { 'xl/workbook.xml': '<?xml version="1.0"?><workbook xmlns="urn:nonsense"/>' }
    ]) {
      expect(readExternalLinks(parts).workbookRead).toBe(false);
      expect(readExternalLinks(parts).references).toEqual([]);
    }
  });

  it('only counts externalReference children of externalReferences', () => {
    const parts = healthy({
      'xl/workbook.xml': workbook(
        `<externalReferences><externalReference r:id="rId5"/><notAReference r:id="rId5"/></externalReferences>`
      )
    });

    expect(readExternalLinks(parts).references).toHaveLength(1);
  });
});

describe('every hop that can break, and the workbook that still shows its numbers', () => {
  it('reports an externalReference with no r:id, and names the position it still occupies', () => {
    const parts = healthy({ 'xl/workbook.xml': references(null) });
    const [problem] = readExternalLinks(parts).references[0].problems;

    expect(problem.code).toBe('externalLink/reference-no-relationship-id');
    expect(problem.severity).toBe('error');
    expect(problem.silent).toBe(true);
    expect(problem.part).toBe('xl/workbook.xml');
    expect(problem.subject?.index).toBe('1');
  });

  it('treats an empty r:id as declared-and-useless, not as absent', () => {
    const parts = healthy({ 'xl/workbook.xml': references('') });
    const [reference] = readExternalLinks(parts).references;

    expect(reference.relationshipId).toBe('');
    expect(reference.problems[0].code).toBe('externalLink/reference-no-relationship-id');
  });

  it('reports a relationship the workbook rels does not declare', () => {
    const parts = healthy({ 'xl/_rels/workbook.xml.rels': rels(rel('rId1', 'externalLinks/externalLink1.xml')) });
    const [problem] = readExternalLinks(parts).references[0].problems;

    expect(problem.code).toBe('externalLink/reference-relationship-missing');
    expect(problem.subject?.relationshipId).toBe('rId5');
    expect(problem.message).toContain('does not declare');
    expect(problem.remediation).toContain(EXTERNAL_LINK_RELATIONSHIP_TYPE);
  });

  it('distinguishes a rels part that is missing entirely from one that omits the id', () => {
    const parts = healthy();
    delete parts['xl/_rels/workbook.xml.rels'];
    const [problem] = readExternalLinks(parts).references[0].problems;

    expect(problem.code).toBe('externalLink/reference-relationship-missing');
    expect(problem.message).toContain('does not exist');
    expect(problem.message).not.toContain('does not declare');
  });

  it('reports a relationship that resolves to a part the package does not contain', () => {
    const parts = healthy();
    delete parts['xl/externalLinks/externalLink1.xml'];
    const [reference] = readExternalLinks(parts).references;

    expect(reference.partPath).toBe('xl/externalLinks/externalLink1.xml');
    expect(reference.kind).toBeNull();
    expect(reference.problems[0].code).toBe('externalLink/link-part-missing');
    expect(reference.problems[0].subject?.target).toBe('xl/externalLinks/externalLink1.xml');
  });

  it('resolves a relationship target relative to the workbook part, not to the package root', () => {
    const parts = healthy({
      'xl/_rels/workbook.xml.rels': rels(rel('rId5', '../elsewhere/link.xml')),
      'elsewhere/link.xml': book(CACHE),
      'elsewhere/_rels/link.xml.rels': rels(rel('rId1', 'source.xlsx', true))
    });
    delete parts['xl/externalLinks/externalLink1.xml'];
    delete parts['xl/externalLinks/_rels/externalLink1.xml.rels'];

    const [reference] = readExternalLinks(parts).references;
    expect(reference.partPath).toBe('elsewhere/link.xml');
    expect(reference.problems).toEqual([]);
  });

  it('reports a link part that does not parse', () => {
    const parts = healthy({ 'xl/externalLinks/externalLink1.xml': '<externalLink><unclosed>' });
    const [reference] = readExternalLinks(parts).references;

    expect(reference.kind).toBe('unreadable');
    expect(reference.book).toBeNull();
    expect(reference.problems[0].code).toBe('externalLink/link-part-unreadable');
    expect(reference.problems[0].message).toContain('does not parse');
  });

  it('names the root element it found when a link part is something else entirely', () => {
    const parts = healthy({ 'xl/externalLinks/externalLink1.xml': `<?xml version="1.0"?><worksheet ${S}/>` });
    const [problem] = readExternalLinks(parts).references[0].problems;

    expect(problem.code).toBe('externalLink/link-part-unreadable');
    expect(problem.message).toContain('worksheet');
  });
});

describe('the externalBook and the source it names', () => {
  it('reports an externalBook with no r:id, where not even the path survives', () => {
    const parts = healthy({ 'xl/externalLinks/externalLink1.xml': book(CACHE, null) });
    const [reference] = readExternalLinks(parts).references;

    expect(reference.book?.relationshipId).toBeNull();
    expect(reference.book?.target).toBeNull();
    expect(reference.problems.map(p => p.code)).toContain('externalLink/book-no-relationship-id');
    // The cached values are still read: they are what a cell is displaying.
    expect(reference.book?.cachedSheets).toHaveLength(1);
  });

  it('reports an externalBook r:id the link part rels does not declare', () => {
    const parts = healthy({
      'xl/externalLinks/_rels/externalLink1.xml.rels': rels(rel('rIdOther', 'x.xlsx', true))
    });
    const [problem] = readExternalLinks(parts).references[0].problems;

    expect(problem.code).toBe('externalLink/book-relationship-missing');
    expect(problem.part).toBe('xl/externalLinks/externalLink1.xml');
    expect(problem.subject?.relationshipId).toBe('rId1');
    expect(problem.message).toContain('does not declare');
  });

  it('distinguishes a missing link rels part from one that omits the id', () => {
    const parts = healthy();
    delete parts['xl/externalLinks/_rels/externalLink1.xml.rels'];
    const [problem] = readExternalLinks(parts).references[0].problems;

    expect(problem.code).toBe('externalLink/book-relationship-missing');
    expect(problem.message).toContain('does not exist');
  });

  it('reports a relationship with no TargetMode="External" that resolves inside the package', () => {
    const parts = healthy({
      'xl/externalLinks/_rels/externalLink1.xml.rels': rels(rel('rId1', '../other.xlsx')),
      'xl/other.xlsx': 'BINARY'
    });
    const [reference] = readExternalLinks(parts).references;

    expect(reference.book?.targetIsExternal).toBe(false);
    expect(reference.book?.target).toBe('xl/other.xlsx');
    expect(reference.problems.map(p => p.code)).toEqual(['externalLink/source-not-external']);
    expect(reference.problems[0].severity).toBe('warning');
  });

  it('reports a non-external relationship whose package path is not there', () => {
    const parts = healthy({
      'xl/externalLinks/_rels/externalLink1.xml.rels': rels(rel('rId1', '../other.xlsx'))
    });
    const [problem] = readExternalLinks(parts).references[0].problems;

    expect(problem.code).toBe('externalLink/source-part-missing');
    expect(problem.severity).toBe('error');
    expect(problem.subject?.target).toBe('xl/other.xlsx');
  });
});

describe('the three states of "is the source there"', () => {
  it('answers null for a healthy external link, because a package cannot see outside itself', () => {
    const [reference] = readExternalLinks(healthy()).references;

    expect(reference.book?.sourceIsPresent).toBeNull();
    expect(externalSourceIsPresent(reference)).toBeNull();
    // And nothing in the findings calls the unreachable source a defect.
    expect(codes(healthy())).toEqual([]);
  });

  it('answers true only when the relationship pointed inside the package and the part is there', () => {
    const parts = healthy({
      'xl/externalLinks/_rels/externalLink1.xml.rels': rels(rel('rId1', '../other.xlsx')),
      'xl/other.xlsx': 'BINARY'
    });

    expect(externalSourceIsPresent(readExternalLinks(parts).references[0])).toBe(true);
  });

  it('answers false only when the relationship pointed inside the package and the part is gone', () => {
    const parts = healthy({
      'xl/externalLinks/_rels/externalLink1.xml.rels': rels(rel('rId1', '../other.xlsx'))
    });

    expect(externalSourceIsPresent(readExternalLinks(parts).references[0])).toBe(false);
  });

  it('answers null, not false, when there is no book to ask about', () => {
    const parts = healthy({ 'xl/externalLinks/externalLink1.xml': link('<ddeLink/>') });

    expect(externalSourceIsPresent(readExternalLinks(parts).references[0])).toBeNull();
  });
});

describe('an external-link part is not necessarily a workbook link', () => {
  it('reports a ddeLink as what it is, not as a malformed workbook link', () => {
    const parts = healthy({ 'xl/externalLinks/externalLink1.xml': link('<ddeLink ddeService="x" ddeTopic="y"/>') });
    const [reference] = readExternalLinks(parts).references;

    expect(reference.kind).toBe('ddeLink');
    expect(reference.book).toBeNull();
    expect(reference.problems).toEqual([]);
  });

  it('reports an oleLink as what it is', () => {
    const parts = healthy({ 'xl/externalLinks/externalLink1.xml': link(`<oleLink r:id="rId1" progId="Word.Document.12"/>`) });

    expect(readExternalLinks(parts).references[0].kind).toBe('oleLink');
  });

  it('reports a link part declaring none of the three, which the schema requires one of', () => {
    const parts = healthy({ 'xl/externalLinks/externalLink1.xml': link('<extLst/>') });
    const [reference] = readExternalLinks(parts).references;

    expect(reference.kind).toBe('unknown');
    expect(reference.problems[0].code).toBe('externalLink/link-kind-unknown');
    expect(reference.problems[0].severity).toBe('warning');
  });
});

describe('the cached values, and what they can and cannot be attributed to', () => {
  it('distinguishes no sheetDataSet at all from an empty one', () => {
    const none = healthy({ 'xl/externalLinks/externalLink1.xml': book('<sheetNames><sheetName val="A"/></sheetNames>') });
    const empty = healthy({
      'xl/externalLinks/externalLink1.xml': book('<sheetNames><sheetName val="A"/></sheetNames><sheetDataSet/>')
    });

    expect(readExternalLinks(none).references[0].book?.cachedSheets).toBeNull();
    expect(readExternalLinks(empty).references[0].book?.cachedSheets).toEqual([]);
    expect(codes(none)).toEqual(['externalLink/no-cached-values']);
    expect(codes(empty)).toEqual([]);
  });

  it('reads sheetId as a zero-based index into sheetNames', () => {
    const parts = healthy({
      'xl/externalLinks/externalLink1.xml': book(
        '<sheetNames><sheetName val="Budget"/><sheetName val="Actuals"/></sheetNames>' +
          '<sheetDataSet><sheetData sheetId="1"><row r="1"><cell r="A1"><v>7</v></cell></row></sheetData></sheetDataSet>'
      )
    });

    // One-based would name "Actuals" for sheetId="0"; zero-based names it for "1".
    expect(readExternalLinks(parts).references[0].book?.cachedSheets?.[0].sheetName).toBe('Actuals');
    expect(codes(parts)).toEqual([]);
  });

  it('warns when a sheetId selects none of the listed sheet names', () => {
    const parts = healthy({
      'xl/externalLinks/externalLink1.xml': book(
        '<sheetNames><sheetName val="Budget"/></sheetNames>' +
          '<sheetDataSet><sheetData sheetId="4"><row r="1"><cell r="A1"><v>7</v></cell></row></sheetData></sheetDataSet>'
      )
    });
    const [problem] = readExternalLinks(parts).references[0].problems;

    expect(problem.code).toBe('externalLink/cached-sheet-id-unknown');
    // A warning, not an error: the zero-based reading comes from one implementation
    // rather than from the schema, which types @sheetId only as UInt32.
    expect(problem.severity).toBe('warning');
    expect(problem.subject?.sheetId).toBe('4');
    expect(readExternalLinks(parts).references[0].book?.cachedSheets?.[0].sheetName).toBeNull();
  });

  it('does not read a missing or non-numeric sheetId as position zero', () => {
    const parts = (sheetId: string) =>
      healthy({
        'xl/externalLinks/externalLink1.xml': book(
          '<sheetNames><sheetName val="Budget"/></sheetNames>' +
            `<sheetDataSet><sheetData ${sheetId}><row r="1"><cell r="A1"><v>7</v></cell></row></sheetData></sheetDataSet>`
        )
      });

    for (const attribute of ['', 'sheetId=""', 'sheetId=" "', 'sheetId="1.5"', 'sheetId="-1"', 'sheetId="x"']) {
      const sheet = readExternalLinks(parts(attribute)).references[0].book?.cachedSheets?.[0];
      expect(sheet?.sheetName).toBeNull();
      expect(codes(parts(attribute))).toEqual(['externalLink/cached-sheet-id-unknown']);
    }
  });

  it('reports a sheet whose last refresh failed, with the name it resolves to', () => {
    const parts = healthy({
      'xl/externalLinks/externalLink1.xml': book(
        '<sheetNames><sheetName val="Budget"/></sheetNames>' +
          '<sheetDataSet><sheetData sheetId="0" refreshError="1">' +
          '<row r="1"><cell r="A1"><v>7</v></cell><cell r="B1"><v>8</v></cell></row>' +
          '</sheetData></sheetDataSet>'
      )
    });
    const [problem] = readExternalLinks(parts).references[0].problems;

    expect(problem.code).toBe('externalLink/refresh-error');
    expect(problem.subject).toEqual({ sheetId: '0', sheetName: 'Budget' });
    expect(problem.message).toContain('"Budget"');
    expect(problem.message).toContain('2 cached cell value(s)');
    expect(readExternalLinks(parts).references[0].book?.cachedSheets?.[0].refreshError).toBe(true);
  });

  it('reports a failed refresh even when the sheet cannot be named', () => {
    const parts = healthy({
      'xl/externalLinks/externalLink1.xml': book(
        '<sheetDataSet><sheetData sheetId="9" refreshError="true"><row r="1"/></sheetData></sheetDataSet>'
      )
    });
    const problem = readExternalLinks(parts).references[0].problems.find(
      p => p.code === 'externalLink/refresh-error'
    );

    expect(problem?.message).toContain('sheetId "9"');
    expect(problem?.subject).toEqual({ sheetId: '9' });
  });

  it('treats refreshError="0" as what it says', () => {
    const parts = healthy({
      'xl/externalLinks/externalLink1.xml': book(
        '<sheetNames><sheetName val="Budget"/></sheetNames>' +
          '<sheetDataSet><sheetData sheetId="0" refreshError="0"><row r="1"/></sheetData></sheetDataSet>'
      )
    });

    expect(codes(parts)).toEqual([]);
  });

  it('reports a sheetNames element with nothing in it, which the schema forbids', () => {
    const parts = healthy({
      'xl/externalLinks/externalLink1.xml': book(
        '<sheetNames/><sheetDataSet><sheetData sheetId="0"><row r="1"/></sheetData></sheetDataSet>'
      )
    });

    expect(codes(parts)).toEqual(['externalLink/empty-sheet-names']);
  });

  it('says nothing about sheet names when the optional element is simply absent', () => {
    const parts = healthy({
      'xl/externalLinks/externalLink1.xml': book(
        '<sheetDataSet><sheetData sheetId="0"><row r="1"/></sheetData></sheetDataSet>'
      )
    });

    expect(readExternalLinks(parts).references[0].book?.sheetNames).toEqual([]);
    expect(codes(parts)).toEqual([]);
  });
});

describe('link parts nothing points at', () => {
  it('reports a part no externalReference reaches, as a note rather than damage', () => {
    const parts = healthy({ 'xl/externalLinks/externalLink7.xml': book(CACHE) });
    const set = readExternalLinks(parts);

    expect(set.unreferencedParts).toEqual(['xl/externalLinks/externalLink7.xml']);
    expect(set.problems).toHaveLength(1);
    expect(set.problems[0].code).toBe('externalLink/unreferenced-link-part');
    expect(set.problems[0].severity).toBe('note');
    expect(set.problems[0].part).toBe('xl/externalLinks/externalLink7.xml');
  });

  it('finds parts under the Open XML SDK naming convention as well as Excel\'s', () => {
    const parts = healthy({ 'xl/externalReferences/externalReference1.xml': book(CACHE) });

    expect(readExternalLinks(parts).unreferencedParts).toEqual(['xl/externalReferences/externalReference1.xml']);
  });

  it('does not count a part that a reference does reach', () => {
    expect(readExternalLinks(healthy()).unreferencedParts).toEqual([]);
  });

  it('does not count the rels of an external-link part as an orphan link part', () => {
    expect(EXTERNAL_LINK_PART.test('xl/externalLinks/_rels/externalLink1.xml.rels')).toBe(false);
    expect(readExternalLinks(healthy()).unreferencedParts).toEqual([]);
  });

  it('says nothing about orphans when the workbook could not be read at all', () => {
    // Without the list there is no such thing as "unreferenced": nothing was checked.
    const set = readExternalLinks({ 'xl/externalLinks/externalLink1.xml': book(CACHE) });

    expect(set.workbookRead).toBe(false);
    expect(set.unreferencedParts).toEqual([]);
    expect(set.problems).toEqual([]);
  });
});

describe('the bracketed index in a formula', () => {
  it('finds the index in the forms Excel writes', () => {
    expect(externalIndexesIn('[1]Sheet1!A1')).toEqual([1]);
    expect(externalIndexesIn(`SUM('[2]Q1 Actuals'!$A$1:$A$9)`)).toEqual([2]);
    expect(externalIndexesIn('[1]S!A1+[2]S!A1')).toEqual([1, 2]);
    expect(externalIndexesIn('[12]S!A1')).toEqual([12]);
  });

  it('ignores a bracketed number inside a string literal', () => {
    expect(externalIndexesIn('IF(A1="[9]",1,2)')).toEqual([]);
    expect(externalIndexesIn('IF(A1="say ""[9]""",1,2)')).toEqual([]);
    // Two separate literals must not be read as one span that swallows the middle.
    expect(externalIndexesIn('"a" & [3]S!A1 & "b"')).toEqual([3]);
  });

  it('ignores a structured table reference, which is a column in this workbook', () => {
    expect(externalIndexesIn('SUM(Table1[2])')).toEqual([]);
    expect(externalIndexesIn('SUM(_x1[2])')).toEqual([]);
    expect(externalIndexesIn('SUM(A1.b[2])')).toEqual([]);
  });

  it('finds nothing in a formula with no external reference', () => {
    expect(externalIndexesIn('SUM(A1:A10)')).toEqual([]);
  });
});

describe('a formula naming a reference the workbook does not list', () => {
  const sheet = (formula: string, cellRef = 'B2') =>
    `<?xml version="1.0"?><worksheet ${S}><sheetData><row r="2"><c r="${cellRef}"><f>${formula}</f><v>1</v></c></row></sheetData></worksheet>`;

  it('reports an index past the end of the list, against the worksheet that writes it', () => {
    const parts = healthy({ 'xl/worksheets/sheet1.xml': sheet('[2]Sheet1!A1') });
    const [problem] = externalLinkFindings(parts).filter(f => f.code === 'externalLink/formula-index-unresolved');

    expect(problem.part).toBe('xl/worksheets/sheet1.xml');
    expect(problem.severity).toBe('error');
    expect(problem.silent).toBe(true);
    expect(problem.subject).toEqual({ index: '2', where: 'B2' });
    expect(problem.message).toContain('lists 1 external reference(s)');
    expect(problem.remediation).toContain('[1]–[1]');
  });

  it('says nothing about an index the list does cover', () => {
    expect(codes(healthy({ 'xl/worksheets/sheet1.xml': sheet('[1]Sheet1!A1') }))).toEqual([]);
  });

  it('never reports [0], which means this workbook rather than an out-of-range link', () => {
    expect(externalIndexesIn('[0]!Global_Range_Name')).toEqual([0]);
    expect(codes(healthy({ 'xl/worksheets/sheet1.xml': sheet('[0]!Global_Range_Name') }))).toEqual([]);
  });

  it('offers a different fix when the workbook lists no references at all', () => {
    const parts: PackageParts = {
      'xl/workbook.xml': workbook(''),
      'xl/worksheets/sheet1.xml': sheet('[1]Sheet1!A1')
    };
    const [problem] = externalLinkFindings(parts);

    expect(problem.code).toBe('externalLink/formula-index-unresolved');
    expect(problem.remediation).toContain('Add the externalReference');
    expect(problem.remediation).not.toContain('Repoint');
  });

  it('finds the index in a workbook-level defined name, whose expression is element text', () => {
    const parts = healthy({
      'xl/workbook.xml': workbook(
        `<externalReferences><externalReference r:id="rId5"/></externalReferences>` +
          `<definedNames><definedName name="Total">[4]Sheet1!$A$1</definedName></definedNames>`
      )
    });
    const [problem] = externalLinkFindings(parts).filter(f => f.code === 'externalLink/formula-index-unresolved');

    expect(problem.part).toBe('xl/workbook.xml');
    expect(problem.subject?.where).toBe('defined name "Total"');
  });

  it('reports one finding per cell and index, not one per occurrence', () => {
    const parts = healthy({ 'xl/worksheets/sheet1.xml': sheet('[3]S!A1+[3]S!A2') });

    expect(codes(parts)).toEqual(['externalLink/formula-index-unresolved']);
  });

  it('reports the same index in two cells separately, because each is a place to fix', () => {
    const parts = healthy({
      'xl/worksheets/sheet1.xml':
        `<?xml version="1.0"?><worksheet ${S}><sheetData><row r="2">` +
        `<c r="B2"><f>[3]S!A1</f><v>1</v></c><c r="C2"><f>[3]S!A2</f><v>1</v></c>` +
        `</row></sheetData></worksheet>`
    });

    expect(externalLinkFindings(parts).map(f => f.subject?.where)).toEqual(['B2', 'C2']);
  });

  it('reports nothing when the workbook could not be read, because there is no list to be out of', () => {
    const parts: PackageParts = {
      'xl/workbook.xml': '<workbook><unclosed>',
      'xl/worksheets/sheet1.xml': sheet('[7]Sheet1!A1')
    };

    expect(externalLinkFindings(parts)).toEqual([]);
  });

  it('skips a cell with no formula and a worksheet that does not parse', () => {
    const parts = healthy({
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet ${S}><sheetData><row r="1"><c r="A1"><v>[9]</v></c></row></sheetData></worksheet>`,
      'xl/worksheets/sheet2.xml': '<worksheet><unclosed>'
    });

    expect(codes(parts)).toEqual([]);
  });
});

describe('the parts this analyzer speaks for', () => {
  it('claims the workbook and the link parts as hosts, under either naming convention', () => {
    expect(EXTERNAL_LINK_HOST_PART.test('xl/workbook.xml')).toBe(true);
    expect(EXTERNAL_LINK_HOST_PART.test('xl/externalLinks/externalLink1.xml')).toBe(true);
    expect(EXTERNAL_LINK_HOST_PART.test('xl/externalReferences/externalReference1.xml')).toBe(true);
    expect(EXTERNAL_LINK_HOST_PART.test('xl/worksheets/sheet1.xml')).toBe(false);
    expect(EXTERNAL_LINK_HOST_PART.test('xl/externalLinks/_rels/externalLink1.xml.rels')).toBe(false);
  });

  it('claims worksheets as the parts that write the indexes', () => {
    expect(EXTERNAL_LINK_FORMULA_PART.test('xl/worksheets/sheet1.xml')).toBe(true);
    expect(EXTERNAL_LINK_FORMULA_PART.test('xl/worksheets/_rels/sheet1.xml.rels')).toBe(false);
    expect(EXTERNAL_LINK_FORMULA_PART.test('xl/workbook.xml')).toBe(false);
  });

  it('namespaces every finding it produces, and marks every one silent', () => {
    const parts = healthy({
      'xl/workbook.xml': references(null, 'rId5'),
      'xl/externalLinks/externalLink9.xml': book(CACHE)
    });
    const findings = externalLinkFindings(parts);

    expect(findings.length).toBeGreaterThan(1);
    for (const found of findings) {
      expect(found.code.startsWith('externalLink/')).toBe(true);
      // The whole thesis: the sheet displays its cached numbers either way.
      expect(found.silent).toBe(true);
      expect(found.remediation.length).toBeGreaterThan(0);
    }
  });
});

describe('evidence for the AI panel', () => {
  it('says nothing at all about a package with no external links', () => {
    expect(computeExternalLinkEvidenceForMarkup({ 'xl/workbook.xml': workbook('') })).toBeNull();
    expect(computeExternalLinkEvidenceForMarkup({})).toBeNull();
  });

  it('leads with the cache framing and explains what the bracketed index counts', () => {
    const evidence = computeExternalLinkEvidenceForMarkup(healthy());

    expect(evidence?.lines[0]).toContain('1 external workbook reference(s)');
    expect(evidence?.lines[0]).toContain('1-based index');
    expect(evidence?.lines[0]).toContain('cached');
  });

  it('describes each reference by its index, its part and its cache', () => {
    const evidence = computeExternalLinkEvidenceForMarkup(healthy());

    expect(evidence?.lines[1]).toContain('[1] is xl/externalLinks/externalLink1.xml');
    expect(evidence?.lines[1]).toContain('file:///Z:/finance/Source.xlsx');
    expect(evidence?.lines[1]).toContain('Budget, Actuals');
    expect(evidence?.lines[1]).toContain('1 cached cell value(s) across 1 sheet(s)');
  });

  it('records the external source as unresolved rather than as verified or broken', () => {
    const evidence = computeExternalLinkEvidenceForMarkup(healthy());

    expect(evidence?.unresolved.some(u => u.includes('was not read'))).toBe(true);
    expect(evidence?.unresolved.some(u => u.includes('their age cannot be established'))).toBe(true);
    expect(evidence?.lines.some(l => l.includes('cannot be determined from this package'))).toBe(true);
  });

  it('does not claim an unreadable source for a link that names none', () => {
    const parts = healthy({ 'xl/externalLinks/externalLink1.xml': book(CACHE, null) });
    const evidence = computeExternalLinkEvidenceForMarkup(parts);

    expect(evidence?.lines[1]).toContain('a source this package does not name');
    expect(evidence?.unresolved.some(u => u.includes('outside this package and was not read'))).toBe(false);
  });

  it('carries the findings into the evidence lines', () => {
    const parts = healthy();
    delete parts['xl/externalLinks/externalLink1.xml'];
    const evidence = computeExternalLinkEvidenceForMarkup(parts);

    expect(evidence?.lines.some(l => l.includes('is not in the package'))).toBe(true);
  });

  it('reports an orphaned link part in the evidence', () => {
    const parts: PackageParts = {
      'xl/workbook.xml': workbook(''),
      'xl/externalLinks/externalLink1.xml': book(CACHE)
    };
    const evidence = computeExternalLinkEvidenceForMarkup(parts);

    expect(evidence?.lines.some(l => l.includes('no externalReference'))).toBe(true);
  });

  it('reports a dangling formula index even when the workbook lists no references', () => {
    // The most alarming arrangement in the module: the formula names a source workbook
    // the file does not describe at all. Returning null here would say nothing.
    const parts: PackageParts = {
      'xl/workbook.xml': workbook(''),
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet ${S}><sheetData><row r="1"><c r="A1"><f>[1]S!A1</f><v>5</v></c></row></sheetData></worksheet>`
    };
    const evidence = computeExternalLinkEvidenceForMarkup(parts);

    expect(evidence?.lines.some(l => l.includes('there is no [1] to resolve'))).toBe(true);
  });

  it('describes a link part that turned out to be a DDE conversation', () => {
    const parts = healthy({ 'xl/externalLinks/externalLink1.xml': link('<ddeLink/>') });
    const evidence = computeExternalLinkEvidenceForMarkup(parts);

    expect(evidence?.lines[1]).toContain('which is a ddeLink rather than a workbook link');
  });
});
