import { describe, it, expect } from 'vitest';
import {
  readPivotTables,
  resolvePivotCacheChain,
  findSilentlyBrokenPivotTables,
  pivotTableErrors,
  describeBrokenHop
} from '../services/excelPivotTables';
import type { PackageParts } from '../services/packageIntegrity';

const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const rels = (body: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;

const rel = (id: string, type: string, target: string, external = false) =>
  `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`;

const workbook = (pivotCaches = '<pivotCaches><pivotCache cacheId="4" r:id="rId8"/></pivotCaches>') =>
  `<?xml version="1.0"?><workbook xmlns="${S}" xmlns:r="${R}">
    <sheets><sheet name="Sales" sheetId="1" r:id="rId1"/></sheets>${pivotCaches}</workbook>`;

const pivotTable = (attrs = 'cacheId="4"', body = '') =>
  `<?xml version="1.0"?><pivotTableDefinition xmlns="${S}" name="PivotTable1" dataCaption="Values" ${attrs}>
    <location ref="A3:D12" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
    <pivotFields count="3"><pivotField axis="axisRow" showAll="0"/><pivotField axis="axisCol" showAll="0"/><pivotField dataField="1" showAll="0"/></pivotFields>
    <rowFields count="1"><field x="0"/></rowFields>
    <colFields count="1"><field x="1"/></colFields>
    <dataFields count="1"><dataField name="Sum of Amount" fld="2" baseField="0" baseItem="0"/></dataFields>
    ${body}</pivotTableDefinition>`;

const cacheDefinition = (
  attrs = 'r:id="rId1" refreshOnLoad="0"',
  source = '<cacheSource type="worksheet"><worksheetSource ref="A1:C20" sheet="Sales"/></cacheSource>',
  fields = '<cacheFields count="3"><cacheField name="Region"/><cacheField name="Quarter"/><cacheField name="Amount"/></cacheFields>'
) => `<?xml version="1.0"?><pivotCacheDefinition xmlns="${S}" xmlns:r="${R}" ${attrs} recordCount="19">
    ${source}${fields}</pivotCacheDefinition>`;

/** A workbook with one pivot table whose three-hop chain resolves end to end. */
const pkg = (over: Partial<PackageParts> = {}): PackageParts => {
  const parts: PackageParts = {
    '_rels/.rels': rels(rel('rIdW', 'officeDocument', 'xl/workbook.xml')),
    'xl/workbook.xml': workbook(),
    'xl/_rels/workbook.xml.rels': rels(
      rel('rId1', 'worksheet', 'worksheets/sheet1.xml') +
        rel('rId8', 'pivotCacheDefinition', 'pivotCache/pivotCacheDefinition1.xml')
    ),
    'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="${S}"><sheetData/></worksheet>`,
    'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('rId3', 'pivotTable', '../pivotTables/pivotTable1.xml')),
    'xl/pivotTables/pivotTable1.xml': pivotTable(),
    'xl/pivotTables/_rels/pivotTable1.xml.rels': rels(
      rel('rId1', 'pivotCacheDefinition', '../pivotCache/pivotCacheDefinition1.xml')
    ),
    'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition(),
    'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': rels(
      rel('rId1', 'pivotCacheRecords', 'pivotCacheRecords1.xml')
    ),
    'xl/pivotCache/pivotCacheRecords1.xml': `<?xml version="1.0"?><pivotCacheRecords xmlns="${S}" count="19"><r><x v="0"/></r></pivotCacheRecords>`
  };
  for (const [path, content] of Object.entries(over)) {
    if (content === undefined) delete parts[path];
    else parts[path] = content;
  }
  return parts;
};

const only = (parts: PackageParts) => readPivotTables(parts)[0];
const kinds = (parts: PackageParts) => pivotTableErrors(only(parts)).map(p => p.code);

describe('the healthy chain', () => {
  it('resolves all three hops and reports nothing', () => {
    const table = only(pkg());

    expect(table.partPath).toBe('xl/pivotTables/pivotTable1.xml');
    expect(table.name).toBe('PivotTable1');
    expect(table.location).toBe('A3:D12');
    expect(table.chain.cacheId).toBe('4');
    expect(table.chain.cacheRelationshipId).toBe('rId8');
    expect(table.chain.cacheDefinitionPath).toBe('xl/pivotCache/pivotCacheDefinition1.xml');
    expect(table.chain.cacheRecordsPath).toBe('xl/pivotCache/pivotCacheRecords1.xml');
    expect(table.chain.cacheRecordsPresent).toBe(true);
    expect(table.chain.brokenHop).toBeNull();
    expect(describeBrokenHop(table.chain)).toBeNull();
    expect(pivotTableErrors(table)).toEqual([]);
    expect(findSilentlyBrokenPivotTables([table])).toEqual([]);
  });

  it('names the worksheet as owner, not the workbook that owns the cache', () => {
    // The two halves of a pivot table have different owners: the table part hangs off
    // the sheet it is drawn on, the cache off the workbook. A converter that walks only
    // one of them sees only half the pivot.
    const table = only(pkg());

    expect(table.ownerPath).toBe('xl/worksheets/sheet1.xml');
    expect(table.chain.workbookPath).toBe('xl/workbook.xml');
  });

  it('finds the workbook through _rels/.rels rather than assuming xl/workbook.xml', () => {
    const parts = pkg({ 'xl/workbook.xml': undefined as unknown as string });
    parts['_rels/.rels'] = rels(rel('rIdW', 'officeDocument', 'spreadsheet/book.xml'));
    parts['spreadsheet/book.xml'] = workbook();
    parts['spreadsheet/_rels/book.xml.rels'] = rels(
      rel('rId8', 'pivotCacheDefinition', '../xl/pivotCache/pivotCacheDefinition1.xml')
    );

    const table = only(parts);
    expect(table.chain.workbookPath).toBe('spreadsheet/book.xml');
    expect(table.chain.cacheDefinitionPath).toBe('xl/pivotCache/pivotCacheDefinition1.xml');
    expect(pivotTableErrors(table)).toEqual([]);
  });

  it('reads a Strict package, whose namespace drops the year', () => {
    // ISO Strict says http://purl.oclc.org/ooxml/spreadsheetml/main. Matching only the
    // Transitional URI would read no pivot tables and look like a clean workbook.
    const strict = (xml: string) =>
      xml
        .split('http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        .join('http://purl.oclc.org/ooxml/spreadsheetml/main')
        .split('http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        .join('http://purl.oclc.org/ooxml/officeDocument/relationships');

    const parts = Object.fromEntries(Object.entries(pkg()).map(([path, xml]) => [path, strict(xml)]));

    const table = only(parts);
    expect(table.chain.cacheDefinitionPath).toBe('xl/pivotCache/pivotCacheDefinition1.xml');
    expect(table.chain.brokenHop).toBeNull();
    expect(table.cacheFieldCount).toBe(3);
    expect(pivotTableErrors(table)).toEqual([]);
  });
});

describe('hop 1 — cacheId into the workbook', () => {
  it('names hop 1 when no pivotCache carries the cacheId', () => {
    const parts = pkg({ 'xl/workbook.xml': workbook('<pivotCaches><pivotCache cacheId="9" r:id="rId8"/></pivotCaches>') });
    const table = only(parts);

    expect(table.chain.brokenHop).toBe('table-to-workbook');
    expect(table.chain.problems.map(p => p.code)).toContain('pivot/cache-id-not-in-workbook');
    expect(describeBrokenHop(table.chain)).toContain('Hop 1 of 3');
    expect(table.chain.cacheDefinitionPath).toBeNull();
  });

  it('does not resolve a cacheId as if it were a relationship id', () => {
    // The two identifier spaces are the trap. Writing the pivotCache's r:id into
    // @cacheId is a real authoring bug, and it must not accidentally resolve.
    const parts = pkg({ 'xl/pivotTables/pivotTable1.xml': pivotTable('cacheId="rId8"') });
    const problem = only(parts).chain.problems.find(p => p.code === 'pivot/cache-id-not-in-workbook');

    expect(problem?.message).toContain('matched by value');
    expect(only(parts).chain.cacheDefinitionPath).toBeNull();
  });

  it('reports a workbook with no pivotCaches at all', () => {
    const parts = pkg({ 'xl/workbook.xml': workbook('') });

    expect(kinds(parts)).toContain('pivot/no-pivot-caches');
    expect(only(parts).chain.brokenHop).toBe('table-to-workbook');
  });

  it('reports a missing workbook part rather than blaming the pivot table', () => {
    const parts = pkg({ 'xl/workbook.xml': undefined as unknown as string });

    expect(kinds(parts)).toContain('pivot/workbook-missing');
  });

  it('reports a pivot table with no cacheId at all', () => {
    const parts = pkg({ 'xl/pivotTables/pivotTable1.xml': pivotTable('') });
    const table = only(parts);

    expect(table.chain.cacheId).toBeNull();
    expect(table.chain.brokenHop).toBe('table-to-workbook');
    expect(table.chain.problems[0].code).toBe('pivot/missing-required-attribute');
  });

  it('flags two pivotCaches claiming the same cacheId', () => {
    const parts = pkg({
      'xl/workbook.xml': workbook(
        '<pivotCaches><pivotCache cacheId="4" r:id="rId8"/><pivotCache cacheId="4" r:id="rId9"/></pivotCaches>'
      )
    });

    expect(kinds(parts)).toContain('pivot/duplicate-cache-id');
  });
});

describe('hop 2 — r:id into the workbook relationships', () => {
  it('reports a pivotCache missing its required r:id', () => {
    const parts = pkg({ 'xl/workbook.xml': workbook('<pivotCaches><pivotCache cacheId="4"/></pivotCaches>') });
    const table = only(parts);

    expect(table.chain.brokenHop).toBe('workbook-to-cache-definition');
    expect(table.chain.problems[0].code).toBe('pivot/missing-required-attribute');
    expect(table.chain.problems[0].message).toContain('not interchangeable');
  });

  it('names hop 2 when the relationship is not declared', () => {
    const parts = pkg({ 'xl/_rels/workbook.xml.rels': rels(rel('rId1', 'worksheet', 'worksheets/sheet1.xml')) });
    const table = only(parts);

    expect(table.chain.brokenHop).toBe('workbook-to-cache-definition');
    expect(table.chain.problems.map(p => p.code)).toContain('pivot/cache-relationship-missing');
    expect(describeBrokenHop(table.chain)).toContain('Hop 2 of 3');
  });

  it('catches a cache definition part that is simply gone', () => {
    const parts = pkg({ 'xl/pivotCache/pivotCacheDefinition1.xml': undefined as unknown as string });
    const problem = only(parts).chain.problems.find(p => p.code === 'pivot/cache-definition-missing');

    expect(problem?.silent).toBe(true);
    expect(problem?.message).toContain('renders');
    expect(only(parts).cacheFieldCount).toBeNull();
  });

  it('does not accept a pivot table part relationship as the cache relationship', () => {
    // pivotTable1.xml has its own .rels naming the same cache definition. That is not
    // the hop: the cache is reached from the workbook, and only from the workbook.
    const parts = pkg({ 'xl/_rels/workbook.xml.rels': undefined as unknown as string });
    const table = only(parts);

    expect(table.chain.cacheDefinitionPath).toBeNull();
    expect(table.chain.problems.map(p => p.code)).toContain('pivot/cache-relationship-missing');
  });
});

describe('hop 3 — the cache definition to its records', () => {
  it('names hop 3 when the records part is missing', () => {
    const parts = pkg({ 'xl/pivotCache/pivotCacheRecords1.xml': undefined as unknown as string });
    const table = only(parts);

    expect(table.chain.brokenHop).toBe('cache-definition-to-records');
    expect(table.chain.cacheRecordsPresent).toBe(false);
    expect(describeBrokenHop(table.chain)).toContain('Hop 3 of 3');
    expect(table.chain.problems.find(p => p.code === 'pivot/cache-records-missing')?.silent).toBe(true);
  });

  it('resolves the records relationship against the cache definition, not the workbook', () => {
    // The last hop uses the cache definition's own .rels. Declaring the records
    // relationship in the workbook's .rels instead looks plausible and resolves nothing.
    const parts = pkg({
      'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': undefined as unknown as string,
      'xl/_rels/workbook.xml.rels': rels(
        rel('rId1', 'pivotCacheRecords', 'pivotCache/pivotCacheRecords1.xml') +
          rel('rId8', 'pivotCacheDefinition', 'pivotCache/pivotCacheDefinition1.xml')
      )
    });
    const problem = only(parts).chain.problems.find(p => p.code === 'pivot/cache-records-missing');

    expect(problem?.message).toContain('own');
    expect(only(parts).chain.cacheRecordsPath).toBeNull();
  });

  it('treats absent records under refreshOnLoad as a note, not an error', () => {
    // A real, deliberate state: the cache rebuilds itself on open, so there is nothing
    // to store. Calling this damage would be a confidently wrong answer.
    const parts = pkg({
      'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition('refreshOnLoad="1"'),
      'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': undefined as unknown as string,
      'xl/pivotCache/pivotCacheRecords1.xml': undefined as unknown as string
    });
    const table = only(parts);
    const problem = table.chain.problems.find(p => p.code === 'pivot/cache-records-absent');

    expect(problem?.severity).toBe('note');
    expect(problem?.message).toContain('not damage');
    expect(table.chain.recordsAbsentIsExpected).toBe(true);
    expect(table.chain.brokenHop).toBeNull();
    expect(pivotTableErrors(table)).toEqual([]);
    expect(findSilentlyBrokenPivotTables([table])).toEqual([]);
  });

  it('accepts saveData="0" as the same deliberate state', () => {
    const parts = pkg({
      'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition('saveData="0"'),
      'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': undefined as unknown as string,
      'xl/pivotCache/pivotCacheRecords1.xml': undefined as unknown as string
    });

    expect(only(parts).chain.recordsAbsentIsExpected).toBe(true);
    expect(kinds(parts)).toEqual([]);
  });

  it('does flag absent records when nothing says they will be rebuilt', () => {
    const parts = pkg({
      'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition('refreshOnLoad="0"'),
      'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': undefined as unknown as string,
      'xl/pivotCache/pivotCacheRecords1.xml': undefined as unknown as string
    });
    const problem = only(parts).chain.problems.find(p => p.code === 'pivot/cache-records-absent');

    expect(problem?.severity).toBe('error');
    expect(problem?.message).toContain('not a schema violation');
    expect(only(parts).chain.recordsAbsentIsExpected).toBe(false);
  });
});

describe('cacheSource — where the data came from', () => {
  it('describes a worksheet range', () => {
    const table = only(pkg());

    expect(table.cacheSource?.type).toBe('worksheet');
    expect(table.cacheSource?.sheet).toBe('Sales');
    expect(table.cacheSource?.ref).toBe('A1:C20');
    expect(table.cacheSource?.description).toBe('A worksheet range: Sales!A1:C20.');
  });

  it('describes a defined name used instead of a sheet and range', () => {
    const parts = pkg({
      'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition(
        'r:id="rId1"',
        '<cacheSource type="worksheet"><worksheetSource name="SalesData"/></cacheSource>'
      )
    });

    expect(only(parts).cacheSource?.definedName).toBe('SalesData');
    expect(only(parts).cacheSource?.description).toContain('SalesData');
  });

  it('notes when the source range lives in another workbook', () => {
    const parts = pkg({
      'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition(
        'r:id="rId1"',
        `<cacheSource type="worksheet"><worksheetSource ref="A1:C20" sheet="Sales" r:id="rId7"/></cacheSource>`
      )
    });

    expect(only(parts).cacheSource?.externalRelationshipId).toBe('rId7');
    expect(only(parts).cacheSource?.description).toContain('another workbook');
  });

  it('describes an external connection', () => {
    const parts = pkg({
      'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition(
        'r:id="rId1"',
        '<cacheSource type="external" connectionId="2"/>'
      )
    });

    expect(only(parts).cacheSource?.connectionId).toBe('2');
    expect(only(parts).cacheSource?.description).toContain('connections.xml');
  });

  it('reports a cache with no cacheSource', () => {
    const parts = pkg({ 'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition('r:id="rId1"', '') });

    expect(kinds(parts)).toContain('pivot/no-cache-source');
    expect(only(parts).cacheSource).toBeNull();
  });

  it('reports a cacheSource with no type', () => {
    const parts = pkg({
      'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition(
        'r:id="rId1"',
        '<cacheSource><worksheetSource ref="A1:C20" sheet="Sales"/></cacheSource>'
      )
    });
    const table = only(parts);

    expect(table.problems.map(p => p.code)).toContain('pivot/missing-required-attribute');
    expect(table.cacheSource?.description).toContain('does not say');
  });
});

describe('field indices into the cache', () => {
  it('resolves each index to the cache field it names', () => {
    const table = only(pkg());

    expect(table.cacheFieldCount).toBe(3);
    expect(table.pivotFieldCount).toBe(3);
    expect(table.fieldReferences).toEqual([
      { origin: 'rowFields', index: 0, cacheFieldName: 'Region' },
      { origin: 'colFields', index: 1, cacheFieldName: 'Quarter' },
      { origin: 'dataFields', index: 2, cacheFieldName: 'Amount' }
    ]);
  });

  it('catches an index past the end of the cache fields', () => {
    const parts = pkg({
      'xl/pivotTables/pivotTable1.xml': pivotTable().replace('fld="2"', 'fld="7"')
    });
    const problem = only(parts).problems.find(p => p.code === 'pivot/field-index-out-of-range');

    expect(problem?.message).toContain('only 3 fields');
    expect(problem?.message).toContain('blank');
    expect(only(parts).fieldReferences.find(r => r.index === 7)?.cacheFieldName).toBeNull();
  });

  it('catches the index one past the last field, which is the one that gets written', () => {
    // Indices are zero-based, so with three cache fields index 3 is already off the
    // end. This is the off-by-one an author actually makes, and a bound that only
    // rejects indices strictly greater than the count lets it through as valid.
    const parts = pkg({ 'xl/pivotTables/pivotTable1.xml': pivotTable().replace('fld="2"', 'fld="3"') });

    expect(kinds(parts)).toContain('pivot/field-index-out-of-range');
    expect(only(parts).problems.find(p => p.code === 'pivot/field-index-out-of-range')?.message).toContain('indices 0–2');
  });

  it('accepts the last valid index', () => {
    // The other half of the boundary: index 2 of 3 fields must not be flagged.
    expect(kinds(pkg())).toEqual([]);
    expect(only(pkg()).fieldReferences.at(-1)?.cacheFieldName).toBe('Amount');
  });

  it('reads pageField/@fld as a field index too', () => {
    const parts = pkg({
      'xl/pivotTables/pivotTable1.xml': pivotTable('cacheId="4"', '<pageFields count="1"><pageField fld="5"/></pageFields>')
    });

    expect(only(parts).fieldReferences.map(r => r.origin)).toContain('pageFields');
    expect(kinds(parts)).toContain('pivot/field-index-out-of-range');
  });

  it('does not treat x="-2" as a dangling reference', () => {
    // -2 marks the "values" pseudo field rather than a real cache field index. Range
    // checking it would report a false break on most multi-measure pivots.
    const parts = pkg({
      'xl/pivotTables/pivotTable1.xml': pivotTable().replace('<field x="1"/>', '<field x="1"/><field x="-2"/>')
    });

    expect(kinds(parts)).toEqual([]);
    expect(only(parts).fieldReferences.some(r => r.index === -2)).toBe(false);
  });

  it('cannot check indices when the cache did not resolve, and does not pretend to', () => {
    const parts = pkg({
      'xl/pivotCache/pivotCacheDefinition1.xml': undefined as unknown as string,
      'xl/pivotTables/pivotTable1.xml': pivotTable().replace('fld="2"', 'fld="7"')
    });
    const table = only(parts);

    expect(table.cacheFieldCount).toBeNull();
    expect(table.problems.map(p => p.code)).not.toContain('pivot/field-index-out-of-range');
    expect(table.fieldReferences.find(r => r.index === 7)?.cacheFieldName).toBeNull();
  });

  it('flags pivotFields and cacheFields disagreeing in length', () => {
    const parts = pkg({
      'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition(
        'r:id="rId1"',
        undefined,
        '<cacheFields count="2"><cacheField name="Region"/><cacheField name="Quarter"/></cacheFields>'
      )
    });
    const problem = only(parts).problems.find(p => p.code === 'pivot/field-count-mismatch');

    expect(problem?.message).toContain('3 pivotField');
    expect(problem?.message).toContain('no index attribute');
  });

  it('flags a cacheFields/@count that disagrees with the elements present', () => {
    const parts = pkg({
      'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition(
        'r:id="rId1"',
        undefined,
        '<cacheFields count="9"><cacheField name="Region"/><cacheField name="Quarter"/><cacheField name="Amount"/></cacheFields>'
      )
    });
    const messages = only(parts)
      .problems.filter(p => p.code === 'pivot/field-count-mismatch')
      .map(p => p.message);

    expect(messages.some(m => m.includes('@count says 9'))).toBe(true);
    expect(only(parts).cacheFieldCount).toBe(3);
  });
});

describe('discovery and reporting', () => {
  it('reports a pivot table part no worksheet relates to', () => {
    // Nothing in worksheet XML names a pivot table, so deleting the relationship leaves
    // a perfectly valid worksheet and a pivot table that will never appear.
    const parts = pkg({ 'xl/worksheets/_rels/sheet1.xml.rels': rels('') });
    const table = only(parts);

    expect(table.ownerPath).toBeNull();
    expect(table.problems.map(p => p.code)).toContain('pivot/orphan-pivot-table-part');
    expect(table.problems.find(p => p.code === 'pivot/orphan-pivot-table-part')?.message).toContain('implicit relationship');
  });

  it('finds a relationship-declared pivot table outside the conventional folder', () => {
    const parts = pkg({ 'xl/pivotTables/pivotTable1.xml': undefined as unknown as string });
    parts['xl/pt/custom.xml'] = pivotTable();
    parts['xl/worksheets/_rels/sheet1.xml.rels'] = rels(rel('rId3', 'pivotTable', '../pt/custom.xml'));

    const table = only(parts);
    expect(table.partPath).toBe('xl/pt/custom.xml');
    expect(table.ownerPath).toBe('xl/worksheets/sheet1.xml');
    expect(pivotTableErrors(table)).toEqual([]);
  });

  it('reports missing required attributes on the definition itself', () => {
    const parts = pkg({
      'xl/pivotTables/pivotTable1.xml': `<?xml version="1.0"?><pivotTableDefinition xmlns="${S}" cacheId="4"/>`
    });

    expect(only(parts).problems.map(p => p.code)).toEqual([
      'pivot/missing-required-attribute',
      'pivot/missing-required-attribute'
    ]);
  });

  it('lists every break as silent, because none of them change what is displayed', () => {
    const parts = pkg({ 'xl/pivotCache/pivotCacheRecords1.xml': undefined as unknown as string });
    const tables = readPivotTables(parts);

    expect(findSilentlyBrokenPivotTables(tables)).toHaveLength(1);
    expect(pivotTableErrors(tables[0]).every(p => p.silent)).toBe(true);
  });

  it('returns nothing for a workbook with no pivot tables', () => {
    expect(readPivotTables({ 'xl/workbook.xml': workbook('') })).toEqual([]);
  });

  it('reports an unparseable pivot table part instead of throwing', () => {
    const chain = resolvePivotCacheChain(pkg({ 'xl/pivotTables/pivotTable1.xml': '<not xml' }), 'xl/pivotTables/pivotTable1.xml');

    expect(chain.brokenHop).toBe('table-to-workbook');
    expect(chain.problems[0].message).toContain('well-formed');
  });

  it('reports an absent pivot table part asked for by path', () => {
    const chain = resolvePivotCacheChain(pkg(), 'xl/pivotTables/nope.xml');

    expect(chain.problems[0].code).toBe('pivot/missing-required-attribute');
    expect(chain.cacheId).toBeNull();
  });
});
