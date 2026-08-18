/**
 * SpreadsheetML pivot tables — three parts, two identifier spaces, one silent failure.
 *
 * A pivot table is not one part. It is a chain of three, joined by two *different kinds*
 * of identifier, and a break at any link leaves a workbook that opens cleanly, renders
 * the pivot exactly as before, and is dead the moment anyone refreshes it:
 *
 *   xl/pivotTables/pivotTable1.xml   <pivotTableDefinition cacheId="4" …>
 *           │  @cacheId — an application-level number, NOT a relationship id
 *           ▼
 *   xl/workbook.xml                  <pivotCaches><pivotCache cacheId="4" r:id="rId8"/>
 *           │  @r:id — resolved through xl/_rels/workbook.xml.rels
 *           ▼
 *   xl/pivotCache/pivotCacheDefinition1.xml   <pivotCacheDefinition r:id="rId1" …>
 *           │  @r:id — resolved through *that part's own* .rels
 *           ▼
 *   xl/pivotCache/pivotCacheRecords1.xml      the cached source data
 *
 * ⚠️ `@cacheId` AND `@r:id` LIVE IN DIFFERENT IDENTIFIER SPACES. `pivotCache` carries
 * both, and both are required (verified: `CT_PivotCache` declares `:cacheId` and `r:id`,
 * each with a RequiredValidator, in the Open XML SDK schema). `cacheId` is matched by
 * *value* against `pivotTableDefinition/@cacheId`; `r:id` is resolved through the
 * relationship part. Code that treats `cacheId` as an rId, or looks up `r:id` values in
 * `pivotCaches`, finds nothing and usually reports the wrong part as missing. The
 * `pivotCaches` element is the only place the two spaces are bridged.
 *
 * ⚠️ WHY THE FAILURE IS INVISIBLE. Nothing in `pivotTable1.xml` looks wrong when the
 * chain breaks — every hop points *away* from it. And the cells the reader actually sees
 * are ordinary cells in the worksheet, carrying cached values written at the last
 * refresh. So the numbers still display, the file still opens, and the damage surfaces
 * only when someone clicks Refresh, months later.
 *
 * ⚠️ TWO DIFFERENT OWNERS. The pivot *table* part is related from the **worksheet** it
 * sits on; the *cache* is related from the **workbook**. So a converter walking one
 * worksheet's relationships never sees the cache, and a converter walking the workbook's
 * never sees the table. Worse, `CT_Worksheet`'s content model has no pivot child at all
 * (verified against the SDK schema): the worksheet→pivotTable link is a purely implicit
 * relationship with *nothing in the worksheet XML referring to it*, exactly like
 * slideLayout in PresentationML. Drop the relationship and the worksheet is still valid.
 *
 * INTEROPERABILITY. This brief's source reports [MS-OI29500] logging 67 normative
 * variations against Part 1 §18.10, second only to formulas — plausible given the size
 * of the clause, but **not verified here**: this module was written against the Open XML
 * SDK schema, and no MS-OI29500 text was consulted. Treat the number as hearsay. What
 * *is* verified below is marked as such; what is inferred says so.
 */

import { relsPathFor, resolveTarget, type PackageParts } from './packageIntegrity';
import { finding, renderFindings, type Finding, type Severity } from './findings';

/**
 * Namespace matching tolerates Strict as well as Transitional packages, which use
 * different URIs for the same vocabulary. Comparing whole URIs reports every Strict
 * package as broken.
 *
 * Note the two suffixes are genuinely different, not just a different host: Transitional
 * is `http://schemas.openxmlformats.org/spreadsheetml/2006/main` and Strict is
 * `http://purl.oclc.org/ooxml/spreadsheetml/main` — Strict **drops the year**. Matching
 * on `/spreadsheetml/2006/main` alone silently reads nothing out of a Strict workbook,
 * which looks exactly like a workbook with no pivot tables in it.
 */
const nsEndsWith = (uri: string | null, suffix: string) => uri !== null && uri.endsWith(suffix);

const isSpreadsheet = (el: Element) =>
  nsEndsWith(el.namespaceURI, '/spreadsheetml/2006/main') || nsEndsWith(el.namespaceURI, '/spreadsheetml/main');

/** Reads an `r:`-namespace attribute without pinning the exact relationships URI. */
const relAttr = (el: Element, local: string): string | null => {
  for (const a of Array.from(el.attributes)) {
    if (a.localName === local && a.namespaceURI?.includes('/relationships')) return a.value;
  }
  return null;
};

const descendants = (root: ParentNode, localName: string): Element[] =>
  Array.from(root.querySelectorAll('*')).filter(el => isSpreadsheet(el) && el.localName === localName);

const firstDescendant = (root: ParentNode, localName: string): Element | null =>
  descendants(root, localName)[0] ?? null;

/** OOXML booleans are `1`/`0` as often as `true`/`false`; both spellings are legal. */
const readBool = (el: Element | null, name: string): boolean | null => {
  const raw = el?.getAttribute(name);
  if (raw === null || raw === undefined) return null;
  return raw === '1' || raw === 'true';
};

/** The three links of the chain, named so a report can say *which one* broke. */
export type PivotChainHop =
  /** pivotTableDefinition/@cacheId → workbook.xml pivotCaches/pivotCache/@cacheId */
  | 'table-to-workbook'
  /** pivotCache/@r:id → workbook rels → pivotCacheDefinition part */
  | 'workbook-to-cache-definition'
  /** pivotCacheDefinition/@r:id → its own rels → pivotCacheRecords part */
  | 'cache-definition-to-records';

export type PivotProblemKind =
  | 'orphan-pivot-table-part'
  | 'missing-required-attribute'
  | 'workbook-missing'
  | 'no-pivot-caches'
  | 'cache-id-not-in-workbook'
  | 'duplicate-cache-id'
  | 'cache-relationship-missing'
  | 'cache-definition-missing'
  | 'cache-records-missing'
  | 'cache-records-absent'
  | 'no-cache-source'
  | 'field-index-out-of-range'
  | 'field-count-mismatch';

/**
 * Turns a pivot problem into a Finding.
 *
 * Severity and silence stay per-call-site here rather than moving to a table, because
 * unlike the other analyzers they are not a property of the kind: a missing records
 * part is an `error` normally and a `note` under `refreshOnLoad`, and only the call
 * site knows which. `hop` becomes part of the subject — it is the single most useful
 * thing to know about a broken chain, since it names where the walk stopped.
 */
const pivotFinding = (input: {
  kind: PivotProblemKind;
  /** Which hop this break sits on; null for problems inside a single part. */
  hop: PivotChainHop | null;
  severity: Severity;
  message: string;
  remediation: string;
  silent: boolean;
  /** The part the fault is in. Defaults to the pivot table part being walked. */
  part?: string;
}): Finding =>
  finding(`pivot/${input.kind}`, input.part ?? '', input.message, input.remediation, {
    severity: input.severity,
    silent: input.silent,
    ...(input.hop ? { subject: { hop: input.hop } } : {})
  });

/** Where a cache says its data came from. */
export interface PivotCacheSource {
  /** `worksheet` | `external` | `consolidation` | `scenario` — required, per the schema. */
  type: string | null;
  /** worksheetSource/@sheet — the sheet name, absent when a defined name is used. */
  sheet: string | null;
  /** worksheetSource/@ref — an A1-style range. */
  ref: string | null;
  /** worksheetSource/@name — a defined name, used *instead of* sheet+ref. */
  definedName: string | null;
  /** cacheSource/@connectionId — points into xl/connections.xml for external data. */
  connectionId: string | null;
  /** worksheetSource/@r:id — set when the source range lives in another workbook. */
  externalRelationshipId: string | null;
  /** One sentence a reader can act on. */
  description: string;
}

/** One index written by the pivot table that has to land inside the cache's fields. */
export interface PivotFieldReference {
  /** The element the index was read from: `rowFields`, `colFields`, `pageFields`, `dataFields`. */
  origin: 'rowFields' | 'colFields' | 'pageFields' | 'dataFields';
  index: number;
  /** The `cacheField/@name` the index lands on, or null when it lands nowhere. */
  cacheFieldName: string | null;
}

/**
 * The resolved (or broken) three-hop chain for one pivot table.
 *
 * Shaped after `resolveSlideChain` in `powerpointFormattingAnalysis`: every hop is
 * reported separately, because a break below the top is invisible from the top.
 */
export interface PivotCacheChain {
  pivotTablePath: string;
  /** The value of `pivotTableDefinition/@cacheId`, as written. */
  cacheId: string | null;
  workbookPath: string | null;
  /** `pivotCache/@r:id` — hop two's input, from a different id space than `cacheId`. */
  cacheRelationshipId: string | null;
  cacheDefinitionPath: string | null;
  cacheDefinition: Document | null;
  /** `pivotCacheDefinition/@r:id` — absent is legal; see `recordsAbsentIsExpected`. */
  cacheRecordsRelationshipId: string | null;
  cacheRecordsPath: string | null;
  /** null when no records relationship was declared at all. */
  cacheRecordsPresent: boolean | null;
  /** True when the cache says it refreshes on load, which makes absent records normal. */
  recordsAbsentIsExpected: boolean;
  /** The *first* hop that failed, or null when all three resolved. */
  brokenHop: PivotChainHop | null;
  problems: Finding[];
}

export interface PivotTable {
  partPath: string;
  /** The part that relates to this pivot table — a worksheet, in a well-formed package. */
  ownerPath: string | null;
  /** `pivotTableDefinition/@name` — required by the schema. */
  name: string | null;
  /** `location/@ref` — where on the sheet the pivot is drawn. */
  location: string | null;
  chain: PivotCacheChain;
  cacheSource: PivotCacheSource | null;
  /** Number of `cacheField` elements, or null when the cache did not resolve. */
  cacheFieldCount: number | null;
  /** Number of `pivotField` elements in this table. */
  pivotFieldCount: number;
  /** Every index this table writes into the cache's field list. */
  fieldReferences: PivotFieldReference[];
  problems: Finding[];
}

interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

const parseXml = (xml: string): Document | null => {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
};

const readRelationships = (parts: PackageParts, ownerPart: string): Map<string, Relationship> | null => {
  const relsXml = parts[relsPathFor(ownerPart)];
  if (relsXml === undefined) return null;
  const doc = parseXml(relsXml);
  if (!doc) return null;

  const map = new Map<string, Relationship>();
  for (const rel of Array.from(doc.getElementsByTagName('Relationship'))) {
    const id = rel.getAttribute('Id');
    if (id === null) continue;
    map.set(id, {
      id,
      type: rel.getAttribute('Type') ?? '',
      target: rel.getAttribute('Target') ?? '',
      external: rel.getAttribute('TargetMode') === 'External'
    });
  }
  return map;
};

/**
 * Finds the workbook part through the package relationships rather than assuming a path.
 *
 * `xl/workbook.xml` is a convention, not a rule — the part name comes from the
 * officeDocument relationship in `_rels/.rels`, and macro-enabled and Strict packages do
 * vary. Falls back to the conventional name only when the package relationships are
 * unreadable, so a package missing `_rels/.rels` is still analysable.
 */
const findWorkbookPath = (parts: PackageParts): string | null => {
  const rels = readRelationships(parts, '');
  if (rels) {
    for (const rel of rels.values()) {
      if (rel.type.endsWith('/officeDocument') && !rel.external) {
        const target = resolveTarget('', rel.target);
        if (parts[target] !== undefined) return target;
      }
    }
  }
  return parts['xl/workbook.xml'] !== undefined ? 'xl/workbook.xml' : null;
};

/**
 * Walks pivot table → workbook → cache definition → cache records, naming the hop that
 * broke.
 *
 * Every break here is silent: the pivot's rendered cells live in the worksheet and carry
 * cached values, so the workbook opens and displays identical numbers no matter which
 * hop is severed. Only Refresh reveals it.
 */
export const resolvePivotCacheChain = (
  parts: PackageParts,
  pivotTablePath: string,
  /** An already-parsed pivot table definition, when the caller has one. */
  parsedTable?: Document
): PivotCacheChain => {
  const problems: Finding[] = [];
  const chain: PivotCacheChain = {
    pivotTablePath,
    cacheId: null,
    workbookPath: null,
    cacheRelationshipId: null,
    cacheDefinitionPath: null,
    cacheDefinition: null,
    cacheRecordsRelationshipId: null,
    cacheRecordsPath: null,
    cacheRecordsPresent: null,
    recordsAbsentIsExpected: false,
    brokenHop: null,
    problems
  };

  const xml = parts[pivotTablePath];
  const table = parsedTable ?? (xml !== undefined ? parseXml(xml) : null);
  const definition = table?.documentElement ?? null;
  if (!definition) {
    chain.brokenHop = 'table-to-workbook';
    problems.push(pivotFinding({
      part: pivotTablePath,
      kind: 'missing-required-attribute',
      hop: 'table-to-workbook',
      severity: 'error',
      message: `${pivotTablePath} is missing or is not well-formed XML, so no pivot table can be read from it.`,
      remediation: `Restore ${pivotTablePath}.`,
      silent: true
    }));
    return chain;
  }

  // ---- hop one: @cacheId, matched by value inside the workbook -------------------
  chain.cacheId = definition.getAttribute('cacheId');
  if (chain.cacheId === null) {
    chain.brokenHop = 'table-to-workbook';
    problems.push(pivotFinding({
      part: pivotTablePath,
      kind: 'missing-required-attribute',
      hop: 'table-to-workbook',
      severity: 'error',
      message:
        'pivotTableDefinition/@cacheId is absent. It is required, and it is the only thing joining this pivot table to its cached data — without it the table names no cache at all.',
      remediation: 'Set cacheId to the cacheId of the workbook pivotCache holding this table’s data.',
      silent: true
    }));
    return chain;
  }

  const workbookPath = findWorkbookPath(parts);
  chain.workbookPath = workbookPath;
  const workbookDoc = workbookPath !== null ? parseXml(parts[workbookPath] ?? '') : null;
  if (!workbookDoc?.documentElement) {
    chain.brokenHop = 'table-to-workbook';
    problems.push(pivotFinding({
      part: workbookPath ?? 'xl/workbook.xml',
      kind: 'workbook-missing',
      hop: 'table-to-workbook',
      severity: 'error',
      message: `The workbook part${
        workbookPath === null ? '' : ` (${workbookPath})`
      } is missing or unparseable, so cacheId "${chain.cacheId}" cannot be looked up and no pivot cache can be found.`,
      remediation: 'Restore the workbook part named by the officeDocument relationship in _rels/.rels.',
      silent: true
    }));
    return chain;
  }

  const pivotCachesEl = firstDescendant(workbookDoc.documentElement, 'pivotCaches');
  const entries = pivotCachesEl ? descendants(pivotCachesEl, 'pivotCache') : [];
  if (entries.length === 0) {
    chain.brokenHop = 'table-to-workbook';
    problems.push(pivotFinding({
      part: workbookPath ?? 'xl/workbook.xml',
      kind: 'no-pivot-caches',
      hop: 'table-to-workbook',
      severity: 'error',
      message: `${workbookPath} declares no pivotCaches, so cacheId "${chain.cacheId}" resolves to nothing. pivotCaches is the only place a cacheId is bound to a relationship id, so with it gone every pivot table in the workbook is cut off from its data at once.`,
      remediation: `Add <pivotCaches><pivotCache cacheId="${chain.cacheId}" r:id="…"/></pivotCaches> to ${workbookPath}.`,
      silent: true
    }));
    return chain;
  }

  const matches = entries.filter(el => el.getAttribute('cacheId') === chain.cacheId);
  if (matches.length === 0) {
    chain.brokenHop = 'table-to-workbook';
    problems.push(pivotFinding({
      part: workbookPath ?? 'xl/workbook.xml',
      kind: 'cache-id-not-in-workbook',
      hop: 'table-to-workbook',
      severity: 'error',
      message: `No pivotCache in ${workbookPath} has cacheId="${
        chain.cacheId
      }" (the workbook offers ${entries
        .map(el => `"${el.getAttribute('cacheId') ?? '?'}"`)
        .join(', ')}). @cacheId is matched by value, not resolved as a relationship id, so an rId written here would also miss.`,
      remediation: `Add a pivotCache with cacheId="${chain.cacheId}", or point the pivot table at a cacheId the workbook actually declares.`,
      silent: true
    }));
    return chain;
  }
  if (matches.length > 1) {
    problems.push(pivotFinding({
      part: workbookPath ?? 'xl/workbook.xml',
      kind: 'duplicate-cache-id',
      hop: 'table-to-workbook',
      severity: 'error',
      message: `${workbookPath} declares ${matches.length} pivotCache entries with cacheId="${chain.cacheId}". cacheId is meant to identify one cache, so which data this pivot table shows is left to the consumer to guess.`,
      remediation: 'Give each pivotCache a distinct cacheId and repoint the pivot tables accordingly.',
      silent: true
    }));
  }

  // ---- hop two: @r:id, resolved through the workbook's own relationships ---------
  const cacheEntry = matches[0];
  chain.cacheRelationshipId = relAttr(cacheEntry, 'id');
  const workbookRels = readRelationships(parts, workbookPath as string);

  if (chain.cacheRelationshipId === null) {
    chain.brokenHop = 'workbook-to-cache-definition';
    problems.push(pivotFinding({
      part: pivotTablePath,
      kind: 'missing-required-attribute',
      hop: 'workbook-to-cache-definition',
      severity: 'error',
      message: `The pivotCache with cacheId="${chain.cacheId}" carries no r:id. Both attributes are required on pivotCache and they are not interchangeable: cacheId names the cache to pivot tables, r:id names the part to the package. Without r:id nothing points at a cache definition part.`,
      remediation: `Add r:id to the pivotCache and declare a matching relationship in ${relsPathFor(workbookPath as string)}.`,
      silent: true
    }));
    return chain;
  }

  const cacheRel = workbookRels?.get(chain.cacheRelationshipId);
  if (!cacheRel) {
    chain.brokenHop = 'workbook-to-cache-definition';
    problems.push(pivotFinding({
      part: workbookPath ?? 'xl/workbook.xml',
      kind: 'cache-relationship-missing',
      hop: 'workbook-to-cache-definition',
      severity: 'error',
      message: `Relationship "${chain.cacheRelationshipId}" is referenced by the pivotCache with cacheId="${
        chain.cacheId
      }" but ${
        workbookRels === null
          ? `${relsPathFor(workbookPath as string)} does not exist`
          : `is not declared in ${relsPathFor(workbookPath as string)}`
      }, so the cache definition part cannot be located.`,
      remediation: `Declare a pivotCacheDefinition relationship with Id="${chain.cacheRelationshipId}" in ${relsPathFor(
        workbookPath as string
      )}.`,
      silent: true
    }));
    return chain;
  }

  chain.cacheDefinitionPath = cacheRel.external
    ? cacheRel.target
    : resolveTarget(workbookPath as string, cacheRel.target);
  const cacheXml = cacheRel.external ? undefined : parts[chain.cacheDefinitionPath];
  chain.cacheDefinition = cacheXml !== undefined ? parseXml(cacheXml) : null;
  if (!chain.cacheDefinition?.documentElement) {
    chain.brokenHop = 'workbook-to-cache-definition';
    chain.cacheDefinition = null;
    problems.push(pivotFinding({
      part: chain.cacheDefinitionPath ?? pivotTablePath,
      kind: 'cache-definition-missing',
      hop: 'workbook-to-cache-definition',
      severity: 'error',
      message: `The pivot cache definition "${chain.cacheDefinitionPath}" is ${
        cacheRel.external ? 'declared TargetMode="External", so it is not in this package' : 'not in the package'
      } or is unparseable. The pivot table still renders from the cell values cached in the worksheet, so nothing looks wrong until someone refreshes it.`,
      remediation: `Restore ${chain.cacheDefinitionPath}, or remove the pivot table, its pivotCache entry and its relationship together.`,
      silent: true
    }));
    return chain;
  }

  // ---- hop three: the cache definition's own @r:id → the records part ------------
  const cacheRoot = chain.cacheDefinition.documentElement;
  chain.cacheRecordsRelationshipId = relAttr(cacheRoot, 'id');
  // `refreshOnLoad` means the cache is rebuilt from the source the moment the workbook
  // opens, so shipping no records at all is a normal, deliberate state — not damage.
  // `saveData="0"` says the same thing more bluntly: do not store the cached data.
  chain.recordsAbsentIsExpected =
    readBool(cacheRoot, 'refreshOnLoad') === true || readBool(cacheRoot, 'saveData') === false;

  if (chain.cacheRecordsRelationshipId === null) {
    // Not a break: r:id is optional on pivotCacheDefinition (verified — the SDK schema
    // declares it with no RequiredValidator, unlike the two on pivotCache).
    problems.push(pivotFinding({
      part: chain.cacheDefinitionPath ?? pivotTablePath,
      kind: 'cache-records-absent',
      hop: 'cache-definition-to-records',
      severity: chain.recordsAbsentIsExpected ? 'note' : 'error',
      message: chain.recordsAbsentIsExpected
        ? `${chain.cacheDefinitionPath} declares no pivotCacheRecords part, which is expected here: the cache is marked to refresh on load, so the records are rebuilt from the source rather than stored. This is a real state, not damage.`
        : `${chain.cacheDefinitionPath} declares no pivotCacheRecords part and is not marked refreshOnLoad, so the pivot table has no stored data and no instruction to rebuild it. Note that r:id is optional here, so this is not a schema violation — it is a workbook that will show an empty pivot on refresh.`,
      remediation: chain.recordsAbsentIsExpected
        ? 'None needed; refreshing the workbook regenerates the records.'
        : 'Add the pivotCacheRecords part and its relationship, or set refreshOnLoad="1" so the data is rebuilt on open.',
      silent: true
    }));
    return chain;
  }

  const cacheRels = readRelationships(parts, chain.cacheDefinitionPath);
  const recordsRel = cacheRels?.get(chain.cacheRecordsRelationshipId);
  if (!recordsRel) {
    chain.brokenHop = 'cache-definition-to-records';
    chain.cacheRecordsPresent = false;
    problems.push(pivotFinding({
      part: chain.cacheDefinitionPath ?? pivotTablePath,
      kind: 'cache-records-missing',
      hop: 'cache-definition-to-records',
      severity: 'error',
      message: `The cache definition references relationship "${chain.cacheRecordsRelationshipId}" for its records, but ${
        cacheRels === null
          ? `${relsPathFor(chain.cacheDefinitionPath)} does not exist`
          : `it is not declared in ${relsPathFor(chain.cacheDefinitionPath)}`
      }. Note this hop uses the cache definition's *own* relationship part, not the workbook's.`,
      remediation: `Declare a pivotCacheRecords relationship with Id="${chain.cacheRecordsRelationshipId}" in ${relsPathFor(
        chain.cacheDefinitionPath
      )}.`,
      silent: true
    }));
    return chain;
  }

  chain.cacheRecordsPath = recordsRel.external
    ? recordsRel.target
    : resolveTarget(chain.cacheDefinitionPath, recordsRel.target);
  chain.cacheRecordsPresent = recordsRel.external ? null : parts[chain.cacheRecordsPath] !== undefined;
  if (chain.cacheRecordsPresent === false) {
    chain.brokenHop = 'cache-definition-to-records';
    problems.push(pivotFinding({
      part: chain.cacheDefinitionPath ?? pivotTablePath,
      kind: 'cache-records-missing',
      hop: 'cache-definition-to-records',
      severity: 'error',
      message: `The cache definition points at records in "${chain.cacheRecordsPath}", which is not in the package. The cached values already written into the worksheet keep displaying, so the workbook looks and opens exactly as before${
        chain.recordsAbsentIsExpected ? '' : ' — and the cache is not marked refreshOnLoad, so nothing will rebuild them'
      }.`,
      remediation: `Restore ${chain.cacheRecordsPath}, or drop the r:id and set refreshOnLoad="1" so the data is rebuilt from the source instead.`,
      silent: true
    }));
  }

  return chain;
};

/** Reads `cacheSource` and says, in one sentence, where the data claims to come from. */
const readCacheSource = (cacheDefinition: Document, problems: Finding[], part: string): PivotCacheSource | null => {
  const root = cacheDefinition.documentElement;
  const el = root ? firstDescendant(root, 'cacheSource') : null;
  if (!el) {
    problems.push(pivotFinding({
      part: part,
      kind: 'no-cache-source',
      hop: null,
      severity: 'error',
      message:
        'The pivot cache definition has no cacheSource, so nothing records where the data came from. cacheSource is a required child, and without it the cache cannot be refreshed even when the source data is still present in the workbook.',
      remediation: 'Add a cacheSource naming the worksheet range or connection the data was read from.',
      silent: true
    }));
    return null;
  }

  const worksheetSource = firstDescendant(el, 'worksheetSource');
  const source: PivotCacheSource = {
    type: el.getAttribute('type'),
    sheet: worksheetSource?.getAttribute('sheet') ?? null,
    ref: worksheetSource?.getAttribute('ref') ?? null,
    definedName: worksheetSource?.getAttribute('name') ?? null,
    connectionId: el.getAttribute('connectionId'),
    externalRelationshipId: worksheetSource ? relAttr(worksheetSource, 'id') : null,
    description: ''
  };

  // @type is required and enumerated: worksheet | external | consolidation | scenario
  // (verified against ST_SourceType in the SDK schema).
  if (source.type === null) {
    problems.push(pivotFinding({
      part: part,
      kind: 'missing-required-attribute',
      hop: null,
      severity: 'error',
      message:
        'cacheSource/@type is absent. It is required, and it is what tells a consumer whether to look for a worksheet range, an external connection, consolidation ranges or a scenario — the child element alone does not settle it.',
      remediation: 'Set type to one of "worksheet", "external", "consolidation" or "scenario".',
      silent: true
    }));
  }

  if (source.type === 'worksheet') {
    const where = source.definedName !== null
      ? `the defined name "${source.definedName}"`
      : source.sheet !== null && source.ref !== null
        ? `${source.sheet}!${source.ref}`
        : source.ref !== null
          ? `the range ${source.ref}`
          : 'an unstated worksheet range';
    source.description =
      source.externalRelationshipId !== null
        ? `A worksheet range, ${where}, in another workbook reached through relationship "${source.externalRelationshipId}".`
        : `A worksheet range: ${where}.`;
  } else if (source.type === 'external') {
    source.description =
      source.connectionId !== null
        ? `An external data source, via connection ${source.connectionId} in xl/connections.xml.`
        : 'An external data source, but no connectionId is given, so which connection is left unsaid.';
  } else if (source.type === null) {
    source.description = 'The cache does not say what kind of source it came from.';
  } else {
    source.description = `A ${source.type} source.`;
  }

  return source;
};

const FIELD_INDEX_ORIGINS: { container: string; child: string; attribute: string; origin: PivotFieldReference['origin'] }[] = [
  // rowFields/colFields hold CT_Field, whose only attribute is the required @x.
  { container: 'rowFields', child: 'field', attribute: 'x', origin: 'rowFields' },
  { container: 'colFields', child: 'field', attribute: 'x', origin: 'colFields' },
  // pageFields and dataFields index with @fld instead — same idea, different spelling.
  { container: 'pageFields', child: 'pageField', attribute: 'fld', origin: 'pageFields' },
  { container: 'dataFields', child: 'dataField', attribute: 'fld', origin: 'dataFields' }
];

/**
 * Collects every field index the pivot table writes, and checks each lands in the cache.
 *
 * The bound is the number of `cacheField` elements. `pivotField` carries no index
 * attribute of its own (verified: `CT_PivotField` declares none) — its index *is* its
 * position in `pivotFields`, and that list is meant to run parallel to `cacheFields`.
 * **That parallelism is inferred, not schema-enforced**: nothing in the SDK schema ties
 * the two lists together, so this module treats the cache field count as the bound for
 * both and says so rather than pretending the schema settles it.
 *
 * `@x = -2` in rowFields/colFields is the conventional marker for the "values" pseudo
 * field rather than a real index. `@x` is a plain Int32 in the schema with no facet
 * restricting it, so **this convention is not schema-verified**; it is excluded from the
 * range check because treating it as a real index would report a false dangling
 * reference on most multi-measure pivots.
 */
const readFieldReferences = (
  definition: Element,
  cacheFieldNames: string[] | null,
  problems: Finding[],
  part: string
): PivotFieldReference[] => {
  const references: PivotFieldReference[] = [];

  for (const { container, child, attribute, origin } of FIELD_INDEX_ORIGINS) {
    const containerEl = firstDescendant(definition, container);
    if (!containerEl) continue;
    for (const el of descendants(containerEl, child)) {
      const raw = el.getAttribute(attribute);
      if (raw === null) continue;
      const index = Number.parseInt(raw, 10);
      if (Number.isNaN(index)) continue;
      if (index < 0) continue; // -2: the "values" pseudo field; see the doc comment.

      const name = cacheFieldNames !== null && index < cacheFieldNames.length ? cacheFieldNames[index] : null;
      references.push({ origin, index, cacheFieldName: name });

      if (cacheFieldNames !== null && index >= cacheFieldNames.length) {
        problems.push(pivotFinding({
          part: part,
          kind: 'field-index-out-of-range',
          hop: null,
          severity: 'error',
          message: `${container}/${child}/@${attribute} is ${index}, but the cache defines only ${cacheFieldNames.length} field${
            cacheFieldNames.length === 1 ? '' : 's'
          } (indices 0–${cacheFieldNames.length - 1}). Fields are referenced by position, so an index past the end is a dangling reference: the column comes out blank, or shows the wrong field once the indices shift.`,
          remediation: `Point the reference at an existing cache field index, or add the missing cacheField to the cache definition.`,
          silent: true
        }));
      }
    }
  }

  return references;
};

/** Reads one pivot table part in full, chain included. */
const readPivotTable = (parts: PackageParts, partPath: string, ownerPath: string | null): PivotTable => {
  const problems: Finding[] = [];
  const doc = parseXml(parts[partPath] ?? '');
  const definition = doc?.documentElement ?? null;
  const chain = resolvePivotCacheChain(parts, partPath, doc ?? undefined);

  if (!definition) {
    return {
      partPath,
      ownerPath,
      name: null,
      location: null,
      chain,
      cacheSource: null,
      cacheFieldCount: null,
      pivotFieldCount: 0,
      fieldReferences: [],
      problems
    };
  }

  if (ownerPath === null) {
    problems.push(pivotFinding({
      part: partPath,
      kind: 'orphan-pivot-table-part',
      hop: null,
      severity: 'error',
      message: `${partPath} is in the package but no part relates to it. A pivot table is reached only through an implicit relationship from the worksheet it sits on — nothing in worksheet XML names it — so an unreferenced pivot table part is dead weight: it never appears, and the worksheet stays perfectly valid without it.`,
      remediation: `Add a pivotTable relationship from the worksheet that shows this pivot, or delete ${partPath}.`,
      silent: true
    }));
  }

  // @name and @dataCaption are required on pivotTableDefinition (verified against the
  // SDK schema), alongside @cacheId which resolvePivotCacheChain reports on.
  for (const attribute of ['name', 'dataCaption']) {
    if (definition.getAttribute(attribute) === null) {
      problems.push(pivotFinding({
        part: partPath,
        kind: 'missing-required-attribute',
        hop: null,
        severity: 'error',
        message: `pivotTableDefinition/@${attribute} is absent, but the schema requires it.`,
        remediation: `Set @${attribute} on the pivotTableDefinition.`,
        silent: true
      }));
    }
  }

  const cacheSource = chain.cacheDefinition ? readCacheSource(chain.cacheDefinition, problems, chain.cacheDefinitionPath ?? partPath) : null;

  let cacheFieldNames: string[] | null = null;
  if (chain.cacheDefinition?.documentElement) {
    const cacheFieldsEl = firstDescendant(chain.cacheDefinition.documentElement, 'cacheFields');
    const cacheFields = cacheFieldsEl ? descendants(cacheFieldsEl, 'cacheField') : [];
    cacheFieldNames = cacheFields.map(el => el.getAttribute('name') ?? '');

    // @count is advisory — the elements are the truth — but a consumer that sizes an
    // array from @count and then reads by index gets the wrong field or none.
    const declared = cacheFieldsEl?.getAttribute('count');
    if (declared !== null && declared !== undefined && Number.parseInt(declared, 10) !== cacheFields.length) {
      problems.push(pivotFinding({
        part: partPath,
        kind: 'field-count-mismatch',
        hop: null,
        severity: 'error',
        message: `cacheFields/@count says ${declared} but ${cacheFields.length} cacheField elements are present. Consumers that size their field table from @count read the wrong field, or none, for every index past the shorter of the two.`,
        remediation: `Set cacheFields/@count to ${cacheFields.length}.`,
        silent: true
      }));
    }
  }

  const pivotFieldsEl = firstDescendant(definition, 'pivotFields');
  const pivotFields = pivotFieldsEl ? descendants(pivotFieldsEl, 'pivotField') : [];

  if (cacheFieldNames !== null && pivotFields.length > 0 && pivotFields.length !== cacheFieldNames.length) {
    problems.push(pivotFinding({
      part: partPath,
      kind: 'field-count-mismatch',
      hop: null,
      severity: 'error',
      message: `The table declares ${pivotFields.length} pivotField element${
        pivotFields.length === 1 ? '' : 's'
      } but the cache defines ${cacheFieldNames.length} cacheField${
        cacheFieldNames.length === 1 ? '' : 's'
      }. pivotField has no index attribute — a pivot field *is* its position in the list — so once the two lists differ in length every index past the difference names a different field than the author meant.`,
      remediation: 'Write one pivotField per cacheField, in the same order.',
      silent: true
    }));
  }

  const fieldReferences = readFieldReferences(definition, cacheFieldNames, problems, partPath);

  return {
    partPath,
    ownerPath,
    name: definition.getAttribute('name'),
    location: firstDescendant(definition, 'location')?.getAttribute('ref') ?? null,
    chain,
    cacheSource,
    cacheFieldCount: cacheFieldNames?.length ?? null,
    pivotFieldCount: pivotFields.length,
    fieldReferences,
    problems
  };
};

/**
 * Every pivot table in the package, found the way a consumer finds them.
 *
 * Discovery runs through the relationships, since that is the only link there is: a
 * `pivotTable` relationship from a worksheet. Parts under `xl/pivotTables/` that no
 * relationship points at are picked up too and reported as orphans — that path is the
 * usual convention rather than a requirement, so a producer using another naming scheme
 * would hide an orphan from this scan. Relationship-discovered tables are found wherever
 * they live.
 */
export function readPivotTables(parts: PackageParts): PivotTable[] {
  const owners = new Map<string, string>();

  for (const path of Object.keys(parts)) {
    if (!path.endsWith('.rels')) continue;
    const ownerPart = path.replace(/(^|\/)_rels\/([^/]+)\.rels$/, '$1$2');
    if (ownerPart === path) continue;
    const doc = parseXml(parts[path]);
    if (!doc) continue;
    for (const rel of Array.from(doc.getElementsByTagName('Relationship'))) {
      if (!(rel.getAttribute('Type') ?? '').endsWith('/pivotTable')) continue;
      if (rel.getAttribute('TargetMode') === 'External') continue;
      const target = resolveTarget(ownerPart, rel.getAttribute('Target') ?? '');
      if (!owners.has(target)) owners.set(target, ownerPart);
    }
  }

  const byConvention = Object.keys(parts).filter(path => /^xl\/pivotTables\/[^/]+\.xml$/.test(path));
  const found = [...new Set([...owners.keys(), ...byConvention])].sort();

  return found
    .filter(path => parts[path] !== undefined)
    .map(path => readPivotTable(parts, path, owners.get(path) ?? null));
}

/**
 * Everything wrong with a pivot table, chain problems and part problems together.
 *
 * `note`-severity entries are excluded: an absent records part under `refreshOnLoad` is
 * a design decision, and listing it beside genuine breakage would train a reader to
 * ignore the list.
 */
export function pivotTableErrors(table: PivotTable): Finding[] {
  return [...table.chain.problems, ...table.problems].filter(p => p.severity === 'error');
}

/**
 * The pivot tables that display correctly and are broken anyway.
 *
 * This is the whole list, in practice: a severed chain never changes what the workbook
 * shows, because the visible cells are ordinary worksheet cells holding the values from
 * the last refresh. No visual diff and no rendering test catches any of these.
 */
export function findSilentlyBrokenPivotTables(tables: PivotTable[]): PivotTable[] {
  return tables.filter(table => pivotTableErrors(table).some(p => p.silent));
}

/**
 * One line naming the hop that broke, for a report that has room for a sentence.
 *
 * Returns null when the chain is whole — including when the records part is absent by
 * design, which is not a break.
 */
export function describeBrokenHop(chain: PivotCacheChain): string | null {
  switch (chain.brokenHop) {
    case 'table-to-workbook':
      return `Hop 1 of 3 broke: pivotTableDefinition/@cacheId="${
        chain.cacheId ?? '(absent)'
      }" does not reach a pivotCache entry in ${chain.workbookPath ?? 'the workbook'}.`;
    case 'workbook-to-cache-definition':
      return `Hop 2 of 3 broke: the workbook's pivotCache r:id="${
        chain.cacheRelationshipId ?? '(absent)'
      }" does not reach a pivot cache definition part.`;
    case 'cache-definition-to-records':
      return `Hop 3 of 3 broke: ${
        chain.cacheDefinitionPath ?? 'the cache definition'
      } references records at r:id="${chain.cacheRecordsRelationshipId ?? '(absent)'}" that are not in the package.`;
    default:
      return null;
  }
}

/**
 * Evidence lines for the AI panel.
 *
 * Needs the workbook and the cache parts alongside the worksheet, because every hop of
 * the chain lives somewhere else. A pivot table part read on its own always looks fine.
 */
export function computePivotEvidenceForMarkup(
  parts: Record<string, string>
): { lines: string[]; unresolved: string[] } | null {
  const tables = readPivotTables(parts);
  if (tables.length === 0) return null;

  const lines: string[] = [];
  const unresolved: string[] = [];

  lines.push(`${tables.length} pivot table(s) in this workbook.`);

  for (const table of tables) {
    lines.push(
      `Pivot table "${table.name ?? 'unnamed'}" (${table.partPath}) uses cacheId ${table.chain.cacheId ?? 'none'}${
        table.chain.cacheDefinitionPath ? `, whose definition is ${table.chain.cacheDefinitionPath}` : ''
      }.`
    );
    const broken = describeBrokenHop(table.chain);
    if (broken) lines.push(broken);
    lines.push(...renderFindings([...table.chain.problems, ...table.problems]));
  }

  const silent = findSilentlyBrokenPivotTables(tables);
  if (silent.length > 0) {
    lines.push(
      `${silent.length} of these pivot table(s) still display the values from their last refresh, so the workbook looks correct and no visual check will catch the break. It surfaces on the next refresh.`
    );
  }

  // The cache records hold what the pivot was built from; whether they still agree with
  // the live source range is not knowable without recomputing the source.
  if (tables.some(t => t.chain.cacheRecordsPath !== null)) {
    unresolved.push(
      'Cached pivot records were located but not compared against their source range, so whether the cache is stale is unverified.'
    );
  }

  return { lines, unresolved };
}
