/**
 * External workbook links — the numbers are real, and nobody knows how old they are.
 *
 * `excelFormulas.ts` detects the reference: a formula containing `[1]Sheet1!A1` points at
 * a *different workbook*, and that module correctly refuses to say more, because the
 * source is outside the package. This module is the other half — it resolves what `[1]`
 * actually names, and reports what the package does and does not know about it.
 *
 * WHY THIS IS WORTH AN ANALYZER.
 *
 *   A WORKBOOK THAT LINKS TO ANOTHER WORKBOOK CACHES THE VALUES IT LAST READ.
 *
 * Open it somewhere the source is unreachable — a different machine, a moved share, a
 * colleague's laptop — and every linked cell keeps showing the numbers from whenever the
 * two files were last open together. Nothing is blank, nothing is red, no dialog appears
 * about the data itself. The spreadsheet looks completely normal and its numbers are from
 * an unknown date. That is the same failure class as a stale formula cache, except the
 * missing input is a file that may not exist anywhere any more.
 *
 * THE CHAIN, AND EVERY PLACE IT BREAKS.
 *
 *   xl/workbook.xml   <externalReferences><externalReference r:id="rId5"/></externalReferences>
 *           │  r:id resolved through xl/_rels/workbook.xml.rels
 *           ▼
 *   xl/externalLinks/externalLink1.xml   <externalLink><externalBook r:id="rId1">
 *           │                              <sheetNames>, <sheetDataSet> ← the cached values
 *           │  r:id resolved through xl/externalLinks/_rels/externalLink1.xml.rels
 *           ▼
 *     TargetMode="External" → the source workbook, outside this package
 *
 * The last hop leaves the package, and this module **never follows it**. An external
 * target is reported as *unverifiable*, never as broken — the same three-state discipline
 * as `oleDataIsPresent` and `hyperlinkResolves`: `null` means "cannot be checked from
 * here", which is not `false`. Nothing is fetched.
 *
 * Every hop before that one is fully checkable, and each break leaves a workbook that
 * opens and shows its numbers exactly as before.
 *
 * PROVENANCE — verified, and against what.
 *
 * Verified against the Open XML SDK schema data
 * (`data/schemas/schemas_openxmlformats_org_spreadsheetml_2006_main.json`):
 *   • `x:CT_ExternalReference/x:externalReference` is a LEAF with exactly one attribute,
 *     `r:id`, and it is **required**. Its only parent is `x:externalReferences`, whose
 *     only parent is `x:workbook` (min 1, max 65534 children).
 *   • `x:CT_ExternalLink/x:externalLink` is the root of the linked part and its content is
 *     a **choice** of `x:externalBook`, `x:ddeLink` or `x:oleLink`, followed by an
 *     optional `x:extLst`. **An external-link part is not necessarily a workbook link** —
 *     a DDE conversation and an OLE link use the same part, so code that assumes
 *     `externalBook` mis-reports both.
 *   • `x:CT_ExternalBook/x:externalBook` has one attribute, `r:id`, **required**. Its
 *     children, all optional and all at most one: `xxl21:alternateUrls` (Microsoft 365
 *     extension), `x:sheetNames`, `x:definedNames`, `x:sheetDataSet`.
 *   • `x:CT_ExternalSheetNames/x:sheetNames` requires **at least one** `x:sheetName`
 *     (max 65534); `x:CT_ExternalSheetName/x:sheetName/@val` is an optional string.
 *   • `x:CT_ExternalSheetDataSet/x:sheetDataSet` requires at least one `x:sheetData`.
 *     `x:CT_ExternalSheetData/x:sheetData` carries `@sheetId` (UInt32, **required**) and
 *     `@refreshError` (boolean, "Last Refresh Resulted in Error").
 *   • `x:CT_ExternalRow/x:row/@r` required; `x:CT_ExternalCell/x:cell/@r` required, with
 *     `@t`, `@vm` and a single `x:v` child.
 *   • `x:CT_ExternalDefinedName/x:definedName` — `@name` **required**, `@refersTo`,
 *     `@sheetId`. Note this is a *different type* from the workbook-level
 *     `x:CT_DefinedName/x:definedName`, which is a leaf-TEXT element (its formula is
 *     element content, not an attribute) and carries fifteen attributes including
 *     `@localSheetId` and `@hidden`. Same element name, two unrelated types.
 *
 * Verified against the SDK's part metadata (`data/parts/ExternalWorkbookPart.json`):
 *   • relationship type `…/officeDocument/2006/relationships/externalLink`
 *   • content type `application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml`
 *   • ⚠️ and a genuine divergence: the SDK's own default target is
 *     **`externalReferences/externalReference1.xml`**, while Excel writes
 *     **`externalLinks/externalLink1.xml`**. Part names are arbitrary in OPC — the
 *     relationship is what binds the part — so both spellings are accepted here and
 *     nothing is resolved by path.
 *
 * `[N]` — VERIFIED, BUT NOT FROM THE SCHEMA.
 *
 * No schema describes formula text, so the meaning of the bracketed index was confirmed
 * against LibreOffice's OOXML importer, in two independent places:
 *   • `sc/source/filter/oox/externallinkbuffer.cxx`, `ExternalLinkBuffer::getExternalLink`:
 *     "OOXML: 0 = this document, otherwise one-based index into link list", implemented as
 *     `maLinks.get(nRefId - 1)` over the links accumulated from `x:externalReference`
 *     elements **in document order**.
 *   • `sc/source/core/tool/compiler.cxx`, `ConventionXL_OOX::makeExternalRefStr`:
 *     "`[N]SheetName!$A$1` … where N is a 1-based positive integer number of a file name
 *     in OOXML xl/externalLinks/externalLinkN.xml".
 *
 * So the brief's claim is right — **but incomplete in one way that matters**: `[0]` is a
 * legal index meaning *this workbook*, not an out-of-range one. `compiler.cxx` skips it
 * explicitly ("`[0]!Global_Range_Name` is a special case in OOXML syntax, where the '0' is
 * referencing to self"). Reporting `[0]` as dangling would be a false positive on every
 * workbook that uses a global defined name in that form, and it is never reported — see
 * `formulaIndexFindings` for why that needs no code of its own.
 *
 * The two LibreOffice sources describe the same index slightly differently — one as an
 * ordinal into the `externalReference` list, the other as the digit in the part's file
 * name. They coincide in every file Excel writes. **The ordinal is what is used here**,
 * because it is the only reading that survives a package whose parts are named by the SDK
 * convention above; and it is the reading the code in `getExternalLink` actually
 * implements.
 *
 * ⚠️ `sheetData/@sheetId` — NOT confirmed by the schema, which only types it `UInt32`.
 * LibreOffice's `externallinkfragment.cxx` resolves it as a **zero-based index into the
 * `sheetNames` list** (`getSheetCache(sheetId)` → `getVectorElement(maSheetCaches, nTabId)`,
 * where the cache vector is filled from `sheetNames` in order). That reading is used, and
 * a `@sheetId` past the end of `sheetNames` is reported as a warning rather than an error
 * because the source is one implementation rather than the specification.
 *
 * ⚠️ NOT VERIFIED — what Excel *displays* when a link cannot be resolved and no cached
 * value exists for it, and whether Excel offers to repair a workbook whose external-link
 * part is missing. Nothing here claims either. Every finding is phrased in terms of what
 * the package does and does not contain.
 *
 * Namespaces are compared by exact equality against `S_NAMESPACE`: `conformance.ts`
 * rewrites ISO Strict URIs to Transitional before any analyzer runs, so Strict-tolerant
 * matching here would be dead code that looks load-bearing.
 */

import { S_NAMESPACE } from './excelStyleResolver';
import { relsPathFor, resolveTarget, type PackageParts } from './packageIntegrity';
import { finding, renderFindings, type Finding, type Severity } from './findings';

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Verified against `data/parts/ExternalWorkbookPart.json`. */
export const EXTERNAL_LINK_RELATIONSHIP_TYPE = `${R_NS}/externalLink`;

/** The one part that lists external references, and the only entry point to the chain. */
export const WORKBOOK_PART = 'xl/workbook.xml';

/**
 * External-link parts, under either naming convention.
 *
 * `externalLinks/externalLink1.xml` is Excel's; `externalReferences/externalReference1.xml`
 * is the Open XML SDK's. Nothing is *resolved* by this pattern — every reference is
 * followed through its relationship — it exists only to find parts that no reference
 * reaches, which is a question about the package's contents rather than about a link.
 * `[^/]+` keeps `_rels/…` subdirectories out.
 */
export const EXTERNAL_LINK_PART = /^xl\/external(?:Links|References)\/[^/]+\.xml$/;

/** Parts this analyzer has something to say about when one is open in the editor. */
export const EXTERNAL_LINK_HOST_PART = /^xl\/(?:workbook\.xml|external(?:Links|References)\/[^/]+\.xml)$/;

/** Worksheets, whose formulas carry the `[N]` indexes that point back into the list. */
export const EXTERNAL_LINK_FORMULA_PART = /^xl\/worksheets\/[^/]+\.xml$/;

/**
 * Severity and silence per kind.
 *
 * **Every one of these is silent**, and that is the whole thesis of the analyzer: a
 * linked cell's value is cached in the worksheet that references it, so the workbook
 * opens and displays exactly the same numbers whether the chain behind them is intact or
 * completely severed. There is nothing to see, in any of these cases.
 *
 * `silent` is a claim about what renders, and it is made deliberately: the values on the
 * sheet come from `x:c/x:v` in the worksheet part, which none of these faults touches.
 * It is *not* a claim that Excel opens without a repair prompt — that was not verified,
 * and no finding here says anything about a prompt.
 *
 * `cached-sheet-id-unknown` is a warning rather than an error because its rule comes from
 * one implementation rather than from the specification (see the header).
 * `unreferenced-link-part` is a note: dead weight, not damage.
 */
const EXTERNAL_LINK_RULES = {
  'reference-no-relationship-id': { severity: 'error',   silent: true },
  'reference-relationship-missing': { severity: 'error', silent: true },
  'link-part-missing':            { severity: 'error',   silent: true },
  'link-part-unreadable':         { severity: 'error',   silent: true },
  'book-no-relationship-id':      { severity: 'error',   silent: true },
  'book-relationship-missing':    { severity: 'error',   silent: true },
  'source-part-missing':          { severity: 'error',   silent: true },
  'formula-index-unresolved':     { severity: 'error',   silent: true },
  'link-kind-unknown':            { severity: 'warning', silent: true },
  'source-not-external':          { severity: 'warning', silent: true },
  'no-cached-values':             { severity: 'warning', silent: true },
  'refresh-error':                { severity: 'warning', silent: true },
  'empty-sheet-names':            { severity: 'warning', silent: true },
  'cached-sheet-id-unknown':      { severity: 'warning', silent: true },
  'unreferenced-link-part':       { severity: 'note',    silent: true }
} as const satisfies Record<string, { severity: Severity; silent: boolean }>;

export type ExternalLinkProblemKind = keyof typeof EXTERNAL_LINK_RULES;

const externalLinkFinding = (
  kind: ExternalLinkProblemKind,
  part: string,
  message: string,
  remediation: string,
  subject?: Record<string, string>
): Finding =>
  finding(`externalLink/${kind}`, part, message, remediation, { ...EXTERNAL_LINK_RULES[kind], subject });

/**
 * What the linked part turned out to contain.
 *
 * `externalBook` is the workbook link everyone means by "external link"; the other two are
 * legal contents of the same part and are reported as what they are rather than as a
 * malformed workbook link. `unreadable` and `unknown` are distinct: one is a part that did
 * not parse, the other a part that parsed and declared none of the three.
 */
export type ExternalLinkKind = 'externalBook' | 'ddeLink' | 'oleLink' | 'unknown' | 'unreadable';

/** One cached sheet from the source workbook. */
export interface CachedSheet {
  /** `@sheetId` verbatim — a zero-based index into `sheetNames` (see header). */
  sheetId: string | null;
  /** The name that index lands on, or null when it lands outside the list. */
  sheetName: string | null;
  /** `@refreshError` — Excel recorded that the last refresh of this sheet failed. */
  refreshError: boolean;
  rowCount: number;
  cellCount: number;
}

export interface ExternalBook {
  /** `@r:id`. `''` is preserved as distinct from absent. */
  relationshipId: string | null;
  /** The source workbook's location as the relationship states it. **Never fetched.** */
  target: string | null;
  /** True when the relationship carries `TargetMode="External"`, as it should. */
  targetIsExternal: boolean;
  /**
   * Whether the source workbook is present. THREE STATES, and the third is the point.
   *
   * `null`  — CANNOT BE CHECKED from the package. This is the normal, healthy answer for
   *           a working external link: the source is another file on someone's disk or
   *           share, and a zip cannot answer questions about it.
   * `false` — the relationship pointed *inside* the package and the part is not there.
   * `true`  — the relationship pointed inside the package and the part is there.
   */
  sourceIsPresent: boolean | null;
  /** `sheetNames/sheetName/@val`, in order. The index `sheetData/@sheetId` counts into. */
  sheetNames: string[];
  /** `definedNames/definedName/@name`, in order. */
  definedNames: string[];
  /** null when there is no `sheetDataSet` element at all — no cache, not an empty one. */
  cachedSheets: CachedSheet[] | null;
}

export interface ExternalReference {
  /** 1-based position in `externalReferences` — the number a formula writes as `[N]`. */
  index: number;
  /** `@r:id` on `x:externalReference`. Required by the schema; `null` when absent. */
  relationshipId: string | null;
  /** The resolved external-link part path, or null when the relationship did not resolve. */
  partPath: string | null;
  kind: ExternalLinkKind | null;
  /** Present only when `kind` is `externalBook`. */
  book: ExternalBook | null;
  problems: Finding[];
}

export interface ExternalLinkSet {
  /** In document order. Position + 1 is the `[N]` a formula uses. */
  references: ExternalReference[];
  /** External-link parts in the package that no `externalReference` resolves to. */
  unreferencedParts: string[];
  /**
   * False when `xl/workbook.xml` is absent or unparseable.
   *
   * Nothing can be resolved without it — not the list, not its order, and therefore not
   * any `[N]` in any formula. Reported so callers can say "not checked" rather than "none".
   */
  workbookRead: boolean;
  /** Findings about the set as a whole rather than about one reference. */
  problems: Finding[];
}

interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

const parseXml = (xml: string | undefined): Document | null => {
  if (xml === undefined) return null;
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
};

const isX = (el: Element, local: string): boolean => el.namespaceURI === S_NAMESPACE && el.localName === local;

const childrenNamed = (parent: Element, local: string): Element[] =>
  Array.from(parent.children).filter(child => isX(child, local));

const firstChild = (parent: Element, local: string): Element | null => childrenNamed(parent, local)[0] ?? null;

/**
 * `@r:id` as written, preserving the difference between absent (`null`) and empty (`''`).
 * `?? null` on an empty string would collapse the two, so it is spelled out.
 */
const relId = (el: Element): string | null => el.getAttributeNS(R_NS, 'id');

const isOn = (value: string | null): boolean => value === '1' || value === 'true';

const readRelationships = (parts: PackageParts, ownerPart: string): Map<string, Relationship> | null => {
  const doc = parseXml(parts[relsPathFor(ownerPart)]);
  if (!doc) return null;

  const map = new Map<string, Relationship>();
  for (const rel of Array.from(doc.getElementsByTagName('Relationship'))) {
    const id = rel.getAttribute('Id');
    if (id === null) continue;
    map.set(id, {
      id,
      type: rel.getAttribute('Type') ?? '',
      target: rel.getAttribute('Target') ?? '',
      external: (rel.getAttribute('TargetMode') ?? '').toLowerCase() === 'external'
    });
  }
  return map;
};

/**
 * The source workbook, as far as the package can describe it.
 *
 * This is the hop that leaves the package, and the one place the analyzer has to be most
 * careful about what it claims. A well-formed external link ends here with
 * `sourceIsPresent === null` and **no finding at all** — "the source is somewhere else"
 * is the definition of an external link, not a defect. The findings below are about the
 * packaging around that target: a relationship that is not declared, or one that does not
 * say the target is external and so resolves to a package path that is not there.
 */
const readBook = (
  book: Element,
  parts: PackageParts,
  linkPart: string,
  problems: Finding[]
): ExternalBook => {
  const id = relId(book);
  const rels = readRelationships(parts, linkPart);

  let target: string | null = null;
  let targetIsExternal = false;
  let sourceIsPresent: boolean | null = null;

  if (id === null || id === '') {
    problems.push(externalLinkFinding(
      'book-no-relationship-id', linkPart,
      `${linkPart} declares an externalBook with no r:id, which the schema makes required. Nothing in the package names the workbook this link reads from, so not even its path is recoverable — the cached values below are attributed to a source that cannot be identified.`,
      'Add an r:id to the externalBook naming a TargetMode="External" relationship that points at the source workbook.'
    ));
  } else {
    const rel = rels?.get(id);
    if (rel === undefined) {
      problems.push(externalLinkFinding(
        'book-relationship-missing', linkPart,
        rels === null
          ? `${linkPart} references relationship "${id}" for its source workbook, but ${relsPathFor(linkPart)} does not exist, so even the path to the source workbook is unknown. The values cached in this part still display in the workbook and nothing indicates where they came from.`
          : `${linkPart} references relationship "${id}" for its source workbook, and ${relsPathFor(linkPart)} does not declare it. Even the path to the source is unknown — the link cannot be repointed or refreshed, only rebuilt.`,
        `Add a Relationship with Id="${id}" and TargetMode="External" to ${relsPathFor(linkPart)}, pointing at the source workbook.`,
        { relationshipId: id }
      ));
    } else if (rel.external) {
      // The healthy case. Reported, never fetched, never judged.
      target = rel.target;
      targetIsExternal = true;
    } else {
      // No TargetMode="External", so OPC resolves the Target as a path inside this
      // package — which is not what an external workbook link means.
      target = resolveTarget(linkPart, rel.target);
      sourceIsPresent = parts[target] !== undefined;
      if (sourceIsPresent) {
        problems.push(externalLinkFinding(
          'source-not-external', linkPart,
          `The externalBook in ${linkPart} points at relationship "${id}", which carries no TargetMode="External". OPC therefore resolves it as the package path "${target}" rather than as a location outside the file. A workbook link names another file; a relationship without TargetMode does not.`,
          `Add TargetMode="External" to relationship "${id}" in ${relsPathFor(linkPart)} if it names a workbook outside this package.`,
          { relationshipId: id, target }
        ));
      } else {
        problems.push(externalLinkFinding(
          'source-part-missing', linkPart,
          `The externalBook in ${linkPart} points at relationship "${id}", which carries no TargetMode="External" and so resolves to the package path "${target}" — and no such part is in the package. The link resolves to nothing, while every cell that reads from it keeps displaying its cached number.`,
          `Add TargetMode="External" to relationship "${id}" if the target is a workbook outside this package, or restore ${target}.`,
          { relationshipId: id, target }
        ));
      }
    }
  }

  const sheetNamesElement = firstChild(book, 'sheetNames');
  const sheetNames = sheetNamesElement === null
    ? []
    : childrenNamed(sheetNamesElement, 'sheetName').map(el => el.getAttribute('val') ?? '');

  if (sheetNamesElement !== null && sheetNames.length === 0) {
    problems.push(externalLinkFinding(
      'empty-sheet-names', linkPart,
      `${linkPart} declares a sheetNames element with no sheetName inside it; the schema requires at least one. Nothing records which sheets of the source workbook this link reads, so a cached sheet cannot be matched to a name.`,
      'Write one sheetName per sheet of the source workbook that this link reads from, or remove the empty sheetNames element.'
    ));
  }

  const definedNamesElement = firstChild(book, 'definedNames');
  const definedNames = definedNamesElement === null
    ? []
    : childrenNamed(definedNamesElement, 'definedName').map(el => el.getAttribute('name') ?? '');

  const dataSet = firstChild(book, 'sheetDataSet');
  const cachedSheets = dataSet === null ? null : childrenNamed(dataSet, 'sheetData').map(sheetData => {
    const rows = childrenNamed(sheetData, 'row');
    const sheetId = sheetData.getAttribute('sheetId');
    // Zero-based into sheetNames — LibreOffice's reading, not the schema's; see header.
    //
    // Matched against `\d+` rather than passed to `Number`, which is far too generous
    // here: `Number('')` is 0 and `Number(' ')` is 0, so `sheetId=""` would silently
    // attribute a sheet's worth of cached values to the FIRST sheet in the list and
    // report nothing. The schema types @sheetId as a required UInt32; anything that is
    // not a run of digits does not name a position, and saying so is the honest answer.
    const position = /^\d+$/.test(sheetId ?? '') ? Number(sheetId) : Number.NaN;
    return {
      sheetId,
      sheetName: position < sheetNames.length ? sheetNames[position] : null,
      refreshError: isOn(sheetData.getAttribute('refreshError')),
      rowCount: rows.length,
      cellCount: rows.reduce((total, row) => total + childrenNamed(row, 'cell').length, 0)
    };
  });

  if (cachedSheets === null) {
    problems.push(externalLinkFinding(
      'no-cached-values', linkPart,
      `${linkPart} carries no sheetDataSet, so the package holds none of the source workbook's values. A cell referencing this link still shows the number cached in its own worksheet, and there is nothing here to check that number against or to fall back on when the source is unreachable.`,
      'Open the workbook in Excel with the source available and save it, so the cached values are written back, or replace the links with static values.'
    ));
  }

  for (const sheet of cachedSheets ?? []) {
    if (sheet.refreshError) {
      problems.push(externalLinkFinding(
        'refresh-error', linkPart,
        `The cached sheet ${sheet.sheetName === null ? `with sheetId "${sheet.sheetId ?? 'absent'}"` : `"${sheet.sheetName}"`} in ${linkPart} is marked refreshError, meaning the last attempt to read it from the source workbook failed. Its ${sheet.cellCount} cached cell value(s) are therefore from some earlier, unrecorded refresh, and every cell reading them displays normally.`,
        'Open the workbook with the source available so the link refreshes successfully, and confirm the values that changed.',
        { sheetId: sheet.sheetId ?? '', ...(sheet.sheetName === null ? {} : { sheetName: sheet.sheetName }) }
      ));
    }
    if (sheet.sheetName === null && sheetNames.length > 0) {
      problems.push(externalLinkFinding(
        'cached-sheet-id-unknown', linkPart,
        `A sheetData in ${linkPart} carries sheetId="${sheet.sheetId ?? 'absent'}", and this link lists ${sheetNames.length} sheet name(s), so that id does not select any of them. Its ${sheet.cellCount} cached value(s) cannot be attributed to a sheet of the source workbook.`,
        'Set sheetId to the zero-based position of the sheet in sheetNames, or add the missing sheetName entries.',
        { sheetId: sheet.sheetId ?? '' }
      ));
    }
  }

  return { relationshipId: id, target, targetIsExternal, sourceIsPresent, sheetNames, definedNames, cachedSheets };
};

/** Reads one external-link part and says which of the three things it turned out to be. */
const readLinkPart = (parts: PackageParts, linkPart: string, problems: Finding[]): {
  kind: ExternalLinkKind;
  book: ExternalBook | null;
} => {
  const doc = parseXml(parts[linkPart]);
  const root = doc?.documentElement;
  if (!root || !isX(root, 'externalLink')) {
    problems.push(externalLinkFinding(
      'link-part-unreadable', linkPart,
      `${linkPart} is referenced as an external link but ${root ? `its root element is ${root.nodeName} rather than externalLink` : 'does not parse as XML'}, so neither the source workbook it names nor the values it caches can be read.`,
      `Repair ${linkPart} so that it parses and its root element is externalLink.`
    ));
    return { kind: 'unreadable', book: null };
  }

  const book = firstChild(root, 'externalBook');
  if (book !== null) return { kind: 'externalBook', book: readBook(book, parts, linkPart, problems) };
  // Legal alternatives, not defects — see the header. Reported as what they are.
  if (firstChild(root, 'ddeLink') !== null) return { kind: 'ddeLink', book: null };
  if (firstChild(root, 'oleLink') !== null) return { kind: 'oleLink', book: null };

  problems.push(externalLinkFinding(
    'link-kind-unknown', linkPart,
    `${linkPart} is an externalLink part containing none of externalBook, ddeLink or oleLink; the schema requires exactly one of the three. The workbook lists this as an external reference and there is nothing in it to reference.`,
    `Add an externalBook (for a link to another workbook), a ddeLink or an oleLink to ${linkPart}, or remove the externalReference that points at it.`
  ));
  return { kind: 'unknown', book: null };
};

/** Resolves one `x:externalReference` down to a part, reporting each hop that breaks. */
const readReference = (
  element: Element,
  index: number,
  parts: PackageParts,
  rels: Map<string, Relationship> | null
): ExternalReference => {
  const problems: Finding[] = [];
  const id = relId(element);

  if (id === null || id === '') {
    problems.push(externalLinkFinding(
      'reference-no-relationship-id', WORKBOOK_PART,
      `External reference [${index}] declares no r:id, which the schema makes required. It still occupies position ${index} in the list, so every formula written as [${index}] resolves to a reference that names nothing, while the cells keep displaying their cached numbers.`,
      `Add an r:id to the externalReference at position ${index}, or delete it — but note that deleting it renumbers every later reference and silently repoints every [N] formula above it.`,
      { index: String(index) }
    ));
    return { index, relationshipId: id, partPath: null, kind: null, book: null, problems };
  }

  const rel = rels?.get(id);
  if (rel === undefined) {
    problems.push(externalLinkFinding(
      'reference-relationship-missing', WORKBOOK_PART,
      rels === null
        ? `External reference [${index}] names relationship "${id}", and ${relsPathFor(WORKBOOK_PART)} does not exist, so nothing can say which workbook it reads from. Formulas written as [${index}] still show their cached values.`
        : `External reference [${index}] names relationship "${id}", which ${relsPathFor(WORKBOOK_PART)} does not declare. The chain to the source workbook is broken at its first hop, and every formula written as [${index}] keeps displaying the number it last read.`,
      `Add a Relationship with Id="${id}" of type ${EXTERNAL_LINK_RELATIONSHIP_TYPE} to ${relsPathFor(WORKBOOK_PART)}.`,
      { index: String(index), relationshipId: id }
    ));
    return { index, relationshipId: id, partPath: null, kind: null, book: null, problems };
  }

  const partPath = resolveTarget(WORKBOOK_PART, rel.target);
  if (parts[partPath] === undefined) {
    problems.push(externalLinkFinding(
      'link-part-missing', WORKBOOK_PART,
      `External reference [${index}] resolves to "${partPath}", which is not in the package. The part that would name the source workbook and hold its cached values is gone, so nothing records where the numbers in every [${index}] formula came from — and those numbers still display.`,
      `Restore ${partPath}, or remove the externalReference and relationship "${id}" together.`,
      { index: String(index), target: partPath }
    ));
    return { index, relationshipId: id, partPath, kind: null, book: null, problems };
  }

  const { kind, book } = readLinkPart(parts, partPath, problems);
  return { index, relationshipId: id, partPath, kind, book, problems };
};

/**
 * The whole chain, from `xl/workbook.xml` outwards.
 *
 * The order of `references` is load-bearing and not incidental: it is the only thing that
 * gives `[N]` a meaning. Document order is preserved exactly.
 */
export function readExternalLinks(parts: PackageParts): ExternalLinkSet {
  const doc = parseXml(parts[WORKBOOK_PART]);
  const root = doc?.documentElement;
  if (!root || !isX(root, 'workbook')) {
    return { references: [], unreferencedParts: [], workbookRead: false, problems: [] };
  }

  const list = firstChild(root, 'externalReferences');
  const elements = list === null ? [] : childrenNamed(list, 'externalReference');
  const rels = readRelationships(parts, WORKBOOK_PART);

  const references = elements.map((element, position) => readReference(element, position + 1, parts, rels));

  const reached = new Set(references.map(reference => reference.partPath).filter(path => path !== null));
  const unreferencedParts = Object.keys(parts).filter(path => EXTERNAL_LINK_PART.test(path) && !reached.has(path));

  const problems = unreferencedParts.map(path => externalLinkFinding(
    'unreferenced-link-part', path,
    `${path} is an external-link part that no externalReference in ${WORKBOOK_PART} points at. Its cached values are carried in the file and reachable by nothing — a formula can only name a link by its position in the externalReferences list, and this part has no position. This is what is left behind when a link is removed from the workbook and its part is not.`,
    `Delete ${path} and its relationship, or add an externalReference in ${WORKBOOK_PART} naming it.`
  ));

  return { references, unreferencedParts, workbookRead: true, problems };
}

/**
 * `[N]` occurrences in a formula or defined-name expression.
 *
 * Two things this deliberately does NOT do, because doing them badly is worse than not
 * doing them:
 *
 *   • It strips double-quoted string literals first (`""` is an escaped quote inside one),
 *     so `IF(A1="[9]",…)` contributes nothing. Without that, any text mentioning a
 *     bracketed number would be read as a workbook reference.
 *   • It requires the `[` to follow something that cannot be part of an identifier, which
 *     is what separates an external index from a **structured table reference**:
 *     `Table1[2]` is a column named `2` in a table in *this* workbook and has nothing to
 *     do with external links. `excelFormulas.ts` uses a bare `/\[\d+\]/` for its own
 *     "this cell references another workbook" note; that is the more conservative choice
 *     for a note that only says "look at this", but it is too loose to resolve an index
 *     against a list and report a specific one as dangling.
 *
 * ⚠️ This is a scanner, not a formula parser. A column named `1` in a table whose name is
 * followed by nothing (`[1]` alone is not valid there) is the kind of edge it cannot see;
 * no tokeniser for the full formula grammar exists in this codebase and inventing a
 * partial one would fail less visibly than this does.
 */
export function externalIndexesIn(formula: string): number[] {
  const withoutStrings = formula.replace(/"(?:[^"]|"")*"/g, '');
  const found: number[] = [];
  for (const match of withoutStrings.matchAll(/(?:^|[^A-Za-z0-9_.$\]])\[(\d+)\]/g)) {
    found.push(Number(match[1]));
  }
  return found;
}

/** Where a `[N]` was written, so a finding can name it. */
interface IndexUse {
  part: string;
  /** `A1` for a formula, or `name "Total"` for a defined name. */
  where: string;
  index: number;
}

const collectIndexUses = (parts: PackageParts): IndexUse[] => {
  const uses: IndexUse[] = [];

  for (const [path, xml] of Object.entries(parts)) {
    if (!EXTERNAL_LINK_FORMULA_PART.test(path)) continue;
    const doc = parseXml(xml);
    if (!doc?.documentElement) continue;
    for (const cell of Array.from(doc.documentElement.querySelectorAll('*'))) {
      if (!isX(cell, 'c')) continue;
      const formula = firstChild(cell, 'f');
      if (formula === null) continue;
      for (const index of externalIndexesIn(formula.textContent ?? '')) {
        uses.push({ part: path, where: cell.getAttribute('r') ?? 'an unnamed cell', index });
      }
    }
  }

  // Workbook-level defined names carry their expression as ELEMENT TEXT, not in an
  // attribute (x:CT_DefinedName is a leaf-text type — see the header), and an external
  // index in one is exactly as invisible as one in a cell.
  const workbook = parseXml(parts[WORKBOOK_PART])?.documentElement;
  const definedNames = workbook === undefined || workbook === null ? null : firstChild(workbook, 'definedNames');
  for (const definedName of definedNames === null ? [] : childrenNamed(definedNames, 'definedName')) {
    for (const index of externalIndexesIn(definedName.textContent ?? '')) {
      uses.push({
        part: WORKBOOK_PART,
        where: `defined name "${definedName.getAttribute('name') ?? 'unnamed'}"`,
        index
      });
    }
  }

  return uses;
};

/**
 * `[N]` indexes that name no external reference.
 *
 * `[0]` means *this workbook*, not an out-of-range link — verified against LibreOffice's
 * compiler (see the header) — and reporting it would be a false positive on a construct
 * Excel writes itself. There is deliberately **no `index === 0` guard** below: an index
 * comes from `\d+` so it is never negative, `references.length` is never negative, and
 * `0 <= length` therefore always holds. An explicit guard would read as protection and
 * provably never fire, which is worse than none — mutation testing removed it and every
 * test still passed, which is how it was found.
 */
const formulaIndexFindings = (parts: PackageParts, set: ExternalLinkSet): Finding[] => {
  // Without the workbook there is no list, so there is no such thing as an index that
  // does not fit it. Reporting anything here would be reporting on evidence not present.
  if (!set.workbookRead) return [];

  const findings: Finding[] = [];
  const seen = new Set<string>();
  for (const use of collectIndexUses(parts)) {
    if (use.index <= set.references.length) continue;
    const key = `${use.part}|${use.where}|${use.index}`;
    if (seen.has(key)) continue;
    seen.add(key);

    findings.push(externalLinkFinding(
      'formula-index-unresolved', use.part,
      `${use.where} references external workbook [${use.index}], and ${WORKBOOK_PART} lists ${set.references.length} external reference(s), so there is no [${use.index}] to resolve. The formula names a source workbook that this file does not describe at all — no path, no sheet names, no cached values — while the cell keeps displaying the number it last read.`,
      set.references.length === 0
        ? `Add the externalReference this formula expects to ${WORKBOOK_PART}, or replace the reference with a value or a link inside this workbook.`
        : `Repoint the formula at one of the ${set.references.length} existing reference(s) [1]–[${set.references.length}], or restore the externalReference that used to occupy position ${use.index}.`,
      { index: String(use.index), where: use.where }
    ));
  }
  return findings;
};

/** Every finding this module produces for a package. */
export function externalLinkFindings(parts: PackageParts): Finding[] {
  const set = readExternalLinks(parts);
  return [
    ...set.references.flatMap(reference => reference.problems),
    ...set.problems,
    ...formulaIndexFindings(parts, set)
  ];
}

/**
 * Whether the source workbook is present, stated plainly.
 *
 * `null` means unknowable from the package alone, and for a healthy external link that is
 * the *expected* answer — the source is another file somewhere else, and answering would
 * mean leaving the package, which this code does not do. Deliberately not `false`: "we
 * cannot check" and "it is missing" are different answers and only one of them is a
 * defect. Same discipline as `oleDataIsPresent`.
 */
export function externalSourceIsPresent(reference: ExternalReference): boolean | null {
  return reference.book?.sourceIsPresent ?? null;
}

/**
 * Evidence lines for the AI panel.
 *
 * Leads with the cache framing, because every other statement depends on the reader
 * understanding that the linked numbers were copied at some unknown past moment rather
 * than read now.
 */
export function computeExternalLinkEvidenceForMarkup(
  parts: PackageParts
): { lines: string[]; unresolved: string[] } | null {
  const set = readExternalLinks(parts);
  // Dangling indexes are computed before the "is there anything to say" test, not after:
  // a formula written as [1] in a workbook that lists NO external references is the most
  // alarming arrangement this module can find — the file names a source it does not
  // describe at all — and testing only `references`/`unreferencedParts` returned null on
  // exactly that package, so the panel said nothing about it.
  const indexProblems = formulaIndexFindings(parts, set);
  if (set.references.length === 0 && set.unreferencedParts.length === 0 && indexProblems.length === 0) {
    return null;
  }

  const lines: string[] = [];
  const unresolved: string[] = [];

  lines.push(
    `${WORKBOOK_PART} lists ${set.references.length} external workbook reference(s). ` +
      'A formula writes one as a bracketed 1-based index into this list, so [1] is the first entry. ' +
      'Every value read through such a link is a copy cached when the two workbooks were last open together; ' +
      'the source is outside this package and nothing here reads it.'
  );

  for (const reference of set.references) {
    const book = reference.book;
    if (book === null) {
      lines.push(
        `[${reference.index}] resolves to ${reference.partPath ?? 'no part'}` +
          `${reference.kind === null ? '' : `, which is a ${reference.kind} rather than a workbook link`}.`
      );
    } else {
      const cached = book.cachedSheets;
      const cells = (cached ?? []).reduce((total, sheet) => total + sheet.cellCount, 0);
      lines.push(
        `[${reference.index}] is ${reference.partPath}, linking to ${book.target === null ? 'a source this package does not name' : `"${book.target}"`}` +
          `${book.sheetNames.length > 0 ? `, covering sheet(s) ${book.sheetNames.join(', ')}` : ''}` +
          `${cached === null ? ', with no cached values at all' : `, with ${cells} cached cell value(s) across ${cached.length} sheet(s)`}.`
      );
      if (book.targetIsExternal) {
        lines.push(
          `Whether "${book.target}" exists, is reachable, or still holds the values cached here cannot be determined from this package: it is a file outside the package, and nothing was opened or fetched to find out.`
        );
        unresolved.push(
          `The source workbook "${book.target}" is outside this package and was not read, so how old the cached values are, and whether the source still exists, is unknown.`
        );
      }
      if (cached !== null && cells > 0) {
        unresolved.push(
          `The ${cells} cached value(s) from "${book.target ?? 'the unnamed source'}" were read as stored; when they were copied is not recorded anywhere in the package, so their age cannot be established.`
        );
      }
    }
    lines.push(...renderFindings(reference.problems));
  }

  lines.push(...renderFindings(set.problems));
  lines.push(...renderFindings(indexProblems));

  return { lines, unresolved };
}
