/**
 * Generates public/rag-data.json from the Open XML SDK's published schema metadata.
 *
 * Run: npx tsx scripts/ingestSchema.ts
 *
 * Why this source: dotnet/Open-XML-SDK publishes machine-readable schema data under
 * data/schemas/*.json (MIT licensed). It carries everything the runtime needs for
 * structural grounding - valid attributes, valid parents, the SDK class name, the real
 * namespace prefix - and encodes constraints the raw ECMA XSDs do not, such as
 * attribute max-lengths and which Office version gated an attribute in.
 *
 * Crucially there is no language model anywhere in this pipeline. Every field it emits
 * is mechanically derived, so it cannot hallucinate a tag, an attribute, or a parent.
 * The one thing it cannot produce is a human-readable definition: the SDK's `Summary`
 * field is boilerplate ("Defines the Table Class.") and occasionally wrong, so records
 * generated here carry no definition and are marked `provenance: "schema"`. The
 * hand-written records in the existing dataset carry real prose and are preserved.
 */

import { writeFileSync, readFileSync, existsSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const REPO_ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const OUTPUT_PATH = join(REPO_ROOT, 'public', 'rag-data.json');
const STATIC_KB_PATH = join(REPO_ROOT, 'services', 'staticKnowledgeBase.ts');

/** Pin via env for reproducible builds; defaults to the tip of main. */
const SDK_REF = process.env.OOXML_SDK_REF || 'main';
const SDK_BASE = `https://raw.githubusercontent.com/dotnet/Open-XML-SDK/${SDK_REF}/data/schemas`;

/**
 * DrawingML maps to `shared` because it is used by all three document types - a
 * `a:solidFill` inside a chart is the same element whether the chart is in a
 * spreadsheet or a slide.
 */
const SOURCES: { file: string; domain: 'docx' | 'xlsx' | 'pptx' | 'shared' }[] = [
  { file: 'schemas_openxmlformats_org_wordprocessingml_2006_main.json', domain: 'docx' },
  { file: 'schemas_openxmlformats_org_spreadsheetml_2006_main.json', domain: 'xlsx' },
  { file: 'schemas_openxmlformats_org_presentationml_2006_main.json', domain: 'pptx' },

  // DrawingML. The largest documented-deviation surface of the three formats
  // (~444 MS-OI29500 entries against PresentationML's 100) and reusable across all
  // of them, which is why it is worth ingesting in full rather than just `a:`.
  { file: 'schemas_openxmlformats_org_drawingml_2006_main.json', domain: 'shared' },
  { file: 'schemas_openxmlformats_org_drawingml_2006_chart.json', domain: 'shared' },
  { file: 'schemas_openxmlformats_org_drawingml_2006_diagram.json', domain: 'shared' },
  { file: 'schemas_openxmlformats_org_drawingml_2006_chartDrawing.json', domain: 'shared' },
  { file: 'schemas_openxmlformats_org_drawingml_2006_picture.json', domain: 'shared' },
  { file: 'schemas_openxmlformats_org_drawingml_2006_lockedCanvas.json', domain: 'shared' },

  // The two format-specific positioning wrappers. DrawingML payloads are identical
  // across formats; only the way they are anchored differs - `wp:` positions against
  // a paginated document, `xdr:` against the cell grid. PowerPoint has no wrapper at
  // all, which is what absolute EMU in `p:spTree` replaces.
  { file: 'schemas_openxmlformats_org_drawingml_2006_wordprocessingDrawing.json', domain: 'docx' },
  { file: 'schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json', domain: 'xlsx' }
];

interface SdkAttribute {
  QName?: string;
  Validators?: { Name?: string; Arguments?: { Name?: string; Value?: string }[] }[];
  Version?: string;
}

interface SdkType {
  /** `"<complexType>/<elementQName>"`; an empty tail means a type-only entry. */
  Name: string;
  ClassName?: string;
  BaseClass?: string;
  Attributes?: SdkAttribute[];
  Children?: { Name: string }[];
}

interface ReferenceDoc {
  tag: string;
  namespace: string;
  domain: 'docx' | 'xlsx' | 'pptx' | 'shared';
  definition?: string;
  attributes: string[];
  parents: string[];
  citation?: string;
  sdkClass?: string;
  reviewerNote?: string;
  priority?: 'high' | 'low';
  /** `curated` records were written by a human; `schema` records are generated here. */
  provenance?: 'curated' | 'schema';
}

/** `"w:CT_Tbl/w:tbl"` -> `"w:tbl"`. Empty string for type-only entries. */
const elementQName = (name: string): string => name.slice(name.indexOf('/') + 1);

/** `"w:tbl"` -> `{ prefix: "w", tag: "tbl" }`. Unprefixed names get an empty prefix. */
const splitQName = (qname: string): { prefix: string; tag: string } => {
  const colon = qname.indexOf(':');
  return colon === -1
    ? { prefix: '', tag: qname }
    : { prefix: qname.slice(0, colon), tag: qname.slice(colon + 1) };
};

/**
 * Collects an element's attributes, walking the BaseClass chain.
 *
 * Only 174 of 726 WordprocessingML entries declare attributes directly; the other 408
 * inherit them from an abstract base (every tracked-change element gets `w:author`,
 * `w:date` and `w:id` from `CT_TrackChange` this way). Reading only the direct
 * attributes would under-report most of the corpus.
 */
const collectAttributes = (
  entry: SdkType,
  byClassName: Map<string, SdkType>,
  seen = new Set<string>()
): string[] => {
  const names: string[] = [];
  for (const attr of entry.Attributes ?? []) {
    if (attr.QName) names.push(attr.QName);
  }
  const base = entry.BaseClass;
  if (base && !seen.has(base)) {
    seen.add(base);
    const baseEntry = byClassName.get(base);
    if (baseEntry) names.push(...collectAttributes(baseEntry, byClassName, seen));
  }
  return names;
};

const fetchNamespace = async (file: string): Promise<{ Types: SdkType[] }> => {
  const url = `${SDK_BASE}/${file}`;
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Failed to fetch ${url}: ${response.status} ${response.statusText}`);
  }
  return response.json() as Promise<{ Types: SdkType[] }>;
};

const main = async () => {
  console.log(`[ingest] Open XML SDK ref: ${SDK_REF}`);

  // Preserve hand-written prose. Generated records cannot supply a definition, so the
  // curated dataset stays authoritative for every tag it covers.
  const curatedByKey = new Map<string, ReferenceDoc>();
  if (existsSync(OUTPUT_PATH)) {
    const existing = JSON.parse(readFileSync(OUTPUT_PATH, 'utf8')) as ReferenceDoc[];
    for (const doc of existing) {
      if (doc.definition) curatedByKey.set(`${doc.domain}:${doc.namespace}:${doc.tag}`, doc);
    }
    console.log(`[ingest] Preserving ${curatedByKey.size} curated records with prose.`);
  }

  const generated = new Map<string, ReferenceDoc>();

  // Load every namespace before emitting anything. Both derivations below have to be
  // global rather than per-file, because the namespaces genuinely cross-reference:
  // `a:graphic` is a child of `wp:inline` AND `xdr:graphicFrame`, and chart types
  // inherit from DrawingML base classes. Computing either per file silently
  // under-reports - the record would simply be missing parents it really has.
  const loaded: { domain: 'docx' | 'xlsx' | 'pptx' | 'shared'; file: string; Types: SdkType[] }[] = [];
  for (const { file, domain } of SOURCES) {
    const { Types } = await fetchNamespace(file);
    loaded.push({ domain, file, Types });
  }

  const byClassName = new Map<string, SdkType>();
  for (const { Types } of loaded) {
    for (const entry of Types) {
      if (entry.ClassName) byClassName.set(entry.ClassName, entry);
    }
  }

  // Parents are not stated anywhere in the SDK data; they are the inverse of the
  // Children lists.
  const parentsByQName = new Map<string, Set<string>>();
  for (const { Types } of loaded) {
    for (const entry of Types) {
      const parent = elementQName(entry.Name);
      if (!parent) continue;
      for (const child of entry.Children ?? []) {
        const childQName = elementQName(child.Name);
        if (!childQName) continue;
        if (!parentsByQName.has(childQName)) parentsByQName.set(childQName, new Set());
        parentsByQName.get(childQName)!.add(parent);
      }
    }
  }

  for (const { domain, file, Types } of loaded) {
    let emitted = 0;
    for (const entry of Types) {
      const qname = elementQName(entry.Name);
      if (!qname) continue; // type-only entry, not an element

      const { prefix, tag } = splitQName(qname);
      // Namespace is part of the key. Within `shared` alone, `ext` exists under
      // both `a:` and `cdr:`, and they are different elements.
      const key = `${domain}:${prefix}:${tag}`;

      // A tag can appear under several complex types (w:rPr sits under seven). Merge
      // rather than letting the last one win, so the record reflects every context.
      const existing = generated.get(key);
      const attributes = new Set(existing?.attributes ?? []);
      for (const a of collectAttributes(entry, byClassName)) attributes.add(a);
      const parents = new Set(existing?.parents ?? []);
      for (const p of parentsByQName.get(qname) ?? []) parents.add(p);

      generated.set(key, {
        tag,
        namespace: prefix,
        domain,
        attributes: [...attributes].sort(),
        parents: [...parents].sort(),
        sdkClass: entry.ClassName ?? existing?.sdkClass,
        provenance: 'schema'
      });
      emitted += 1;
    }
    const nsLabel = file.replace(/^schemas_openxmlformats_org_/, '').replace(/\.json$/, '');
    console.log(`[ingest] ${domain.padEnd(6)} ${String(emitted).padStart(5)} elements  ${nsLabel}`);
  }

  // Merge rule: prose from humans, structure from the schema.
  //
  // Curated records are authoritative only for the fields no machine can produce -
  // definition, citation, reviewer notes. Everything structural comes from the schema
  // whenever a schema record exists, because hand-curation demonstrably drifts: the
  // curated spreadsheet records all carried namespace "r" (the relationships
  // namespace) where the schema says "x", and curated parent lists were inconsistently
  // prefixed. Letting curation win on those fields would preserve both defects.
  const merged = new Map(generated);
  for (const [key, doc] of curatedByKey) {
    const schemaRecord = generated.get(key);
    merged.set(key, {
      ...(schemaRecord ?? doc),
      definition: doc.definition,
      citation: doc.citation,
      reviewerNote: doc.reviewerNote,
      priority: doc.priority,
      // An OPC package element such as <Relationship> has no entry in the four
      // document-markup namespaces, so it keeps its curated structure by necessity.
      sdkClass: schemaRecord?.sdkClass ?? doc.sdkClass,
      provenance: 'curated'
    });
  }

  const output = [...merged.values()].sort((a, b) =>
    a.domain !== b.domain ? a.domain.localeCompare(b.domain)
      : a.namespace !== b.namespace ? a.namespace.localeCompare(b.namespace)
      : a.tag.localeCompare(b.tag)
  );

  writeFileSync(OUTPUT_PATH, JSON.stringify(output, null, 2), 'utf8');

  // The bundled fallback deliberately carries only the curated records. They are the
  // ones with prose, so they are worth the bundle bytes; shipping all 1,500 would add
  // a quarter-megabyte of JavaScript to duplicate what IndexedDB already holds.
  const bundled = output.filter(d => d.provenance === 'curated');
  writeFileSync(STATIC_KB_PATH, `// This file is auto-generated by scripts/ingestSchema.ts.
// To regenerate: pnpm run ingest:schema
// Do not edit this file manually.

export interface ReferenceDoc {
  tag: string;
  namespace: string;
  domain: 'docx' | 'xlsx' | 'pptx' | 'shared';
  /**
   * Human-readable prose. Optional because the bulk of the dataset is generated from
   * the Open XML SDK's schema metadata, which supplies structure but no usable
   * description - its own \`Summary\` field is boilerplate ("Defines the Table Class.")
   * and sometimes wrong, so generating from it would manufacture false authority.
   * Records without a definition are structurally grounded only; see ragRouter.
   */
  definition?: string;
  attributes: string[];
  parents: string[];
  citation?: string;
  sdkClass?: string;
  reviewerNote?: string;
  priority?: 'high' | 'low';
  /** \`curated\` records were written by a human; \`schema\` records are generated. */
  provenance?: 'curated' | 'schema';
}

/**
 * Offline fallback consulted by ragRouter when IndexedDB has no answer - because the
 * store has not been populated yet, the /rag-data.json fetch failed, or IndexedDB is
 * unavailable entirely. Only the curated subset is bundled; the full corpus lives in
 * public/rag-data.json.
 */
export const KNOWLEDGE_BASE: ReferenceDoc[] = ${JSON.stringify(bundled, null, 2)};
`, 'utf8');

  const curatedCount = output.filter(d => d.provenance === 'curated').length;
  const bytes = Buffer.byteLength(JSON.stringify(output));
  console.log(`\n[ingest] Wrote ${output.length} records to public/rag-data.json`);
  console.log(`[ingest]   ${curatedCount} curated (with prose), ${output.length - curatedCount} schema-derived (structure only)`);
  console.log(`[ingest]   ${(bytes / 1024 / 1024).toFixed(2)} MB minified`);
};

main().catch(error => {
  console.error('[ingest] Failed:', error);
  process.exit(1);
});
