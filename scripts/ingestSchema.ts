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

  // VML. Legacy, and unavoidable: an OLE object's preview image is a `v:shape` with a
  // `v:imagedata` inside a `w:pict`, and `services/oleObjects.ts` analyses exactly that.
  // Without this the corpus knew `w:pict` and nothing inside it, so clicking the preview
  // of a broken embedding — the thing that analyzer exists to catch — returned nothing.
  { file: 'schemas-microsoft-com_vml.json', domain: 'shared' },

  // Office Math (OMML). Every equation in every Word document is `m:oMath`, and none of
  // it was covered.
  { file: 'schemas_openxmlformats_org_officeDocument_2006_math.json', domain: 'docx' },

  // docProps/app.xml and docProps/custom.xml, which the tree view shows for every file
  // opened. Variant types are the value side of a custom property.
  { file: 'schemas_openxmlformats_org_officeDocument_2006_extended-properties.json', domain: 'shared' },
  { file: 'schemas_openxmlformats_org_officeDocument_2006_custom-properties.json', domain: 'shared' },
  { file: 'schemas_openxmlformats_org_officeDocument_2006_docPropsVTypes.json', domain: 'shared' },

  // Word's citation store, reached from `w:bibliography` fields.
  { file: 'schemas_openxmlformats_org_officeDocument_2006_bibliography.json', domain: 'docx' },

  // The two format-specific positioning wrappers. DrawingML payloads are identical
  // across formats; only the way they are anchored differs - `wp:` positions against
  // a paginated document, `xdr:` against the cell grid. PowerPoint has no wrapper at
  // all, which is what absolute EMU in `p:spTree` replaces.
  { file: 'schemas_openxmlformats_org_drawingml_2006_wordprocessingDrawing.json', domain: 'docx' },
  { file: 'schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json', domain: 'xlsx' }
];

interface SdkValidator {
  Name?: string;
  Arguments?: { Type?: string; Name?: string; Value?: string }[];
}

interface SdkAttribute {
  QName?: string;
  /** `.NET` type, e.g. `StringValue` or `EnumValue<...JustificationValues>`. */
  Type?: string;
  /** A short human label, e.g. "Alignment Type". Present on every attribute. */
  PropertyComments?: string;
  Validators?: SdkValidator[];
}

/** One enumeration and the values it permits. */
interface SdkEnum {
  Name?: string;
  Facets?: { Value?: string }[];
}

/**
 * What an attribute permits — the half of the schema this ingest used to discard.
 *
 * `QName` alone says an attribute exists. It does not say `w:jc/@w:val` must be one of
 * twelve alignment keywords, which is the fact anyone actually asking about `w:jc`
 * wants. All of it is mechanically derived, so nothing here can be fabricated.
 */
interface AttributeSpec {
  name: string;
  /** Simplified from the SDK's .NET type: `enum`, `string`, `integer`, `boolean`, … */
  type: string;
  /** Every permitted value, for an enumerated attribute. */
  values?: string[];
  /** True only when the schema says so — see `isRequired` for the trap. */
  required?: true;
  label?: string;
  /** Office version that gates this attribute, e.g. `Office2010`. */
  version?: string;
  maxLength?: number;
  min?: number;
  max?: number;
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
  attributes: AttributeSpec[];
  parents: string[];
  /** Elements this one may contain. The inverse of `parents`, and just as useful. */
  children: string[];
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
 * `EnumValue<DocumentFormat.OpenXml.Wordprocessing.JustificationValues>` ->
 * `{ namespace: "DocumentFormat.OpenXml.Wordprocessing", name: "JustificationValues" }`.
 *
 * The .NET namespace is load-bearing, not decoration. `ColorSchemeIndexValues` means
 * `dark1, light1, …` under `Wordprocessing` and `dk1, lt1, …` under `Drawing` — the same
 * bare name, genuinely different permitted values. Eighteen enum names collide this way
 * across the eleven source files, so resolving on the bare name alone would advertise
 * one namespace's values for the other's attribute: plausible, wrong, and invisible.
 */
const enumRefOf = (dotNetType: string): { namespace: string; name: string } | null => {
  const match = /^EnumValue<(?:(.*)\.)?([^.<>]+)>$/.exec(dotNetType);
  return match ? { namespace: match[1] ?? '', name: match[2] } : null;
};

/** The SDK's .NET wrapper types, reduced to words a reader recognises. */
const SIMPLE_TYPES: Record<string, string> = {
  StringValue: 'string',
  OnOffValue: 'boolean',
  BooleanValue: 'boolean',
  TrueFalseValue: 'boolean',
  TrueFalseBlankValue: 'boolean',
  Int32Value: 'integer',
  Int64Value: 'integer',
  UInt32Value: 'integer',
  UInt64Value: 'integer',
  ByteValue: 'integer',
  SByteValue: 'integer',
  Int16Value: 'integer',
  UInt16Value: 'integer',
  DecimalValue: 'decimal',
  DoubleValue: 'decimal',
  SingleValue: 'decimal',
  HexBinaryValue: 'hexBinary',
  Base64BinaryValue: 'base64Binary',
  DateTimeValue: 'dateTime'
};

const simplifyType = (dotNetType: string | undefined, isEnum: boolean): string => {
  if (isEnum) return 'enum';
  if (!dotNetType) return 'unknown';
  if (SIMPLE_TYPES[dotNetType]) return SIMPLE_TYPES[dotNetType];
  // ListValue<StringValue> and friends keep their shape but lose the namespace noise.
  const generic = /^(\w+)<(?:.*\.)?([^.<>]+)>$/.exec(dotNetType);
  if (generic) return `${generic[1] === 'ListValue' ? 'list' : generic[1].toLowerCase()}<${SIMPLE_TYPES[generic[2]] ?? generic[2]}>`;
  return dotNetType;
};

/**
 * Enumerations, keyed by the .NET namespace that declares them.
 *
 * The SDK's `Enums` entries carry only a bare name, but each source file corresponds to
 * one .NET namespace, which its own attributes name in full. So the namespace is derived
 * from the file rather than hardcoded — a mapping table would rot the first time the SDK
 * adds a schema.
 */
interface EnumTable {
  /** Values for `namespace.name`, or undefined when nothing can be said safely. */
  resolve(namespace: string, name: string): string[] | undefined;
  size: number;
  /** Names that mean different things in different namespaces. Reported, not hidden. */
  collisions: string[];
}

const buildEnumTable = (
  files: { Types: SdkType[]; Enums: SdkEnum[] }[]
): EnumTable => {
  const qualified = new Map<string, string[]>();
  const byBareName = new Map<string, string[][]>();

  for (const { Types, Enums } of files) {
    // The file's own .NET namespace, taken from whichever it names most often.
    const counts = new Map<string, number>();
    for (const type of Types) {
      for (const attr of type.Attributes ?? []) {
        const ref = attr.Type ? enumRefOf(attr.Type) : null;
        if (ref?.namespace) counts.set(ref.namespace, (counts.get(ref.namespace) ?? 0) + 1);
      }
    }
    const fileNamespace = [...counts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? '';

    for (const declared of Enums) {
      if (!declared.Name) continue;
      const values = (declared.Facets ?? []).map(f => f.Value).filter((v): v is string => Boolean(v));
      if (values.length === 0) continue;
      qualified.set(`${fileNamespace}.${declared.Name}`, values);
      byBareName.set(declared.Name, [...(byBareName.get(declared.Name) ?? []), values]);
    }
  }

  const collisions = [...byBareName.entries()]
    .filter(([, variants]) => new Set(variants.map(v => v.join('\u0000'))).size > 1)
    .map(([name]) => name);
  const ambiguous = new Set(collisions);

  return {
    size: qualified.size,
    collisions,
    resolve(namespace, name) {
      const exact = qualified.get(`${namespace}.${name}`);
      if (exact) return exact;
      // No exact match: fall back to the bare name only when every namespace that
      // declares it agrees. When they disagree, say nothing — an attribute with no
      // listed values is a gap, an attribute with the wrong ones is a falsehood.
      if (ambiguous.has(name)) return undefined;
      return byBareName.get(name)?.[0];
    }
  };
};

const argumentOf = (validator: SdkValidator, name: string): string | undefined =>
  (validator.Arguments ?? []).find(a => a.Name === name)?.Value;

/**
 * ⚠️ The presence of a `RequiredValidator` does NOT mean the attribute is required.
 *
 * One attribute in WordprocessingML carries `RequiredValidator` with an explicit
 * `IsRequired: "False"` argument. Reading presence alone reports it as mandatory, which
 * is exactly the kind of wrong-but-plausible fact this corpus exists to avoid — and it
 * would be invisible, because 182 of the 183 cases give the same answer either way.
 */
const isRequired = (validators: SdkValidator[]): boolean => {
  const required = validators.find(v => v.Name === 'RequiredValidator');
  if (!required) return false;
  return (argumentOf(required, 'IsRequired') ?? 'True').toLowerCase() !== 'false';
};

const numberOrUndefined = (raw: string | undefined): number | undefined => {
  if (raw === undefined) return undefined;
  const n = Number(raw);
  return Number.isFinite(n) ? n : undefined;
};

/** Builds the full spec for one attribute. */
const specFor = (attr: SdkAttribute, enums: EnumTable): AttributeSpec | null => {
  if (!attr.QName) return null;
  // The SDK spells an unqualified attribute ":t", with an empty prefix before the colon.
  // Most OOXML attributes are unqualified — 2,932 of 4,056 here — so keeping the colon
  // would print ":allowOverlap" to the reader for the majority of the corpus, and make
  // every lookup by attribute name miss.
  const name = attr.QName.startsWith(':') ? attr.QName.slice(1) : attr.QName;
  const validators = attr.Validators ?? [];

  const ref = attr.Type ? enumRefOf(attr.Type) : null;
  const values = ref ? enums.resolve(ref.namespace, ref.name) : undefined;

  const stringValidator = validators.find(v => v.Name === 'StringValidator');
  const numberValidator = validators.find(v => v.Name === 'NumberValidator');
  const versionValidator = validators.find(v => v.Name === 'OfficeVersionValidator');

  return {
    name,
    // An enum whose facets are missing is reported by its underlying shape rather than
    // as an enum with no values, which would read as "nothing is permitted here".
    type: simplifyType(attr.Type, Boolean(values?.length)),
    ...(values?.length ? { values } : {}),
    ...(isRequired(validators) ? { required: true as const } : {}),
    ...(attr.PropertyComments ? { label: attr.PropertyComments } : {}),
    ...(versionValidator ? { version: (versionValidator.Arguments ?? [])[0]?.Value } : {}),
    ...(stringValidator ? { maxLength: numberOrUndefined(argumentOf(stringValidator, 'MaxLength')) } : {}),
    ...(numberValidator ? {
      min: numberOrUndefined(argumentOf(numberValidator, 'MinInclusive')),
      max: numberOrUndefined(argumentOf(numberValidator, 'MaxInclusive'))
    } : {})
  };
};

/**
 * Collects an element's attributes, walking the BaseClass chain.
 *
 * Only 174 of 726 WordprocessingML entries declare attributes directly; the other 408
 * inherit them from an abstract base (every tracked-change element gets `w:author`,
 * `w:date` and `w:id` from `CT_TrackChange` this way). Reading only the direct
 * attributes would under-report most of the corpus.
 *
 * The derived class wins on a name collision, which is what overriding means.
 */
const collectAttributes = (
  entry: SdkType,
  byClassName: Map<string, SdkType>,
  enums: EnumTable,
  seen = new Set<string>()
): AttributeSpec[] => {
  const specs: AttributeSpec[] = [];
  for (const attr of entry.Attributes ?? []) {
    const spec = specFor(attr, enums);
    if (spec) specs.push(spec);
  }
  const base = entry.BaseClass;
  if (base && !seen.has(base)) {
    seen.add(base);
    const baseEntry = byClassName.get(base);
    if (baseEntry) {
      const own = new Set(specs.map(s => s.name));
      for (const inherited of collectAttributes(baseEntry, byClassName, enums, seen)) {
        if (!own.has(inherited.name)) specs.push(inherited);
      }
    }
  }
  return specs;
};

/**
 * Curated records predate attribute specs and list attributes as bare names.
 *
 * A handful of them — the OPC package elements such as `<Relationship>` — have no
 * counterpart in the four markup namespaces, so no schema record ever replaces them and
 * their bare names are all there will ever be. Lifting them to the spec shape keeps one
 * type across the corpus; `type: 'unknown'` is the honest value, because nothing in the
 * schema was consulted to produce it.
 */
const liftCuratedAttributes = (doc: ReferenceDoc): AttributeSpec[] =>
  (doc.attributes as unknown as (string | AttributeSpec)[]).map(a =>
    typeof a === 'string' ? { name: a, type: 'unknown' } : a
  );

const fetchNamespace = async (file: string): Promise<{ Types: SdkType[]; Enums?: SdkEnum[] }> => {
  const url = `${SDK_BASE}/${file}`;
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Failed to fetch ${url}: ${response.status} ${response.statusText}`);
  }
  return response.json() as Promise<{ Types: SdkType[]; Enums?: SdkEnum[] }>;
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
  const loaded: { domain: 'docx' | 'xlsx' | 'pptx' | 'shared'; file: string; Types: SdkType[]; Enums: SdkEnum[] }[] = [];
  for (const { file, domain } of SOURCES) {
    const { Types, Enums } = await fetchNamespace(file);
    loaded.push({ domain, file, Types, Enums: Enums ?? [] });
  }

  const enums = buildEnumTable(loaded);
  console.log(`[ingest] ${enums.size} enumerations loaded`);
  if (enums.collisions.length > 0) {
    console.log(
      `[ingest] ${enums.collisions.length} enum name(s) mean different things in different ` +
      `namespaces and are resolved by namespace: ${enums.collisions.slice(0, 5).join(', ')}…`
    );
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
  const childrenByQName = new Map<string, Set<string>>();
  for (const { Types } of loaded) {
    for (const entry of Types) {
      const parent = elementQName(entry.Name);
      if (!parent) continue;
      for (const child of entry.Children ?? []) {
        const childQName = elementQName(child.Name);
        if (!childQName) continue;
        if (!parentsByQName.has(childQName)) parentsByQName.set(childQName, new Set());
        parentsByQName.get(childQName)!.add(parent);
        if (!childrenByQName.has(parent)) childrenByQName.set(parent, new Set());
        childrenByQName.get(parent)!.add(childQName);
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
      // Merged by name across complex types. The first spec for a name wins, which
      // keeps the derived class's override rather than a base class's declaration.
      const attributes = new Map((existing?.attributes ?? []).map(a => [a.name, a]));
      for (const spec of collectAttributes(entry, byClassName, enums)) {
        if (!attributes.has(spec.name)) attributes.set(spec.name, spec);
      }
      const parents = new Set(existing?.parents ?? []);
      for (const p of parentsByQName.get(qname) ?? []) parents.add(p);
      const children = new Set(existing?.children ?? []);
      for (const c of childrenByQName.get(qname) ?? []) children.add(c);

      generated.set(key, {
        tag,
        namespace: prefix,
        domain,
        attributes: [...attributes.values()].sort((a, b) => a.name.localeCompare(b.name)),
        parents: [...parents].sort(),
        children: [...children].sort(),
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
      ...(schemaRecord ?? { ...doc, attributes: liftCuratedAttributes(doc), children: doc.children ?? [] }),
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

  // One record per line: valid JSON, parses with a plain JSON.parse, and still diffs
  // per record. Pretty-printing every field cost 477 KB — a third of the file — to
  // produce a diff nobody reads, because the file is regenerated wholesale.
  //
  // Nothing beyond this is worth doing, and that is measured rather than assumed. A
  // dictionary-encoded variant (deduplicated attribute specs and element names, records
  // as index arrays) came to 92 KB gzipped against this file's 102 KB, and a
  // protobuf-style binary with varint indices came to 90 KB — 1.5% better than the
  // dictionary for the cost of a decoder and a runtime dependency larger than the
  // saving. The payload is overwhelmingly distinct strings, whose gzipped floor is
  // 49 KB on its own, so no encoding can do much about the rest.
  writeFileSync(
    OUTPUT_PATH,
    `[\n${output.map(record => JSON.stringify(record)).join(',\n')}\n]\n`,
    'utf8'
  );

  // The bundled fallback deliberately carries only the curated records. They are the
  // ones with prose, so they are worth the bundle bytes; shipping all 1,500 would add
  // a quarter-megabyte of JavaScript to duplicate what IndexedDB already holds.
  const bundled = output.filter(d => d.provenance === 'curated');
  writeFileSync(STATIC_KB_PATH, `// This file is auto-generated by scripts/ingestSchema.ts.
// To regenerate: pnpm run ingest:schema
// Do not edit this file manually.

/**
 * What an attribute permits, derived mechanically from the Open XML SDK schema.
 *
 * \`values\` is the half that matters most: it is the difference between knowing
 * \`w:jc\` has a \`w:val\` and knowing that \`w:val\` must be one of twelve alignment
 * keywords. Absent when the attribute is not enumerated, or when two namespaces declare
 * the same enum name with different values and the right one could not be established.
 */
export interface AttributeSpec {
  name: string;
  /** \`enum\`, \`string\`, \`integer\`, \`boolean\`, \`hexBinary\`, \`dateTime\`, … */
  type: string;
  values?: string[];
  required?: true;
  label?: string;
  /** Office version that gates the attribute, e.g. \`Office2010\`. */
  version?: string;
  maxLength?: number;
  min?: number;
  max?: number;
}

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
  attributes: AttributeSpec[];
  parents: string[];
  /** Elements this one may contain — the inverse of \`parents\`. */
  children: string[];
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
