/**
 * Distils `tests/spec-facts.json` from the normative ECMA-376 schemas.
 *
 *   pnpm run spec:facts
 *
 * WHY THIS EXISTS.
 *
 * `public/rag-data.json` is derived from the Open XML SDK's schema metadata — which is
 * Microsoft's description of what Office implements. That is the right source for the
 * corpus, because a user's question is about a file Office wrote. But it means the whole
 * dictionary rests on a single vendor's account of the format, and nothing checks it.
 *
 * This is the second opinion: ECMA-376's own XSDs, published by ECMA as electronic
 * inserts to Part 4 (Transitional Migration Features), which is the variant the corpus
 * targets. Two independent derivations of the same facts. Where they agree, the corpus
 * is corroborated by the standard; where they disagree, `tests/specConformance.test.ts`
 * records it deliberately rather than letting it pass unnoticed.
 *
 * The download is ~8 MB and happens only when this script is run, never in CI or at
 * runtime. Only the distilled facts are committed: element names, attribute names,
 * enumerated values and required flags — the vocabulary of the format itself, which is
 * present in every OOXML file ever written and in every library that reads one.
 *
 * ⚠️ TYPE NAMES ARE NOT GLOBALLY UNIQUE.
 *
 * Fifteen simpleType names are declared in more than one schema file, and `ST_Direction`
 * is declared in three of them with different value sets — `horz|vert` in PresentationML,
 * something else in WordprocessingML and DrawingML diagrams. Resolving a type by bare
 * name gives one file's values for another file's attribute. The first version of this
 * extractor did exactly that and reported 39 confident, entirely false disagreements
 * with the corpus, including that `p:blinds/@dir` accepts `norm|rev`. Every reference is
 * therefore resolved through the referring file's own prefix map.
 */

import { writeFileSync, mkdirSync } from 'node:fs';
import { join, dirname, basename } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { JSDOM } from 'jsdom';

const REPO_ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const OUTPUT_PATH = join(REPO_ROOT, 'tests', 'spec-facts.json');

/** Part 4 carries the Transitional schemas; the corpus targets Transitional. */
const ECMA_PART4 =
  'https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip';
const INNER_ZIP = 'OfficeOpenXML-XMLSchema-Transitional.zip';

const XS = 'http://www.w3.org/2001/XMLSchema';

/**
 * Schema file to the namespace prefix the corpus uses for it.
 *
 * `vml-officeDrawing.xsd` and `vml-spreadsheetDrawing.xsd` are deliberately absent: the
 * Open XML SDK publishes only VML main, so the corpus has no records for them and
 * listing them here would report every one of their elements as a false gap.
 */
const FILE_PREFIX: Record<string, string> = {
  'wml.xsd': 'w',
  'sml.xsd': 'x',
  'pml.xsd': 'p',
  'dml-main.xsd': 'a',
  'dml-chart.xsd': 'c',
  'dml-diagram.xsd': 'dgm',
  'dml-chartDrawing.xsd': 'cdr',
  'dml-picture.xsd': 'pic',
  'dml-lockedCanvas.xsd': 'lc',
  'dml-wordprocessingDrawing.xsd': 'wp',
  'dml-spreadsheetDrawing.xsd': 'xdr',

  // Ingested since the first version of this extractor. VML is here because an OLE
  // preview is a `v:shape`, which `services/oleObjects.ts` analyses.
  'vml-main.xsd': 'v',
  'shared-math.xsd': 'm',
  'shared-bibliography.xsd': 'b',
  'shared-documentPropertiesExtended.xsd': 'ap',
  'shared-documentPropertiesCustom.xsd': 'op',
  'shared-documentPropertiesVariantTypes.xsd': 'vt'
};

interface ComplexType {
  attributes: Map<string, { type: string; required: boolean }>;
  children: string[];
  base: string | null;
}

interface SchemaFile {
  document: Document;
  targetNamespace: string | null;
  /** `xmlns:a="…"` declarations, needed to resolve a prefixed type reference. */
  prefixes: Map<string, string>;
  simpleTypes: Map<string, string[]>;
  complexTypes: Map<string, ComplexType>;
}

/** Direct element children of `node` with the given XSD local name. */
const childrenNamed = (node: Element, localName: string): Element[] =>
  Array.from(node.children).filter(el => el.namespaceURI === XS && el.localName === localName);

const firstNamed = (node: Element, localName: string): Element | null =>
  childrenNamed(node, localName)[0] ?? null;

const parseSchema = (xml: string): SchemaFile => {
  const { DOMParser } = new JSDOM().window;
  const document = new DOMParser().parseFromString(xml, 'application/xml');
  const root = document.documentElement;
  const prefixes = new Map<string, string>();
  // jsdom's NamedNodeMap does not line up with the ambient DOM types, so the shape is
  // named explicitly rather than inferred as unknown.
  const declarations = Array.from(root.attributes) as { name: string; value: string }[];
  for (const attr of declarations) {
    if (attr.name === 'xmlns') prefixes.set('', attr.value);
    else if (attr.name.startsWith('xmlns:')) prefixes.set(attr.name.slice(6), attr.value);
  }

  return {
    document,
    targetNamespace: root.getAttribute('targetNamespace'),
    prefixes,
    simpleTypes: new Map(),
    complexTypes: new Map()
  };
};

/** Fills in the simple and complex type tables for one file. */
const indexTypes = (file: SchemaFile): void => {
  const root = file.document.documentElement;

  for (const st of childrenNamed(root, 'simpleType')) {
    const name = st.getAttribute('name');
    if (!name) continue;
    const values = Array.from(st.getElementsByTagNameNS(XS, 'enumeration'))
      .map(e => e.getAttribute('value'))
      .filter((v): v is string => v !== null);
    if (values.length > 0) file.simpleTypes.set(name, values);
  }

  for (const ct of childrenNamed(root, 'complexType')) {
    const name = ct.getAttribute('name');
    if (!name) continue;

    // Attributes and the content model sit under the extension when there is one, and
    // directly on the complexType when there is not.
    const extension = firstNamed(firstNamed(ct, 'complexContent') ?? ct, 'extension');
    const holder = extension ?? ct;

    const attributes = new Map<string, { type: string; required: boolean }>();
    for (const attr of childrenNamed(holder, 'attribute')) {
      const attrName = attr.getAttribute('name') ?? attr.getAttribute('ref')?.split(':').pop();
      if (!attrName) continue;
      attributes.set(attrName, {
        type: attr.getAttribute('type') ?? '',
        required: attr.getAttribute('use') === 'required'
      });
    }

    // Walk the content model, but never descend into a nested inline complexType — its
    // elements belong to it, not to this one.
    const children: string[] = [];
    const walk = (node: Element): void => {
      for (const child of Array.from(node.children)) {
        if (child.namespaceURI !== XS) continue;
        if (child.localName === 'element') {
          const en = child.getAttribute('name') ?? child.getAttribute('ref')?.split(':').pop();
          if (en) children.push(en);
        } else if (['sequence', 'choice', 'all', 'group'].includes(child.localName)) {
          walk(child);
        }
      }
    };
    walk(holder);

    file.complexTypes.set(name, {
      attributes,
      children,
      base: extension?.getAttribute('base') ?? null
    });
  }
};

interface SchemaSet {
  files: Map<string, SchemaFile>;
  /** Target namespace to the file that declares it. */
  byNamespace: Map<string, string>;
}

/** Resolves `a:ST_Foo` from within `home` to the file and local name that define it. */
const resolveRef = (
  ref: string | null,
  home: string,
  set: SchemaSet
): { file: string; name: string } | null => {
  if (!ref) return null;
  if (!ref.includes(':')) return { file: home, name: ref };

  const [prefix, local] = ref.split(':', 2);
  const uri = set.files.get(home)?.prefixes.get(prefix);
  if (!uri || uri === XS) return null; // xsd:string and friends carry no enumeration
  return { file: set.byNamespace.get(uri) ?? home, name: local };
};

/** Every attribute of a complexType, including those reached through its base chain. */
const attributesOf = (
  typeName: string,
  home: string,
  set: SchemaSet,
  depth = 0
): Map<string, { type: string; required: boolean; home: string }> => {
  const out = new Map<string, { type: string; required: boolean; home: string }>();
  if (depth > 12) return out;

  const complexType = set.files.get(home)?.complexTypes.get(typeName);
  if (!complexType) return out;

  for (const [name, info] of complexType.attributes) out.set(name, { ...info, home });

  const base = resolveRef(complexType.base, home, set);
  if (base) {
    for (const [name, info] of attributesOf(base.name, base.file, set, depth + 1)) {
      // The derived type wins, which is what extension means.
      if (!out.has(name)) out.set(name, info);
    }
  }
  return out;
};

export interface SpecAttribute {
  /**
   * Every distinct value set the specification permits for this element name.
   *
   * Usually one. But an element name is not a type: `w:jc` is declared as `CT_Jc` in a
   * paragraph and `CT_JcTable` in a table, and `w:type` is declared with five different
   * types across WordprocessingML. Both this extractor and the corpus generator keep one
   * declaration per name, and they do not always keep the same one — which produced a
   * dozen confident "divergences" that were nothing but the two sides having picked
   * different contexts. Recording all of them lets the check ask the question that is
   * actually meaningful: is the corpus's value set one the specification permits here?
   */
  valueSets?: string[][];
  required?: true;
}
export type SpecFacts = Record<string, Record<string, SpecAttribute>>;

/** Builds the facts table from a set of parsed schema files. */
export const buildSpecFacts = (set: SchemaSet): SpecFacts => {
  const facts: SpecFacts = {};

  for (const [fileName, file] of set.files) {
    const prefix = FILE_PREFIX[fileName];
    if (!prefix) continue;

    // Element declarations are overwhelmingly LOCAL — `<xsd:element name="jc"
    // type="CT_Jc"/>` inside a sequence — not global. Only 82 of roughly 1,900 sit at
    // the root of a schema, so scanning the root alone sees almost nothing.
    for (const el of Array.from(file.document.getElementsByTagNameNS(XS, 'element'))) {
      const name = el.getAttribute('name');
      const typeRef = el.getAttribute('type');
      if (!name || !typeRef) continue;

      const type = resolveRef(typeRef, fileName, set);
      if (!type) continue;

      const key = `${prefix}:${name}`;
      const entry = (facts[key] ??= {});

      for (const [attrName, info] of attributesOf(type.name, type.file, set)) {
        const simple = resolveRef(info.type, info.home, set);
        const values = simple
          ? set.files.get(simple.file)?.simpleTypes.get(simple.name)
          : undefined;

        const existing = entry[attrName] ?? {};
        const sets = existing.valueSets ?? [];
        if (values) {
          const candidate = [...values].sort();
          const seen = sets.some(s => s.join('\u0000') === candidate.join('\u0000'));
          if (!seen) sets.push(candidate);
        }
        entry[attrName] = {
          ...(sets.length > 0 ? { valueSets: sets } : {}),
          // An element declared under several complex types is required if any of them
          // requires it — the corpus merges the same way.
          ...(info.required || existing.required ? { required: true as const } : {})
        };
      }
    }
  }

  // Elements with no attributes at all carry no facts worth checking.
  for (const key of Object.keys(facts)) {
    if (Object.keys(facts[key]).length === 0) delete facts[key];
  }
  return facts;
};

const main = async (): Promise<void> => {
  console.log(`[spec] fetching ${ECMA_PART4}`);
  const response = await fetch(ECMA_PART4);
  if (!response.ok) throw new Error(`ECMA download failed: ${response.status}`);

  const outer = await JSZip.loadAsync(await response.arrayBuffer());
  const innerFile = outer.file(INNER_ZIP);
  if (!innerFile) throw new Error(`${INNER_ZIP} not found inside the Part 4 archive`);

  const inner = await JSZip.loadAsync(await innerFile.async('arraybuffer'));

  const set: SchemaSet = { files: new Map(), byNamespace: new Map() };
  for (const entry of Object.values(inner.files)) {
    if (entry.dir || !entry.name.endsWith('.xsd')) continue;
    const name = basename(entry.name);
    const file = parseSchema(await entry.async('string'));
    set.files.set(name, file);
    if (file.targetNamespace && !set.byNamespace.has(file.targetNamespace)) {
      set.byNamespace.set(file.targetNamespace, name);
    }
  }
  for (const file of set.files.values()) indexTypes(file);
  console.log(`[spec] parsed ${set.files.size} schema files`);

  const facts = buildSpecFacts(set);
  const elements = Object.keys(facts).length;
  const attributes = Object.values(facts).reduce((n, a) => n + Object.keys(a).length, 0);
  const enumerated = Object.values(facts)
    .reduce((n, a) => n + Object.values(a).filter(x => x.valueSets).length, 0);
  const polymorphic = Object.values(facts)
    .reduce((n, a) => n + Object.values(a).filter(x => (x.valueSets?.length ?? 0) > 1).length, 0);

  mkdirSync(dirname(OUTPUT_PATH), { recursive: true });
  writeFileSync(
    OUTPUT_PATH,
    `{\n${Object.keys(facts).sort()
      .map(k => `${JSON.stringify(k)}:${JSON.stringify(facts[k])}`)
      .join(',\n')}\n}\n`,
    'utf8'
  );

  console.log(`[spec] ${elements} elements, ${attributes} attributes, ${enumerated} enumerated`);
  console.log(`[spec] ${polymorphic} attribute(s) permit more than one value set by context`);
  console.log(`[spec] wrote tests/spec-facts.json`);
};

// Only run when invoked directly, so the helpers above stay importable by tests.
if (process.argv[1] && import.meta.url.endsWith(basename(process.argv[1]))) {
  main().catch(error => {
    console.error(error);
    process.exit(1);
  });
}
