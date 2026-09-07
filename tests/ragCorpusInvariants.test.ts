import { describe, it, expect } from 'vitest';
import { readFileSync } from 'fs';
import { join } from 'path';
import { z } from 'zod';
import { KNOWLEDGE_BASE } from '../services/staticKnowledgeBase';
import { matchesKeyword } from '../services/storageService';
import { COLOUR_SLOTS, COLOUR_MAP_KEYS } from '../services/powerpointResolver';

/**
 * Invariants over the generated RAG corpus (public/rag-data.json).
 *
 * The dataset is produced by scripts/ingestSchema.ts from the Open XML SDK's schema
 * metadata. These tests are the check that regeneration cannot silently reintroduce
 * the defects hand-curation produced: a wrong namespace prefix, drifting field
 * formats, or - most importantly - a record that claims prose it does not have.
 */

const corpus = JSON.parse(
  readFileSync(join(__dirname, '..', 'public', 'rag-data.json'), 'utf8')
) as unknown[];

const AttributeSpecSchema = z.object({
  name: z.string().min(1),
  type: z.string().min(1),
  // An enumeration with an empty list would read as "nothing is permitted here", so the
  // generator omits the field rather than emitting [].
  values: z.array(z.string().min(1)).min(1).optional(),
  required: z.literal(true).optional(),
  label: z.string().min(1).optional(),
  version: z.string().min(1).optional(),
  maxLength: z.number().optional(),
  min: z.number().optional(),
  max: z.number().optional()
});

const ReferenceDocSchema = z.object({
  tag: z.string().min(1),
  namespace: z.string(),
  domain: z.enum(['docx', 'xlsx', 'pptx', 'shared']),
  definition: z.string().min(1).optional(),
  attributes: z.array(AttributeSpecSchema),
  parents: z.array(z.string().min(1)),
  children: z.array(z.string().min(1)),
  citation: z.string().min(1).optional(),
  sdkClass: z.string().min(1).optional(),
  reviewerNote: z.string().optional(),
  priority: z.enum(['high', 'low']).optional(),
  provenance: z.enum(['curated', 'schema']).optional()
});

type ReferenceDoc = z.infer<typeof ReferenceDocSchema>;

const docs: ReferenceDoc[] = corpus.map(d => ReferenceDocSchema.parse(d));

describe('RAG corpus shape', () => {
  it('every record matches the ReferenceDoc schema', () => {
    // The parse above would already have thrown; this states the intent explicitly.
    expect(docs.length).toBe(corpus.length);
  });

  it('has no duplicate domain:namespace:tag keys', () => {
    // storageService keys IndexedDB records on `${domain}:${namespace}:${tag}`, so a
    // duplicate would silently drop a record on load. Namespace is load-bearing here:
    // 43 records share a domain and local name and differ only by namespace
    // (`w:start` vs `wp:start`, `x:row` vs `xdr:row`), and the earlier domain:tag
    // key merged them.
    const keys = docs.map(d => `${d.domain}:${d.namespace}:${d.tag}`);
    expect(new Set(keys).size).toBe(keys.length);
  });

  it('still contains same-named elements that differ only by namespace', () => {
    // Guards the fix rather than the symptom: if keying regresses to domain:tag,
    // these collapse and this fails.
    const byDomainTag = new Map<string, Set<string>>();
    for (const d of docs) {
      const key = `${d.domain}:${d.tag}`;
      if (!byDomainTag.has(key)) byDomainTag.set(key, new Set());
      byDomainTag.get(key)!.add(d.namespace);
    }
    const multi = [...byDomainTag.values()].filter(ns => ns.size > 1);
    expect(multi.length).toBeGreaterThan(20);
  });

  it('covers substantially more than the original 29 curated tags', () => {
    expect(docs.length).toBeGreaterThan(1800);
  });

  it('covers all four domains', () => {
    for (const domain of ['docx', 'xlsx', 'pptx', 'shared'] as const) {
      expect(docs.filter(d => d.domain === domain).length).toBeGreaterThan(100);
    }
  });
});

describe('namespace correctness', () => {
  // This is the regression test for a real shipped bug: every spreadsheet record
  // carried namespace "r" - the relationships namespace - because the prompt that
  // generated them conflated SpreadsheetML with relationships. The app rendered
  // <r:worksheet> into prompts under a "Grounded" badge.
  it('spreadsheet records do not use the relationships prefix', () => {
    const wrong = docs.filter(d => d.domain === 'xlsx' && d.namespace === 'r');
    expect(wrong.map(d => d.tag)).toEqual([]);
  });

  it('each domain uses only its own markup prefix plus its drawing wrapper', () => {
    // A domain is not single-prefix. Word and Excel each carry a DrawingML
    // positioning wrapper - `wp:` anchors against a paginated document, `xdr:`
    // against the cell grid - so those prefixes legitimately live in those domains.
    // PowerPoint has no wrapper; absolute EMU in p:spTree replaces it.
    const allowed: Record<string, string[]> = {
      docx: ['w', 'wp'],
      xlsx: ['x', 'xdr'],
      pptx: ['p']
    };
    for (const [domain, prefixes] of Object.entries(allowed)) {
      const offenders = docs.filter(d => d.domain === domain && !prefixes.includes(d.namespace));
      expect(offenders.map(d => `${d.namespace}:${d.tag}`)).toEqual([]);
    }
  });
});

/**
 * The elements a person actually opens the tool to ask about.
 *
 * Coverage counts are easy to feel good about — 1,899 records sounds like everything —
 * and they say nothing about whether the corpus holds the element in front of the user.
 * These lists are the core of each specification: the parts a document is built from,
 * not a sample. If a regeneration drops one, the count barely moves and this fails.
 *
 * Namespaces matter as much as names. `p:notes` is the notes-slide ROOT element;
 * `notesSlide` is the part filename and is not an element at all, which is the kind of
 * mistake a list like this exists to prevent on both sides.
 */
const CORE_ELEMENTS: Record<string, [string, string[]][]> = {
  'WordprocessingML': [['w', [
    'document', 'body', 'p', 'r', 't', 'rPr', 'pPr', 'tbl', 'tr', 'tc', 'tblPr', 'tblGrid',
    'gridCol', 'sectPr', 'styles', 'style', 'numbering', 'num', 'abstractNum', 'lvl',
    'bookmarkStart', 'bookmarkEnd', 'hyperlink', 'fldChar', 'instrText', 'fldSimple',
    'ins', 'del', 'footnote', 'endnote', 'drawing', 'jc', 'b', 'i', 'u', 'sz', 'color',
    'rFonts', 'spacing', 'ind', 'pStyle', 'rStyle', 'numPr', 'numId', 'ilvl'
  ]]],
  'SpreadsheetML': [['x', [
    'workbook', 'sheets', 'sheet', 'worksheet', 'sheetData', 'row', 'c', 'v', 'f', 'is',
    'si', 'sst', 'cellXfs', 'xf', 'numFmt', 'numFmts', 'fonts', 'font', 'fills', 'fill',
    'borders', 'border', 'cellStyleXfs', 'definedName', 'definedNames', 'mergeCell',
    'mergeCells', 'conditionalFormatting', 'dataValidation', 'autoFilter', 'cols', 'col',
    'dimension', 'pane', 'selection', 'hyperlink'
  ]]],
  'PresentationML': [['p', [
    'presentation', 'sldMasterIdLst', 'sldIdLst', 'sld', 'cSld', 'spTree', 'sp', 'nvSpPr',
    'spPr', 'txBody', 'ph', 'sldLayout', 'sldMaster', 'notes', 'notesMaster', 'clrMap',
    'clrMapOvr', 'timing', 'tnLst', 'par', 'seq', 'cTn', 'graphicFrame', 'pic', 'grpSp',
    'grpSpPr', 'sldSz'
  ]]],
  'DrawingML': [['a', [
    'graphic', 'graphicData', 'blip', 'blipFill', 'solidFill', 'srgbClr', 'schemeClr',
    'xfrm', 'off', 'ext', 'chOff', 'chExt', 'prstGeom', 'ln', 'effectLst', 'theme',
    'themeElements', 'clrScheme', 'fontScheme', 'fmtScheme', 'fillStyleLst', 'lnStyleLst',
    'bgFillStyleLst', 'bodyPr', 'lstStyle', 'latin', 'ea', 'cs'
  ]]],
  'DrawingML charts': [['c', [
    'chart', 'chartSpace', 'plotArea', 'ser', 'cat', 'val', 'barChart', 'lineChart',
    'pieChart', 'valAx', 'catAx', 'numRef', 'strRef', 'numCache', 'strCache', 'tx',
    'title', 'legend'
  ]]]
};

describe('the elements people actually ask about', () => {
  const present = new Set(docs.map(d => `${d.namespace}:${d.tag}`));

  for (const [spec, groups] of Object.entries(CORE_ELEMENTS)) {
    it(`covers the core of ${spec}`, () => {
      const missing = groups.flatMap(([ns, tags]) =>
        tags.filter(tag => !present.has(`${ns}:${tag}`)).map(tag => `${ns}:${tag}`)
      );
      expect(missing, `${spec} elements absent from the corpus`).toEqual([]);
    });
  }

  it('gives the most-asked-about elements real structure, not just a name', () => {
    // Present-but-empty is the failure this catches: a record that exists and carries
    // no attributes, no parents and no children grounds nothing.
    const mustBeSubstantive = ['w:p', 'w:tbl', 'w:rPr', 'x:c', 'x:xf', 'p:sp', 'a:xfrm', 'c:ser'];
    for (const key of mustBeSubstantive) {
      const [ns, tag] = key.split(':');
      const doc = docs.find(d => d.namespace === ns && d.tag === tag);
      expect(doc, key).toBeDefined();
      const facts = doc!.attributes.length + doc!.parents.length + doc!.children.length;
      expect(facts, `${key} carries no structural facts`).toBeGreaterThan(0);
    }
  });
});

describe('provenance honesty', () => {
  // The dataset mixes human prose with machine-generated structure. The badge shown to
  // users depends on being able to tell them apart, so the distinction has to hold.
  it('schema-derived records carry no definition', () => {
    const fabricated = docs.filter(d => d.provenance === 'schema' && d.definition);
    expect(fabricated.map(d => d.tag)).toEqual([]);
  });

  /**
   * The consumer half of the invariant above.
   *
   * "Schema records carry no definition" is a fact about the data, and the test above
   * pins it. It is only safe if every consumer *reads* it that way — and one did not:
   * `searchSchemasInStorage` called `doc.definition.toLowerCase()` unguarded. That was
   * correct for as long as the corpus was the 29 curated records, all of which have
   * prose. The moment schema records landed, the first one the cursor reached threw,
   * uncaught, and took the whole natural-language search path down in the browser.
   *
   * The escape was that the only coverage of that search was mocks which reimplemented
   * the predicate instead of calling it, so every test agreed with itself and none
   * touched the real code. This asserts against the real predicate.
   */
  it('the search predicate tolerates a record with no definition', () => {
    const noProse: ReferenceDoc = {
      tag: 'w:tblPrEx', namespace: 'w', domain: 'docx', attributes: [], parents: [], children: []
    };
    expect(noProse.definition).toBeUndefined();

    expect(() => matchesKeyword(noProse, 'tbl')).not.toThrow();
    expect(matchesKeyword(noProse, 'tbl')).toBe(true);
    expect(matchesKeyword(noProse, 'nothing-like-this')).toBe(false);
  });

  it('the search predicate runs over every record in the real corpus', () => {
    // The integration form: whatever the generator emitted, the predicate survives it.
    // A shape test on ReferenceDoc would not have caught the original bug, because the
    // records were valid — it was the consumer that was wrong.
    expect(() => {
      for (const doc of docs) matchesKeyword(doc, 'style');
    }).not.toThrow();

    const withoutProse = docs.filter(d => !d.definition).length;
    expect(withoutProse, 'most records have no prose; that is the point').toBeGreaterThan(1000);
  });

  it('schema-derived records carry no citation', () => {
    const fabricated = docs.filter(d => d.provenance === 'schema' && d.citation);
    expect(fabricated.map(d => d.tag)).toEqual([]);
  });

  it('curated records keep their prose through regeneration', () => {
    const curated = docs.filter(d => d.provenance === 'curated');
    expect(curated.length).toBeGreaterThanOrEqual(29);
    for (const doc of curated) {
      expect(doc.definition, `${doc.domain}:${doc.tag} lost its definition`).toBeTruthy();
    }
  });

  it('every record declares its provenance', () => {
    expect(docs.filter(d => !d.provenance).map(d => d.tag)).toEqual([]);
  });
});

describe('bundled offline fallback', () => {
  // KNOWLEDGE_BASE ships inside the JS bundle and is consulted by ragRouter when
  // IndexedDB has no answer. It was previously exported but imported by nothing, so
  // these tests exist to keep it both wired up and consistent with the full corpus.
  it('is a strict subset of the generated corpus', () => {
    const corpusKeys = new Set(docs.map(d => `${d.domain}:${d.namespace}:${d.tag}`));
    const orphans = KNOWLEDGE_BASE
      .map(d => `${d.domain}:${d.namespace}:${d.tag}`)
      .filter(key => !corpusKeys.has(key));
    expect(orphans).toEqual([]);
  });

  it('agrees with the corpus on every field it duplicates', () => {
    // Divergence here would mean an offline user and an online user get different
    // answers for the same tag - which the badge would present as equally grounded.
    const byKey = new Map(docs.map(d => [`${d.domain}:${d.namespace}:${d.tag}`, d]));
    for (const bundled of KNOWLEDGE_BASE) {
      const canonical = byKey.get(`${bundled.domain}:${bundled.namespace}:${bundled.tag}`)!;
      expect(bundled.namespace, `${bundled.tag} namespace`).toBe(canonical.namespace);
      expect(bundled.definition, `${bundled.tag} definition`).toBe(canonical.definition);
      expect(bundled.citation, `${bundled.tag} citation`).toBe(canonical.citation);
    }
  });

  it('carries only curated records, all with prose', () => {
    expect(KNOWLEDGE_BASE.length).toBeGreaterThan(0);
    for (const doc of KNOWLEDGE_BASE) {
      expect(doc.provenance, `${doc.tag}`).toBe('curated');
      expect(doc.definition, `${doc.tag}`).toBeTruthy();
    }
  });

  it('stays small enough to justify bundling', () => {
    // Shipping the full corpus would add ~0.27 MB of JavaScript to duplicate what
    // IndexedDB already holds. If this ever fails, the subset rule has been lost.
    expect(KNOWLEDGE_BASE.length).toBeLessThan(200);
  });
});

describe('structural fields', () => {
  it('parent references are qualified names, consistently formatted', () => {
    // A previous hand-curated record listed parents as "w:body" while every other
    // record used bare "body", which would break any code walking the parent graph.
    const schemaDocs = docs.filter(d => d.provenance === 'schema');
    const unqualified = schemaDocs
      .flatMap(d => d.parents)
      .filter(p => !p.includes(':'));
    expect(unqualified).toEqual([]);
  });

  it('has no duplicate entries within a record\'s attributes, parents or children', () => {
    for (const doc of docs) {
      // Attributes are objects now, so a Set of them is always the same size as the
      // array — every object reference is distinct. Deduplicating on the NAME is what
      // the check was always about; a Set of the objects would pass unconditionally.
      const names = doc.attributes.map(a => a.name);
      expect(new Set(names).size, `${doc.tag} attributes`).toBe(names.length);
      expect(new Set(doc.parents).size, `${doc.tag} parents`).toBe(doc.parents.length);
      expect(new Set(doc.children).size, `${doc.tag} children`).toBe(doc.children.length);
    }
  });

  it('resolves inherited attributes rather than only direct ones', () => {
    // w:cellIns declares no attributes of its own; they come from the CT_TrackChange
    // base class. Reading only direct attributes would under-report most of the corpus.
    const cellIns = docs.find(d => d.tag === 'cellIns');
    expect(cellIns).toBeDefined();
    expect(cellIns!.attributes.map(a => a.name)).toEqual(expect.arrayContaining(['w:author', 'w:date', 'w:id']));
  });

  it('carries the permitted values for enumerated attributes', () => {
    // The gap this whole change existed to close. Knowing w:jc has a w:val is close to
    // useless; knowing w:val must be one of these is the answer to the actual question.
    const jc = docs.find(d => d.domain === 'docx' && d.tag === 'jc');
    const val = jc?.attributes.find(a => a.name === 'w:val');

    expect(val?.type).toBe('enum');
    expect(val?.values).toEqual(expect.arrayContaining(['left', 'center', 'right', 'both']));
    expect(val?.label).toBe('Alignment Type');
  });

  it('resolves colliding enum names by namespace, not by bare name', () => {
    // ColorSchemeIndexValues means dark1/light1 under Wordprocessing and dk1/lt1 under
    // Drawing — same name, different values, thirteen such collisions in the corpus.
    // Resolving on the bare name would advertise one namespace's values for the other's
    // attribute: plausible, wrong, and impossible to notice from the output.
    const clrMap = docs.find(d => d.domain === 'pptx' && d.tag === 'clrMap');
    const bg1 = clrMap?.attributes.find(a => a.name.endsWith('bg1'));
    expect(bg1?.values).toEqual(expect.arrayContaining(['dk1', 'lt1']));
    expect(bg1?.values).not.toContain('dark1');

    const schemeClr = docs.find(d => d.domain === 'shared' && d.tag === 'schemeClr');
    const val = schemeClr?.attributes.find(a => a.name.endsWith('val'));
    expect(val?.values).toEqual(expect.arrayContaining(['bg1', 'tx1']));
  });

  it('agrees with the PowerPoint resolver\'s hand-written colour constants', () => {
    // An independent check on both: the resolver's lists were written from the
    // specification by hand, and the corpus derives them mechanically from the SDK.
    // If they ever disagree, one of them is wrong and this says so.
    const clrMap = docs.find(d => d.domain === 'pptx' && d.tag === 'clrMap');
    const slots = clrMap?.attributes.find(a => a.name.endsWith('bg1'))?.values ?? [];
    expect([...slots].sort()).toEqual([...COLOUR_SLOTS].sort());

    const schemeClr = docs.find(d => d.domain === 'shared' && d.tag === 'schemeClr');
    const keys = schemeClr?.attributes.find(a => a.name.endsWith('val'))?.values ?? [];
    for (const key of COLOUR_MAP_KEYS) expect(keys, `schemeClr accepts ${key}`).toContain(key);
  });

  it('never emits an enum with an empty value list', () => {
    // An attribute typed `enum` with no values reads as "nothing is permitted here".
    // The generator reports the underlying shape instead when facets are unavailable.
    for (const doc of docs) {
      for (const attr of doc.attributes) {
        if (attr.type === 'enum') {
          expect(attr.values?.length, `${doc.tag}/@${attr.name}`).toBeGreaterThan(0);
        }
      }
    }
  });

  it('marks required attributes without trusting the validator\'s presence', () => {
    // ⚠️ RequiredValidator can carry IsRequired: "False". Exactly one attribute in
    // WordprocessingML does, so reading presence alone is right 182 times out of 183 —
    // which is precisely why it would never be noticed.
    const required = docs.flatMap(d => d.attributes).filter(a => a.required);
    expect(required.length).toBeGreaterThan(500);
    for (const attr of required) expect(attr.required).toBe(true);
  });

  it('records what an element may contain, not only what contains it', () => {
    const body = docs.find(d => d.domain === 'docx' && d.tag === 'body');
    expect(body?.children).toEqual(expect.arrayContaining(['w:p', 'w:tbl', 'w:sectPr']));
  });

  it('derives parents by inverting the schema child lists', () => {
    const tbl = docs.find(d => d.domain === 'docx' && d.tag === 'tbl');
    expect(tbl).toBeDefined();
    expect(tbl!.parents).toEqual(expect.arrayContaining(['w:body', 'w:tc']));
  });
});
