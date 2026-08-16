import { describe, it, expect } from 'vitest';
import { readFileSync } from 'fs';
import { join } from 'path';
import { z } from 'zod';

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

const ReferenceDocSchema = z.object({
  tag: z.string().min(1),
  namespace: z.string(),
  domain: z.enum(['docx', 'xlsx', 'pptx', 'shared']),
  definition: z.string().min(1).optional(),
  attributes: z.array(z.string().min(1)),
  parents: z.array(z.string().min(1)),
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

  it('has no duplicate domain:tag keys', () => {
    // storageService keys IndexedDB records on `${domain}:${tag}`, so a duplicate
    // would silently drop a record on load.
    const keys = docs.map(d => `${d.domain}:${d.tag}`);
    expect(new Set(keys).size).toBe(keys.length);
  });

  it('covers substantially more than the original 29 curated tags', () => {
    expect(docs.length).toBeGreaterThan(1000);
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

  it('each document domain uses a single consistent prefix', () => {
    const expected: Record<string, string> = { docx: 'w', xlsx: 'x', pptx: 'p' };
    for (const [domain, prefix] of Object.entries(expected)) {
      const offenders = docs.filter(d => d.domain === domain && d.namespace !== prefix);
      expect(offenders.map(d => `${d.namespace}:${d.tag}`)).toEqual([]);
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

  it('has no duplicate entries within a record\'s attributes or parents', () => {
    for (const doc of docs) {
      expect(new Set(doc.attributes).size, `${doc.tag} attributes`).toBe(doc.attributes.length);
      expect(new Set(doc.parents).size, `${doc.tag} parents`).toBe(doc.parents.length);
    }
  });

  it('resolves inherited attributes rather than only direct ones', () => {
    // w:cellIns declares no attributes of its own; they come from the CT_TrackChange
    // base class. Reading only direct attributes would under-report most of the corpus.
    const cellIns = docs.find(d => d.tag === 'cellIns');
    expect(cellIns).toBeDefined();
    expect(cellIns!.attributes).toEqual(expect.arrayContaining(['w:author', 'w:date', 'w:id']));
  });

  it('derives parents by inverting the schema child lists', () => {
    const tbl = docs.find(d => d.domain === 'docx' && d.tag === 'tbl');
    expect(tbl).toBeDefined();
    expect(tbl!.parents).toEqual(expect.arrayContaining(['w:body', 'w:tc']));
  });
});
