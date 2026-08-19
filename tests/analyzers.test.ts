import { describe, it, expect } from 'vitest';
import {
  ANALYZERS,
  analyzePackage,
  capabilityLedger,
  diffFindings,
  explainersFor,
  siblingsFor,
  explainPart,
  type Analyzer
} from '../services/analyzers';
import { analyzerOf } from '../services/findings';
import type { PackageParts } from '../services/packageIntegrity';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const O = 'xmlns:o="urn:schemas-microsoft-com:office:office"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const V = 'xmlns:v="urn:schemas-microsoft-com:vml"';

const CONTENT_TYPES = `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Default Extension="emf" ContentType="image/x-emf"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const rels = (body: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;

/**
 * A Word package with one intact embedded object.
 *
 * The VML preview is not decoration: an OLE object without one is a real finding
 * (`ole/no-preview`), so a fixture missing it is not actually healthy.
 */
const healthy = (): PackageParts => ({
  '[Content_Types].xml': CONTENT_TYPES,
  '_rels/.rels': rels(
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>`
  ),
  'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${O} ${V} ${R}><w:body><w:p><w:r><w:object>
      <v:shape><v:imagedata r:id="rId5"/></v:shape>
      <o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" r:id="rId4"/>
    </w:object></w:r></w:p></w:body></w:document>`,
  'word/_rels/document.xml.rels': rels(
    `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/>` +
    `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.emf"/>`
  ),
  'word/embeddings/oleObject1.bin': 'BINARY',
  'word/media/image1.emf': 'IMAGE'
});

describe('the registry', () => {
  it('gives every analyzer an id matching the codes it emits', () => {
    // The namespace IS the join between a finding and the analyzer that made it. If
    // they drift, grouping and suppression silently stop working.
    const parts = healthy();
    delete parts['word/embeddings/oleObject1.bin'];

    for (const found of analyzePackage(parts).findings) {
      expect(ANALYZERS.map(a => a.id)).toContain(analyzerOf(found.code));
    }
  });

  it('makes every analyzer declare something it cannot determine', () => {
    // An analyzer claiming no limits is almost certainly lying, and the capability
    // ledger would then overstate what a clean run means.
    for (const analyzer of ANALYZERS) {
      expect(analyzer.cannotDetermine.length, analyzer.id).toBeGreaterThan(0);
      expect(analyzer.determines.length, analyzer.id).toBeGreaterThan(0);
    }
  });
});

describe('analyzePackage', () => {
  it('runs every analyzer that applies, not just package integrity', () => {
    // The gap this closes: the validate surface used to run one analyzer, so "is this
    // file correct?" checked content types and relationships and nothing else.
    const parts = healthy();
    delete parts['word/embeddings/oleObject1.bin'];
    const run = analyzePackage(parts);

    expect(run.ran).toContain('package');
    expect(run.ran).toContain('ole');
    expect(run.findings.some(f => f.code === 'ole/data-part-missing')).toBe(true);
  });

  it('skips analyzers with nothing to look at, and says which', () => {
    const run = analyzePackage(healthy());

    expect(run.skipped).toContain('pivot');
    expect(run.ran).not.toContain('pivot');
  });

  it('keeps the other analyzers findings when one throws', () => {
    // A malformed corner of a package is exactly when the rest of the report matters.
    const parts = { ...healthy(), 'word/document.xml': '<w:document><unclosed>' };

    expect(() => analyzePackage(parts)).not.toThrow();
    expect(analyzePackage(parts).ran).toContain('package');
  });

  it('finds nothing wrong with a healthy package', () => {
    expect(analyzePackage(healthy()).findings).toEqual([]);
  });
});

describe('capabilityLedger', () => {
  it('reports what ran, what was skipped, and what the checks cannot see', () => {
    const ledger = capabilityLedger(analyzePackage(healthy()));

    expect(ledger.ran.map(a => a.id)).toContain('ole');
    expect(ledger.skipped.map(a => a.id)).toContain('pivot');
    expect(ledger.limits.length).toBeGreaterThan(0);
  });

  it('lists limits only from analyzers that actually ran', () => {
    // Otherwise a clean report on a .docx would disclaim things about pivot caches,
    // which is noise, and would imply the pivot checks had been considered.
    const ledger = capabilityLedger(analyzePackage(healthy()));

    expect(ledger.limits.some(l => l.includes('cached records'))).toBe(false);
  });

  it('is computed from the registry, so it cannot disagree with what ran', () => {
    const run = analyzePackage(healthy());
    const ledger = capabilityLedger(run);

    expect(ledger.ran.map(a => a.id)).toEqual(run.ran);
    expect(ledger.skipped.map(a => a.id)).toEqual(run.skipped);
  });
});

describe('diffFindings — what a change did to the health of the package', () => {
  it('reports a dropped embedding as introduced, not just as a removed part', () => {
    // The regression case. A structural diff calls this "one part removed", which
    // looks like any other removed part. The finding says the document still renders
    // correctly and is broken anyway.
    const after = healthy();
    delete after['word/embeddings/oleObject1.bin'];
    const delta = diffFindings(healthy(), after);

    expect(delta.introduced.some(f => f.code === 'ole/data-part-missing')).toBe(true);
    expect(delta.introduced.every(f => f.silent)).toBe(true);
    expect(delta.resolved).toEqual([]);
  });

  it('reports a repair as resolved', () => {
    const before = healthy();
    delete before['word/embeddings/oleObject1.bin'];
    const delta = diffFindings(before, healthy());

    expect(delta.resolved.some(f => f.code === 'ole/data-part-missing')).toBe(true);
    expect(delta.introduced).toEqual([]);
  });

  it('does not blame a change for a fault that predates it', () => {
    // Without this the report attributes every pre-existing problem to whoever last
    // touched the file.
    const broken = healthy();
    delete broken['word/embeddings/oleObject1.bin'];
    const alsoBroken = { ...broken, 'word/extra.xml': '<x/>' };
    const delta = diffFindings(broken, alsoBroken);

    expect(delta.introduced.some(f => f.code === 'ole/data-part-missing')).toBe(false);
    expect(delta.unchanged.some(f => f.code === 'ole/data-part-missing')).toBe(true);
  });

  it('reports nothing when the health of both sides is identical', () => {
    const delta = diffFindings(healthy(), healthy());

    expect(delta.introduced).toEqual([]);
    expect(delta.resolved).toEqual([]);
  });

  it('matches findings on identity rather than on wording', () => {
    // Messages interpolate counts and paths, so two genuinely identical faults can
    // word themselves differently. Code, part and subject are what make it the same
    // finding; including the message would report every fault as both fixed and
    // reintroduced.
    const before = healthy();
    delete before['word/embeddings/oleObject1.bin'];
    const after = { ...before, 'word/unrelated.xml': '<x/>' };
    const delta = diffFindings(before, after);

    expect(delta.unchanged.some(f => f.code === 'ole/data-part-missing')).toBe(true);
  });
});

describe('resilience and identity, pinned', () => {
  it('keeps every other analyzer when one throws', () => {
    // Forced rather than simulated: the real analyzers all tolerate malformed input by
    // returning nothing, so only an injected failure exercises this path.
    const exploding = {
      id: 'boom',
      title: 'Always throws',
      formats: ['docx'] as const,
      determines: ['nothing'],
      cannotDetermine: ['anything'],
      appliesTo: () => true,
      analyze: () => {
        throw new Error('analyzer bug');
      }
    };
    const run = analyzePackage(healthy(), [exploding, ...ANALYZERS]);

    expect(run.ran).toContain('package');
    expect(run.ran).toContain('ole');
  });

  it('does not report a finding as fixed-and-reintroduced when only its wording changed', () => {
    // comments-part-missing interpolates the anchor COUNT, so adding a second comment
    // rewrites the message while the fault stays the same fault. Including the message
    // in the identity would report one resolved and one introduced.
    const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const withAnchors = (n: number): PackageParts => ({
      'word/document.xml':
        `<?xml version="1.0"?><w:document ${W_NS}><w:body><w:p>` +
        Array.from({ length: n }, (_, i) =>
          `<w:commentRangeStart w:id="${i}"/><w:r><w:t>x</w:t></w:r><w:commentRangeEnd w:id="${i}"/>` +
          `<w:r><w:commentReference w:id="${i}"/></w:r>`
        ).join('') +
        `</w:p></w:body></w:document>`
    });

    const delta = diffFindings(withAnchors(1), withAnchors(2));
    const partMissing = (f: { code: string }) => f.code === 'comment/comments-part-missing';

    expect(delta.unchanged.some(partMissing)).toBe(true);
    expect(delta.introduced.some(partMissing)).toBe(false);
    expect(delta.resolved.some(partMissing)).toBe(false);
  });

  it('runs more than one analyzer, and each contributes', () => {
    // Pins the whole point of the registry: a package with two different kinds of
    // fault must produce findings from two different analyzers.
    const parts = healthy();
    delete parts['word/embeddings/oleObject1.bin'];
    delete parts['[Content_Types].xml'];

    const analyzers = new Set(analyzePackage(parts).findings.map(f => analyzerOf(f.code)));

    expect(analyzers.size).toBeGreaterThan(1);
    expect(analyzers).toContain('ole');
    expect(analyzers).toContain('package');
  });
});

describe('the part is part of a finding identity', () => {
  it('treats the same fault in a different part as a different finding', () => {
    // A bookmark named "Ref" left unclosed in the body, then instead left unclosed in
    // a header. Same code, same subject, different part. If identity ignored the part
    // this would report as unchanged and a regression that MOVED would be invisible.
    const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const unclosed = `<w:p><w:bookmarkStart w:id="1" w:name="Ref"/><w:r><w:t>x</w:t></w:r></w:p>`;
    const story = (body: string) => `<?xml version="1.0"?><w:document ${W_NS}><w:body>${body}</w:body></w:document>`;

    const delta = diffFindings(
      { 'word/document.xml': story(unclosed), 'word/header1.xml': story('') },
      { 'word/document.xml': story(''), 'word/header1.xml': story(unclosed) }
    );

    const unmatched = (f: { code: string }) => f.code === 'bookmark/unmatched-start';
    expect(delta.introduced.filter(unmatched).map(f => f.part)).toEqual(['word/header1.xml']);
    expect(delta.resolved.filter(unmatched).map(f => f.part)).toEqual(['word/document.xml']);
    expect(delta.unchanged.some(unmatched)).toBe(false);
  });
});

describe('explain routing', () => {
  it('covers every part the hand-maintained table used to', () => {
    // A regression guard for the migration off ANALYSIS_TARGETS: each of these routed
    // to an analyzer before, and losing one would silently drop the Verified tier for
    // that part rather than failing loudly.
    const covered = [
      'word/document.xml',
      'word/header1.xml',
      'word/footnotes.xml',
      'xl/worksheets/sheet1.xml',
      'ppt/slides/slide1.xml',
      'word/charts/chart1.xml',
      'ppt/charts/chart2.xml'
    ];

    for (const path of covered) {
      expect(explainersFor(path).length, path).toBeGreaterThan(0);
    }
  });

  it('routes one part to several analyzers, which is the point', () => {
    // word/document.xml carries formatting AND bookmarks AND comments AND possibly OLE.
    const ids = explainersFor('word/document.xml').map(a => a.id);

    expect(ids).toContain('word-formatting');
    expect(ids).toContain('bookmark');
    expect(ids).toContain('comment');
    expect(ids.length).toBeGreaterThan(2);
  });

  it('returns nothing for a part no analyzer covers', () => {
    // The signal the panel uses to record a coverage gap. Silence here must mean "no
    // check applies", not "nothing is wrong".
    expect(explainersFor('customXml/item1.xml')).toEqual([]);
    expect(explainersFor('docProps/core.xml')).toEqual([]);
  });

  it('keeps registry order, so an evidence bundle does not reshuffle between runs', () => {
    const first = explainersFor('word/document.xml').map(a => a.id);
    const second = explainersFor('word/document.xml').map(a => a.id);

    expect(first).toEqual(second);
  });
});

describe('siblingsFor', () => {
  const available = [
    'word/document.xml',
    'word/styles.xml',
    'word/numbering.xml',
    'word/comments.xml',
    'word/commentsExtended.xml',
    'word/_rels/document.xml.rels',
    'word/embeddings/oleObject1.bin',
    'word/media/image1.emf',
    'docProps/app.xml'
  ];

  it('unions what every matching analyzer wants, without duplicates', () => {
    const siblings = siblingsFor('word/document.xml', available);

    expect(siblings).toContain('word/styles.xml');
    expect(siblings).toContain('word/commentsExtended.xml');
    expect(siblings).toContain('word/_rels/document.xml.rels');
    expect(new Set(siblings).size).toBe(siblings.length);
  });

  it('never asks for the open part itself', () => {
    expect(siblingsFor('word/document.xml', available)).not.toContain('word/document.xml');
  });

  it('drops a named sibling the package does not have', () => {
    // The analyses report an absent part themselves; requesting one that cannot be
    // supplied just wastes a fetch and muddies the failure.
    const siblings = siblingsFor('word/document.xml', ['word/document.xml', 'word/styles.xml']);

    expect(siblings).toEqual(['word/styles.xml']);
  });

  it('does not pull in unrelated parts', () => {
    expect(siblingsFor('word/document.xml', available)).not.toContain('docProps/app.xml');
  });
});

describe('explainPart', () => {
  it('merges what several analyzers say about one part', () => {
    const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const parts: PackageParts = {
      'word/document.xml': `<?xml version="1.0"?><w:document ${W_NS}><w:body><w:p>
        <w:bookmarkStart w:id="1" w:name="Ref"/><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>`
    };

    const result = explainPart(parts, 'word/document.xml', '<w:bookmarkStart w:id="1" w:name="Ref"/>');

    expect(result).not.toBeNull();
    expect(result!.contributors).toContain('bookmark');
    expect(result!.evidence.lines.some(l => l.includes('never closes'))).toBe(true);
  });

  it('returns null for a part nothing covers, rather than an empty bundle', () => {
    // An empty bundle would read as "checked, nothing found" and could earn a tier.
    expect(explainPart({ 'docProps/core.xml': '<x/>' }, 'docProps/core.xml', '<x/>')).toBeNull();
  });

  it('survives one analyzer throwing and still returns the others', () => {
    const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    // styles.xml present but malformed: the cascade analyzer has to cope, and whatever
    // it does must not take the bookmark evidence down with it.
    const parts: PackageParts = {
      'word/document.xml': `<?xml version="1.0"?><w:document ${W_NS}><w:body><w:p>
        <w:bookmarkStart w:id="1" w:name="Ref"/></w:p></w:body></w:document>`,
      'word/styles.xml': '<w:styles><unclosed>'
    };

    const result = explainPart(parts, 'word/document.xml', '<w:bookmarkStart w:id="1" w:name="Ref"/>');

    expect(result?.contributors).toContain('bookmark');
  });

  it('de-duplicates lines two analyzers both produce', () => {
    const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const parts: PackageParts = {
      'word/document.xml': `<?xml version="1.0"?><w:document ${W_NS}><w:body><w:p>
        <w:bookmarkStart w:id="1" w:name="Ref"/></w:p></w:body></w:document>`
    };

    const lines = explainPart(parts, 'word/document.xml', '')!.evidence.lines;

    expect(new Set(lines).size).toBe(lines.length);
  });
});

describe('routing edge cases, pinned with stubs', () => {
  const stub = (id: string, over: Partial<Analyzer['explain']> = {}, lines: string[] = [`${id} says something`]) =>
    ({
      id,
      title: id,
      formats: ['docx'] as const,
      determines: ['x'],
      cannotDetermine: ['y'],
      appliesTo: () => true,
      explain: {
        matches: () => true,
        compute: () => ({ lines, unresolved: [] }),
        ...over
      }
    }) as unknown as Analyzer;

  it('never returns the open part as its own sibling', () => {
    // Some sibling patterns legitimately match the part they are attached to -
    // ppt/slides/slide1.xml matches the PowerPoint pattern - so the guard is
    // load-bearing rather than defensive.
    const registry = [stub('self', { siblingPattern: /^ppt\/slides\/[^/]+$/ })];
    const siblings = siblingsFor('ppt/slides/slide1.xml', ['ppt/slides/slide1.xml', 'ppt/slides/slide2.xml'], registry);

    expect(siblings).toEqual(['ppt/slides/slide2.xml']);
  });

  it('keeps the other analyzers when one explainer throws', () => {
    const exploding = stub('boom');
    exploding.explain!.compute = () => {
      throw new Error('explainer bug');
    };
    const result = explainPart({ 'word/document.xml': '<x/>' }, 'word/document.xml', '', [exploding, stub('ok')]);

    expect(result?.contributors).toEqual(['ok']);
  });

  it('de-duplicates a line two analyzers both produce', () => {
    // Real overlap: bookmarks and the Word cascade both read document.xml and can
    // reach the same conclusion. A repeated line wastes prompt budget and reads as
    // corroboration when it is one fact stated twice.
    const registry = [stub('a', {}, ['the same line']), stub('b', {}, ['the same line'])];
    const result = explainPart({ 'word/document.xml': '<x/>' }, 'word/document.xml', '', registry);

    expect(result!.evidence.lines).toEqual(['the same line']);
    expect(result!.contributors).toEqual(['a', 'b']);
  });
});
