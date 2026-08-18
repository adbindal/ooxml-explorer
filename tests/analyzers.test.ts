import { describe, it, expect } from 'vitest';
import { ANALYZERS, analyzePackage, capabilityLedger, diffFindings } from '../services/analyzers';
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
