import { describe, it, expect } from 'vitest';
import {
  reportPackage,
  reportComparison,
  summariseReport,
  serialiseReport,
  REPORT_SCHEMA_VERSION
} from '../services/report';
import type { PackageParts } from '../services/packageIntegrity';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const O = 'xmlns:o="urn:schemas-microsoft-com:office:office"';
const V = 'xmlns:v="urn:schemas-microsoft-com:vml"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const rels = (body: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;

const healthy = (): PackageParts => ({
  '[Content_Types].xml': `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
    <Default Extension="emf" ContentType="image/x-emf"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  </Types>`,
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

const broken = (): PackageParts => {
  const parts = healthy();
  delete parts['word/embeddings/oleObject1.bin'];
  return parts;
};

const AT = new Date('2026-08-19T00:00:00.000Z');

describe('the report is a contract', () => {
  it('stamps a real semver schema version, because other code will depend on these names', () => {
    // Asserting only `=== REPORT_SCHEMA_VERSION` is a tautology - it passes just as
    // happily if the constant is blank or the field is dropped from the type.
    const report = reportPackage(healthy(), AT);

    expect(REPORT_SCHEMA_VERSION).toMatch(/^\d+\.\d+\.\d+$/);
    expect(report.schemaVersion).toBe(REPORT_SCHEMA_VERSION);
    expect(JSON.parse(serialiseReport(report)).schemaVersion).toMatch(/^\d+\.\d+\.\d+$/);
  });

  it('is byte-identical across runs once the timestamp is pinned', () => {
    // What makes a report diffable in CI: everything except generatedAt is a pure
    // function of the package, so a changed report means a changed file.
    expect(serialiseReport(reportPackage(broken(), AT))).toBe(serialiseReport(reportPackage(broken(), AT)));
  });

  it('sorts findings worst-first rather than in discovery order', () => {
    // The package analyzer runs FIRST and contributes a warning here (an undeclared
    // part); the OLE analyzer runs later and contributes an error. So discovery order
    // and severity order genuinely disagree, which is the only arrangement where the
    // sort is observable at all.
    const mixed = { ...broken(), 'word/undeclared.xml': '<x/>' };
    const findings = reportPackage(mixed, AT).findings;
    const rank = { error: 0, warning: 1, note: 2 } as const;

    expect(findings.length).toBeGreaterThan(1);
    expect(new Set(findings.map(f => f.severity)).size).toBeGreaterThan(1);
    expect(findings[0].severity).toBe('error');
    for (let i = 1; i < findings.length; i++) {
      expect(rank[findings[i].severity]).toBeGreaterThanOrEqual(rank[findings[i - 1].severity]);
    }
  });

  it('round-trips through JSON with nothing lost', () => {
    const report = reportPackage(broken(), AT);

    expect(JSON.parse(serialiseReport(report))).toEqual(report);
  });
});

describe('counts', () => {
  it('counts by severity and calls out how many are silent', () => {
    const report = reportPackage(broken(), AT);

    expect(report.counts.total).toBe(report.findings.length);
    expect(report.counts.error + report.counts.warning + report.counts.note).toBe(report.counts.total);
    expect(report.counts.silent).toBeGreaterThan(0);
  });

  it('reports a healthy package as zero without claiming it is correct', () => {
    const report = reportPackage(healthy(), AT);

    expect(report.counts.total).toBe(0);
    expect(summariseReport(report)).toContain('not the same as the file being correct');
  });

  it('says when nothing found would have been visible anyway', () => {
    const summary = summariseReport(reportPackage(broken(), AT));

    expect(summary).toContain('render correctly and are broken anyway');
  });
});

describe('capabilities travel with the report', () => {
  it('names what ran, what was skipped, and what cannot be established', () => {
    const report = reportPackage(healthy(), AT);

    expect(report.capabilities.ran.length).toBeGreaterThan(0);
    expect(report.capabilities.skipped.length).toBeGreaterThan(0);
    expect(report.capabilities.limits.length).toBeGreaterThan(0);
  });

  it('carries the limits into the serialised form, not just the in-memory object', () => {
    // The consumer of this report is another program reading JSON. A limit that only
    // exists in memory tells that program nothing.
    const parsed = JSON.parse(serialiseReport(reportPackage(healthy(), AT)));

    expect(parsed.capabilities.limits.length).toBeGreaterThan(0);
  });
});

describe('comparison report', () => {
  it('separates what a change broke from what it fixed', () => {
    const report = reportComparison(healthy(), broken(), AT);

    expect(report.introduced.some(f => f.code === 'ole/data-part-missing')).toBe(true);
    expect(report.resolved).toEqual([]);
    expect(report.counts.introduced.total).toBe(report.introduced.length);
  });

  it('reports a repair the other way round', () => {
    const report = reportComparison(broken(), healthy(), AT);

    expect(report.resolved.some(f => f.code === 'ole/data-part-missing')).toBe(true);
    expect(report.introduced).toEqual([]);
  });

  it('keeps pre-existing faults out of the introduced list', () => {
    const report = reportComparison(broken(), { ...broken(), 'word/note.xml': '<x/>' }, AT);

    expect(report.introduced.some(f => f.code === 'ole/data-part-missing')).toBe(false);
    expect(report.unchanged.some(f => f.code === 'ole/data-part-missing')).toBe(true);
  });

  it('counts parts on the after side, since that is the package under review', () => {
    const after = { ...broken(), 'word/note.xml': '<x/>' };

    expect(reportComparison(broken(), after, AT).parts).toBe(Object.keys(after).length);
  });

  it('is byte-identical across runs once the timestamp is pinned', () => {
    expect(serialiseReport(reportComparison(healthy(), broken(), AT))).toBe(
      serialiseReport(reportComparison(healthy(), broken(), AT))
    );
  });
});
