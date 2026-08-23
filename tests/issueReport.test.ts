import { describe, it, expect } from 'vitest';
import { buildIssueReport, promptFor, ISSUE_KINDS } from '../services/issueReport';
import type { PackageParts } from '../services/packageIntegrity';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const AT = new Date('2026-08-23T00:00:00.000Z');

/** A package with a bookmark that opens and never closes. */
const broken = (): PackageParts => ({
  'word/document.xml': `<?xml version="1.0"?><w:document ${W}><w:body><w:p>
    <w:bookmarkStart w:id="1" w:name="ConfidentialProjectCodename"/>
    <w:r><w:t>Acme merger closes in March at $4.2bn</w:t></w:r></w:p></w:body></w:document>`
});

describe('the report never carries document content', () => {
  it('omits body text even though the analyzer read it', () => {
    // The whole tool defaults to on-device AI so confidential documents never leave the
    // machine. A bug-report button that quietly undoes that is worse than none, because
    // the person clicking it is trying to help and has no reason to suspect otherwise.
    const report = buildIssueReport(
      { kind: 'false-positive', partPath: 'word/document.xml', description: 'looks fine to me', parts: broken() },
      AT
    );

    expect(report).not.toContain('Acme');
    expect(report).not.toContain('4.2bn');
    expect(report).not.toContain('merger');
  });

  it('omits bookmark names, which are author-chosen and can be sensitive', () => {
    const report = buildIssueReport(
      { kind: 'false-positive', partPath: 'word/document.xml', description: 'x', parts: broken() },
      AT
    );

    expect(report).not.toContain('ConfidentialProjectCodename');
  });

  it('normalises the part path, so a numbered part carries no information', () => {
    const report = buildIssueReport({ kind: 'not-covered', partPath: 'word/header7.xml', description: 'x' }, AT);

    expect(report).toContain('word/header#.xml');
    expect(report).not.toContain('header7');
  });

  it('does carry engine output, which is what makes it actionable', () => {
    const report = buildIssueReport(
      { kind: 'false-positive', partPath: 'word/document.xml', description: 'x', parts: broken() },
      AT
    );

    expect(report).toContain('bookmark/unmatched-start');
    expect(report).toContain('Likely owner');
  });
});

describe('the report is useful to whoever picks it up', () => {
  it('names which analyzers ran and how many were skipped', () => {
    const report = buildIssueReport({ kind: 'missed-fault', description: 'x', parts: broken() }, AT);

    expect(report).toContain('Analyzers that ran');
    expect(report).toContain('Skipped');
  });

  it('says plainly when the engine found nothing, which is the point of a miss report', () => {
    // A genuinely clean package, not just a bare document.xml -- one of those is missing
    // [Content_Types].xml and correctly reports it, which is the engine being right.
    const clean: PackageParts = {
      '[Content_Types].xml': `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      </Types>`,
      '_rels/.rels': `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>`,
      'word/document.xml': `<?xml version="1.0"?><w:document ${W}><w:body/></w:document>`
    };
    const report = buildIssueReport({ kind: 'missed-fault', description: 'x', parts: clean }, AT);

    expect(report).toContain('Findings: **none**');
  });

  it('groups a repeated fault into one line, so it cannot flood the report', () => {
    // Five unclosed bookmarks is one bug. Five identical lines buries the ledger and the
    // description underneath them, and the person triaging stops reading.
    const many: PackageParts = {
      'word/document.xml': `<?xml version="1.0"?><w:document ${W}><w:body><w:p>` +
        Array.from({ length: 5 }, (_, i) => `<w:bookmarkStart w:id="${i}" w:name="B${i}"/>`).join('') +
        `</w:p></w:body></w:document>`
    };
    const report = buildIssueReport({ kind: 'false-positive', description: 'x', parts: many }, AT);

    expect(report).toContain('`bookmark/unmatched-start` ×5');
    expect(report.split('bookmark/unmatched-start').length - 1).toBe(1);
  });

  it('works with no package at all, for a report filed from memory', () => {
    const report = buildIssueReport({ kind: 'not-covered', description: 'SmartArt is ignored' }, AT);

    expect(report).toContain('SmartArt is ignored');
    expect(report).not.toContain('Analyzers that ran');
  });

  it('marks an empty description rather than producing a silently blank report', () => {
    expect(buildIssueReport({ kind: 'missed-fault', description: '   ' }, AT)).toContain('no description given');
  });

  it('is deterministic apart from the timestamp', () => {
    const input = { kind: 'false-positive' as const, partPath: 'word/document.xml', description: 'x', parts: broken() };

    expect(buildIssueReport(input, AT)).toBe(buildIssueReport(input, AT));
  });
});

describe('the prompts ask for what each kind actually needs', () => {
  it('asks a false-positive reporter for the Office standard', () => {
    // Generic prompts produce generic reports. This is the standard the engine is held
    // to, so it is the standard the report should cite.
    expect(promptFor('false-positive')).toContain('repair prompt');
  });

  it('asks a miss reporter whether it renders correctly anyway', () => {
    expect(promptFor('missed-fault')).toContain('renders correctly and is broken anyway');
  });

  it('offers every kind with a distinct prompt', () => {
    const prompts = ISSUE_KINDS.map(k => promptFor(k.kind));

    expect(ISSUE_KINDS.length).toBe(4);
    expect(new Set(prompts).size).toBe(4);
  });
});
