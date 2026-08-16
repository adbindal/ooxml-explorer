import { describe, it, expect } from 'vitest';
import {
  checkPackageIntegrity,
  relsPathFor,
  resolveTarget,
  type PackageParts
} from '../services/packageIntegrity';

/** A minimal but structurally valid .docx package to mutate in each test. */
const validPackage = (): PackageParts => ({
  '[Content_Types].xml': `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>`,
  '_rels/.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://purl.org/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  'word/document.xml': `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p>
    <w:sectPr><w:headerReference r:id="rId10"/></w:sectPr>
  </w:body>
</w:document>`,
  'word/_rels/document.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId10" Type="http://purl.org/header" Target="header1.xml"/>
</Relationships>`,
  'word/header1.xml': `<?xml version="1.0"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p><w:r><w:drawing><a:blip xmlns:a="http://x" r:embed="rId20"/></w:drawing></w:r></w:p>
</w:hdr>`,
  'word/_rels/header1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId20" Type="http://purl.org/image" Target="../media/image1.png"/>
</Relationships>`,
  'word/media/image1.png': '',
  'media/image1.png': ''
});

const rules = (parts: PackageParts) => checkPackageIntegrity(parts).map(f => f.rule);

describe('path helpers', () => {
  it('puts a part\'s relationships in its own _rels directory', () => {
    expect(relsPathFor('word/document.xml')).toBe('word/_rels/document.xml.rels');
    expect(relsPathFor('word/header1.xml')).toBe('word/_rels/header1.xml.rels');
    expect(relsPathFor('xl/worksheets/sheet1.xml')).toBe('xl/worksheets/_rels/sheet1.xml.rels');
  });

  it('resolves relative targets against the owning part\'s directory', () => {
    expect(resolveTarget('word/document.xml', 'header1.xml')).toBe('word/header1.xml');
    expect(resolveTarget('word/document.xml', 'media/image1.png')).toBe('word/media/image1.png');
  });

  it('resolves parent-relative targets, which headers routinely use', () => {
    expect(resolveTarget('word/header1.xml', '../media/image1.png')).toBe('media/image1.png');
  });

  it('resolves absolute targets', () => {
    expect(resolveTarget('word/document.xml', '/word/styles.xml')).toBe('word/styles.xml');
  });

  it('resolves package-root targets, which _rels/.rels uses', () => {
    // The package-level relationships part has no owning part, so its targets are
    // relative to the root rather than to any directory.
    expect(resolveTarget('', 'word/document.xml')).toBe('word/document.xml');
    expect(resolveTarget('', 'xl/workbook.xml')).toBe('xl/workbook.xml');
  });
});

describe('package-level relationships', () => {
  it('validates the root _rels/.rels target', () => {
    const parts = validPackage();
    parts['_rels/.rels'] = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://purl.org/officeDocument" Target="word/missing.xml"/>
</Relationships>`;
    const finding = checkPackageIntegrity(parts).find(f => f.rule === 'missing-relationship-target');
    expect(finding?.message).toContain('word/missing.xml');
  });

  it('accepts a correct root relationship', () => {
    expect(rules(validPackage())).not.toContain('missing-relationship-target');
  });
});

describe('a well-formed package', () => {
  it('produces no findings', () => {
    expect(checkPackageIntegrity(validPackage())).toEqual([]);
  });
});

describe('content types', () => {
  it('flags a package with no [Content_Types].xml', () => {
    const parts = validPackage();
    delete parts['[Content_Types].xml'];
    expect(rules(parts)).toContain('missing-content-types');
  });

  it('flags a part that is declared nowhere', () => {
    // Adding a header without its Override is one of the most common ways a
    // generated document becomes unopenable.
    const parts = validPackage();
    parts['word/footer9.docbin'] = '';
    const finding = checkPackageIntegrity(parts).find(f => f.rule === 'untyped-part');
    expect(finding?.part).toBe('word/footer9.docbin');
    expect(finding?.message).toContain('docbin');
  });

  it('accepts a part covered by a Default extension rather than an Override', () => {
    const parts = validPackage();
    parts['word/extra.png'] = '';
    expect(rules(parts)).not.toContain('untyped-part');
  });

  it('does not treat [Content_Types].xml as a part needing its own declaration', () => {
    expect(rules(validPackage())).not.toContain('untyped-part');
  });

  it('flags malformed content types rather than throwing', () => {
    const parts = validPackage();
    parts['[Content_Types].xml'] = '<Types><unclosed>';
    expect(rules(parts)).toContain('malformed-xml');
  });
});

describe('relationship references', () => {
  it('flags an r:id that is not declared', () => {
    const parts = validPackage();
    parts['word/_rels/document.xml.rels'] = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
    const finding = checkPackageIntegrity(parts).find(f => f.rule === 'dangling-relationship-id');
    expect(finding?.part).toBe('word/document.xml');
    expect(finding?.message).toContain('rId10');
  });

  it('does NOT accept a relationship declared in a different part\'s rels', () => {
    // The header's image is declared in document.xml.rels instead of header1.xml.rels.
    // Some readers tolerate this; Word does not. This is the check's whole point.
    const parts = validPackage();
    delete parts['word/_rels/header1.xml.rels'];
    parts['word/_rels/document.xml.rels'] = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId10" Type="http://purl.org/header" Target="header1.xml"/>
  <Relationship Id="rId20" Type="http://purl.org/image" Target="media/image1.png"/>
</Relationships>`;
    const finding = checkPackageIntegrity(parts).find(f => f.rule === 'dangling-relationship-id');
    expect(finding?.part).toBe('word/header1.xml');
    expect(finding?.message).toContain('word/_rels/header1.xml.rels');
  });

  it('detects references through r:embed, not just r:id', () => {
    const parts = validPackage();
    parts['word/_rels/header1.xml.rels'] = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
    parts['media/image1.png'] = '';
    const finding = checkPackageIntegrity(parts).find(f => f.rule === 'dangling-relationship-id');
    expect(finding?.message).toContain('rId20');
  });

  it('flags a relationship whose target is missing from the package', () => {
    const parts = validPackage();
    delete parts['media/image1.png'];
    const finding = checkPackageIntegrity(parts).find(f => f.rule === 'missing-relationship-target');
    expect(finding?.message).toContain('../media/image1.png');
  });

  it('ignores external targets, which are URLs rather than parts', () => {
    const parts = validPackage();
    parts['word/_rels/document.xml.rels'] = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId10" Type="http://purl.org/header" Target="header1.xml"/>
  <Relationship Id="rId11" Type="http://purl.org/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>`;
    expect(rules(parts)).not.toContain('missing-relationship-target');
  });

  it('warns about a rels part whose owner is absent', () => {
    const parts = validPackage();
    delete parts['word/header1.xml'];
    const finding = checkPackageIntegrity(parts).find(f => f.rule === 'orphaned-rels-part');
    expect(finding?.severity).toBe('warning');
  });
});

describe('reporting', () => {
  it('orders errors before warnings', () => {
    const parts = validPackage();
    delete parts['word/header1.xml'];
    parts['word/stray.docbin'] = '';
    const severities = checkPackageIntegrity(parts).map(f => f.severity);
    expect(severities.indexOf('error')).toBeLessThan(severities.indexOf('warning'));
  });

  it('is format-agnostic - the same checks run against a spreadsheet', () => {
    const xlsx: PackageParts = {
      '[Content_Types].xml': `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`,
      'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`
    };
    const finding = checkPackageIntegrity(xlsx).find(f => f.rule === 'dangling-relationship-id');
    expect(finding?.part).toBe('xl/workbook.xml');
    expect(finding?.message).toContain('xl/_rels/workbook.xml.rels');
  });
});
