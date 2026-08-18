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

const rules = (parts: PackageParts) => checkPackageIntegrity(parts).map(f => f.code.replace(/^package\//, ''));

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
    const finding = checkPackageIntegrity(parts).find(f => f.code === 'package/missing-relationship-target');
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
    const finding = checkPackageIntegrity(parts).find(f => f.code === 'package/untyped-part');
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
    const finding = checkPackageIntegrity(parts).find(f => f.code === 'package/dangling-relationship-id');
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
    const finding = checkPackageIntegrity(parts).find(f => f.code === 'package/dangling-relationship-id');
    expect(finding?.part).toBe('word/header1.xml');
    expect(finding?.message).toContain('word/_rels/header1.xml.rels');
  });

  it('detects references through r:embed, not just r:id', () => {
    const parts = validPackage();
    parts['word/_rels/header1.xml.rels'] = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
    parts['media/image1.png'] = '';
    const finding = checkPackageIntegrity(parts).find(f => f.code === 'package/dangling-relationship-id');
    expect(finding?.message).toContain('rId20');
  });

  it('flags a relationship whose target is missing from the package', () => {
    const parts = validPackage();
    delete parts['media/image1.png'];
    const finding = checkPackageIntegrity(parts).find(f => f.code === 'package/missing-relationship-target');
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
    const finding = checkPackageIntegrity(parts).find(f => f.code === 'package/orphaned-rels-part');
    expect(finding?.severity).toBe('warning');
  });
});

// --- Implicit relationships -------------------------------------------------

const PML = 'application/vnd.openxmlformats-officedocument.presentationml';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const rels = (...entries: Array<[id: string, type: string, target: string]>) =>
  `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${entries.map(([id, type, target]) => `  <Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join('\n')}
</Relationships>`;

/**
 * A minimal .pptx whose full inheritance chain is intact:
 * slide -> layout -> master -> theme, plus notesSlide -> notesMaster.
 */
const validPresentation = (): PackageParts => ({
  '[Content_Types].xml': `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="${PML}.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="${PML}.slide+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="${PML}.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${PML}.slideMaster+xml"/>
  <Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="${PML}.notesSlide+xml"/>
  <Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="${PML}.notesMaster+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`,
  '_rels/.rels': rels(['rId1', `${REL}/officeDocument`, 'ppt/presentation.xml']),
  'ppt/presentation.xml': '<?xml version="1.0"?><p:presentation xmlns:p="http://x"/>',
  'ppt/slides/slide1.xml': '<?xml version="1.0"?><p:sld xmlns:p="http://x"/>',
  'ppt/slides/_rels/slide1.xml.rels': rels(
    ['rId1', `${REL}/slideLayout`, '../slideLayouts/slideLayout1.xml'],
    ['rId2', `${REL}/notesSlide`, '../notesSlides/notesSlide1.xml']
  ),
  'ppt/slideLayouts/slideLayout1.xml': '<?xml version="1.0"?><p:sldLayout xmlns:p="http://x"/>',
  'ppt/slideLayouts/_rels/slideLayout1.xml.rels': rels(
    ['rId1', `${REL}/slideMaster`, '../slideMasters/slideMaster1.xml']
  ),
  'ppt/slideMasters/slideMaster1.xml': '<?xml version="1.0"?><p:sldMaster xmlns:p="http://x"/>',
  'ppt/slideMasters/_rels/slideMaster1.xml.rels': rels(
    ['rId1', `${REL}/slideLayout`, '../slideLayouts/slideLayout1.xml'],
    ['rId2', `${REL}/theme`, '../theme/theme1.xml']
  ),
  'ppt/notesSlides/notesSlide1.xml': '<?xml version="1.0"?><p:notes xmlns:p="http://x"/>',
  'ppt/notesSlides/_rels/notesSlide1.xml.rels': rels(
    ['rId1', `${REL}/notesMaster`, '../notesMasters/notesMaster1.xml'],
    ['rId2', `${REL}/slide`, '../slides/slide1.xml']
  ),
  'ppt/notesMasters/notesMaster1.xml': '<?xml version="1.0"?><p:notesMaster xmlns:p="http://x"/>',
  'ppt/theme/theme1.xml': '<?xml version="1.0"?><a:theme xmlns:a="http://x"/>'
});

describe('implicit relationships', () => {
  it('produces no findings for an intact inheritance chain', () => {
    expect(checkPackageIntegrity(validPresentation())).toEqual([]);
  });

  it('flags a slide with no slideLayout relationship', () => {
    // Nothing in slide1.xml names the layout, so no rId dangles and no target is
    // missing - every other check passes and PowerPoint opens the deck silently.
    const parts = validPresentation();
    parts['ppt/slides/_rels/slide1.xml.rels'] = rels(
      ['rId2', `${REL}/notesSlide`, '../notesSlides/notesSlide1.xml']
    );
    const found = checkPackageIntegrity(parts);
    expect(found.map(f => f.code.replace(/^package\//, ''))).not.toContain('dangling-relationship-id');
    expect(found.map(f => f.code.replace(/^package\//, ''))).not.toContain('missing-relationship-target');

    const finding = found.find(f => f.code === 'package/missing-implicit-relationship');
    expect(finding?.part).toBe('ppt/slides/slide1.xml');
    expect(finding?.severity).toBe('error');
    expect(finding?.message).toContain('slideLayout');
  });

  it('flags a slide whose rels part is absent entirely', () => {
    const parts = validPresentation();
    delete parts['ppt/slides/_rels/slide1.xml.rels'];
    const finding = checkPackageIntegrity(parts).find(
      f => f.code === 'package/missing-implicit-relationship'
    );
    expect(finding?.part).toBe('ppt/slides/slide1.xml');
  });

  it('flags a slide carrying two slideLayout relationships', () => {
    // Unrepairable by inspecting the slide: the XML expresses no preference between them.
    const parts = validPresentation();
    parts['ppt/slides/_rels/slide1.xml.rels'] = rels(
      ['rId1', `${REL}/slideLayout`, '../slideLayouts/slideLayout1.xml'],
      ['rId3', `${REL}/slideLayout`, '../slideLayouts/slideLayout1.xml']
    );
    const finding = checkPackageIntegrity(parts).find(
      f => f.code === 'package/ambiguous-implicit-relationship'
    );
    expect(finding?.part).toBe('ppt/slides/slide1.xml');
    expect(finding?.message).toContain('rId1, rId3');
  });

  it('flags a layout with no slideMaster relationship', () => {
    const parts = validPresentation();
    parts['ppt/slideLayouts/_rels/slideLayout1.xml.rels'] = rels();
    const finding = checkPackageIntegrity(parts).find(
      f => f.code === 'package/missing-implicit-relationship'
    );
    expect(finding?.part).toBe('ppt/slideLayouts/slideLayout1.xml');
    expect(finding?.message).toContain('slideMaster');
  });

  it('flags a master with no theme relationship', () => {
    const parts = validPresentation();
    parts['ppt/slideMasters/_rels/slideMaster1.xml.rels'] = rels(
      ['rId1', `${REL}/slideLayout`, '../slideLayouts/slideLayout1.xml']
    );
    const finding = checkPackageIntegrity(parts).find(
      f => f.code === 'package/missing-implicit-relationship'
    );
    expect(finding?.part).toBe('ppt/slideMasters/slideMaster1.xml');
    expect(finding?.message).toContain('theme');
  });

  it('flags a notes slide with no notesMaster relationship', () => {
    const parts = validPresentation();
    parts['ppt/notesSlides/_rels/notesSlide1.xml.rels'] = rels(
      ['rId2', `${REL}/slide`, '../slides/slide1.xml']
    );
    const finding = checkPackageIntegrity(parts).find(
      f => f.code === 'package/missing-implicit-relationship'
    );
    expect(finding?.part).toBe('ppt/notesSlides/notesSlide1.xml');
    expect(finding?.message).toContain('notesMaster');
  });

  it('identifies parts by content type, not by filename convention', () => {
    // Same package, unconventional filenames. The Overrides still declare what each
    // part is, so the missing layout relationship must still be caught.
    const parts = validPresentation();
    parts['[Content_Types].xml'] = parts['[Content_Types].xml'].replace(
      '/ppt/slides/slide1.xml',
      '/ppt/decks/first.xml'
    );
    parts['ppt/decks/first.xml'] = parts['ppt/slides/slide1.xml'];
    parts['ppt/decks/_rels/first.xml.rels'] = rels();
    delete parts['ppt/slides/slide1.xml'];
    delete parts['ppt/slides/_rels/slide1.xml.rels'];
    parts['ppt/notesSlides/_rels/notesSlide1.xml.rels'] = rels(
      ['rId1', `${REL}/notesMaster`, '../notesMasters/notesMaster1.xml']
    );

    const finding = checkPackageIntegrity(parts).find(
      f => f.code === 'package/missing-implicit-relationship'
    );
    expect(finding?.part).toBe('ppt/decks/first.xml');
  });

  it('accepts ISO Strict relationship type URIs', () => {
    // Strict spells the same relationship http://purl.oclc.org/ooxml/officeDocument/
    // relationships/slideLayout. Matching the full Transitional URI would report every
    // Strict package as broken.
    const parts = validPresentation();
    parts['ppt/slides/_rels/slide1.xml.rels'] = rels(
      [
        'rId1',
        'http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout',
        '../slideLayouts/slideLayout1.xml'
      ],
      ['rId2', `${REL}/notesSlide`, '../notesSlides/notesSlide1.xml']
    );
    expect(rules(parts)).not.toContain('missing-implicit-relationship');
  });

  it('does not accept an unrelated URI that merely ends in the same word', () => {
    const parts = validPresentation();
    parts['ppt/slides/_rels/slide1.xml.rels'] = rels(
      ['rId1', 'http://example.com/custom/slideLayout', '../slideLayouts/slideLayout1.xml'],
      ['rId2', `${REL}/notesSlide`, '../notesSlides/notesSlide1.xml']
    );
    expect(rules(parts)).toContain('missing-implicit-relationship');
  });

  it('does not fire for a Word package, which has no such implicit chain', () => {
    expect(rules(validPackage())).not.toContain('missing-implicit-relationship');
    expect(rules(validPackage())).not.toContain('ambiguous-implicit-relationship');
  });

  it('does not fire for a spreadsheet package', () => {
    const xlsx: PackageParts = {
      '[Content_Types].xml': `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
      '_rels/.rels': rels(['rId1', `${REL}/officeDocument`, 'xl/workbook.xml']),
      'xl/workbook.xml': '<?xml version="1.0"?><workbook xmlns="http://x"/>',
      'xl/worksheets/sheet1.xml': '<?xml version="1.0"?><worksheet xmlns="http://x"/>'
    };
    expect(checkPackageIntegrity(xlsx)).toEqual([]);
  });

  it('reports nothing extra when the rels part is malformed', () => {
    // The malformed-xml finding is the real problem; inventing a second "no
    // slideLayout relationship" finding from unparseable XML would only mislead.
    const parts = validPresentation();
    parts['ppt/slides/_rels/slide1.xml.rels'] = '<Relationships><unclosed>';
    const found = checkPackageIntegrity(parts);
    expect(found.map(f => f.code.replace(/^package\//, ''))).toContain('malformed-xml');
    expect(found.map(f => f.code.replace(/^package\//, ''))).not.toContain('missing-implicit-relationship');
  });

  it('stays quiet when [Content_Types].xml is missing, since no part has an identity', () => {
    const parts = validPresentation();
    delete parts['[Content_Types].xml'];
    expect(rules(parts)).not.toContain('missing-implicit-relationship');
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
    const finding = checkPackageIntegrity(xlsx).find(f => f.code === 'package/dangling-relationship-id');
    expect(finding?.part).toBe('xl/workbook.xml');
    expect(finding?.message).toContain('xl/_rels/workbook.xml.rels');
  });
});
