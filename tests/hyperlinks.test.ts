import { describe, it, expect } from 'vitest';
import {
  readHyperlinks,
  hyperlinkFindings,
  hyperlinkResolves,
  findSilentlyBrokenHyperlinks,
  computeHyperlinkEvidenceForMarkup,
  HYPERLINK_HOST_PART
} from '../services/hyperlinks';
import type { PackageParts } from '../services/packageIntegrity';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const X = 'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const rels = (body: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;

/** A hyperlink relationship, external unless told otherwise. */
const link = (id: string, target: string, external = true) =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${target}"${
    external ? ' TargetMode="External"' : ''
  }/>`;

const slideRel = (id: string, target: string) =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="${target}"/>`;

/** word/document.xml with one external link and one internal link to a live bookmark. */
const wordPackage = (overrides: Partial<PackageParts> = {}): PackageParts => ({
  'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${R}><w:body>
      <w:p><w:hyperlink r:id="rId4" w:tooltip="Our site" w:tgtFrame="_blank"><w:r><w:t>anthropic.com</w:t></w:r></w:hyperlink></w:p>
      <w:p><w:bookmarkStart w:id="1" w:name="Chapter2"/><w:r><w:t>Chapter 2</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>
      <w:p><w:hyperlink w:anchor="Chapter2"><w:r><w:t>see chapter 2</w:t></w:r></w:hyperlink></w:p>
    </w:body></w:document>`,
  'word/_rels/document.xml.rels': rels(link('rId4', 'https://www.anthropic.com/')),
  ...overrides
});

/** xl/worksheets/sheet1.xml with one external link and one location link. */
const excelPackage = (overrides: Partial<PackageParts> = {}): PackageParts => ({
  'xl/workbook.xml': `<?xml version="1.0"?><x:workbook ${X} ${R}><x:sheets>
      <x:sheet name="Sheet1" sheetId="1" r:id="rId1"/><x:sheet name="Summary" sheetId="2" r:id="rId2"/>
    </x:sheets><x:definedNames><x:definedName name="Total_Sales">Summary!$B$2</x:definedName></x:definedNames></x:workbook>`,
  'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><x:worksheet ${X} ${R}><x:sheetData>
      <x:row r="1"><x:c r="A1" t="s"><x:v>0</x:v></x:c><x:c r="B1"><x:v>42</x:v></x:c></x:row>
      <x:row r="5"><x:c r="A5" s="3"/></x:row>
    </x:sheetData><x:hyperlinks>
      <x:hyperlink ref="A1" r:id="rId9" tooltip="Docs"/>
      <x:hyperlink ref="B1" location="Summary!A1"/>
    </x:hyperlinks></x:worksheet>`,
  'xl/worksheets/_rels/sheet1.xml.rels': rels(link('rId9', 'https://example.com/docs')),
  ...overrides
});

/** ppt/slides/slide1.xml with a slide jump and an external link. */
const slidePackage = (overrides: Partial<PackageParts> = {}): PackageParts => ({
  'ppt/slides/slide1.xml': `<?xml version="1.0"?><p:sld ${P} ${A} ${R}><p:cSld><p:spTree>
      <p:sp><p:nvSpPr><p:cNvPr id="2" name="Jump">
        <a:hlinkClick r:id="rId2" action="ppaction://hlinksldjump"/>
      </p:cNvPr></p:nvSpPr></p:sp>
      <p:sp><p:txBody><a:p><a:r><a:rPr><a:hlinkClick r:id="rId3"/></a:rPr><a:t>site</a:t></a:r></a:p></p:txBody></p:sp>
    </p:spTree></p:cSld></p:sld>`,
  'ppt/slides/_rels/slide1.xml.rels': rels(slideRel('rId2', '../slides/slide2.xml') + link('rId3', 'https://example.com/')),
  'ppt/slides/slide2.xml': `<?xml version="1.0"?><p:sld ${P}/>`,
  ...overrides
});

describe('the link that goes nowhere and looks exactly like the one that works', () => {
  it('reports a healthy Word document with no problems', () => {
    const links = readHyperlinks(wordPackage(), 'word/document.xml');

    expect(links).toHaveLength(2);
    expect(links.flatMap(l => l.problems)).toEqual([]);
    expect(links[0].kind).toBe('external');
    expect(links[1].kind).toBe('internal');
  });

  it('catches an anchor whose bookmark is gone — the headline case', () => {
    // The bookmark markers are deleted; the hyperlink is untouched. The paragraph text
    // is identical, the run is still styled as a link, and clicking does nothing.
    const parts = wordPackage({
      'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${R}><w:body>
        <w:p><w:r><w:t>Chapter 2</w:t></w:r></w:p>
        <w:p><w:hyperlink w:anchor="Chapter2"><w:r><w:t>see chapter 2</w:t></w:r></w:hyperlink></w:p>
      </w:body></w:document>`
    });
    const [internal] = readHyperlinks(parts, 'word/document.xml');

    const problem = internal.problems.find(p => p.code === 'hyperlink/dangling-anchor');
    expect(problem?.silent).toBe(true);
    expect(problem?.severity).toBe('error');
    expect(problem?.subject).toEqual({ anchor: 'Chapter2' });
    expect(hyperlinkResolves(internal)).toBe(false);
  });

  it('does not report an anchor whose bookmark opens but never closes', () => {
    // readBookmarks still indexes an unmatched start by name, and wordBookmarks already
    // reports it. Reporting it a second time here as a dangling anchor would be a
    // different, wrong story: the name IS declared.
    const parts = wordPackage({
      'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${R}><w:body>
        <w:p><w:bookmarkStart w:id="1" w:name="Chapter2"/><w:r><w:t>Chapter 2</w:t></w:r></w:p>
        <w:p><w:hyperlink w:anchor="Chapter2"><w:r><w:t>see</w:t></w:r></w:hyperlink></w:p>
      </w:body></w:document>`
    });
    const [internal] = readHyperlinks(parts, 'word/document.xml');

    expect(internal.problems).toEqual([]);
  });

  it('finds a bookmark declared in another part before calling an anchor dangling', () => {
    // Whether Word resolves a header's anchor against the body was not confirmed, so the
    // lookup is package-wide. This pins that choice: a part-local lookup would report a
    // dangling anchor here, which would be a confident wrong answer.
    const parts: PackageParts = {
      'word/document.xml': `<?xml version="1.0"?><w:document ${W}><w:body>
        <w:p><w:bookmarkStart w:id="1" w:name="Appendix"/><w:bookmarkEnd w:id="1"/></w:p></w:body></w:document>`,
      'word/header1.xml': `<?xml version="1.0"?><w:hdr ${W} ${R}>
        <w:p><w:hyperlink w:anchor="Appendix"><w:r><w:t>appendix</w:t></w:r></w:hyperlink></w:p></w:hdr>`
    };
    const [header] = readHyperlinks(parts, 'word/header1.xml');

    expect(header.problems).toEqual([]);
    expect(hyperlinkResolves(header)).toBe(true);
  });

  it('explains a dangling anchor that no bookmark could ever have carried', () => {
    // A bookmark name is capped at 40 characters, so a longer anchor is not merely
    // missing — it is unreachable by construction, which is a different diagnosis.
    const anchor = 'A'.repeat(60);
    const parts = wordPackage({
      'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${R}><w:body>
        <w:p><w:hyperlink w:anchor="${anchor}"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p></w:body></w:document>`
    });
    const [internal] = readHyperlinks(parts, 'word/document.xml');

    expect(internal.problems[0].message).toContain('cannot exceed 40');
  });

  it('does not report Word\'s reserved _top anchor as dangling', () => {
    // _top is never a bookmark. Checking it against the index would report a dangling
    // anchor on every document that uses Word's own "Top of the Document" destination.
    const parts = wordPackage({
      'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${R}><w:body>
        <w:p><w:hyperlink w:anchor="_top"><w:r><w:t>back to top</w:t></w:r></w:hyperlink></w:p></w:body></w:document>`
    });

    expect(readHyperlinks(parts, 'word/document.xml')[0].problems).toEqual([]);
  });

  it('lists only the links that render correctly and are broken anyway', () => {
    const parts = wordPackage({
      'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${R}><w:body>
        <w:p><w:hyperlink r:id="rId4"><w:r><w:t>ok</w:t></w:r></w:hyperlink></w:p>
        <w:p><w:hyperlink w:anchor="Gone"><w:r><w:t>broken</w:t></w:r></w:hyperlink></w:p>
      </w:body></w:document>`
    });

    expect(findSilentlyBrokenHyperlinks(readHyperlinks(parts, 'word/document.xml'))).toHaveLength(1);
  });
});

describe('external links: reported, never fetched, never guessed at', () => {
  it('reports an intact external link as unverifiable rather than as working or broken', () => {
    const [external] = readHyperlinks(wordPackage(), 'word/document.xml');

    expect(external.externalTarget).toBe('https://www.anthropic.com/');
    expect(external.problems).toEqual([]);
    // null, not true and not false: "we cannot check" is its own answer.
    expect(hyperlinkResolves(external)).toBeNull();
  });

  it('catches a relationship id that names no relationship', () => {
    const parts = wordPackage({ 'word/_rels/document.xml.rels': rels('') });
    const [external] = readHyperlinks(parts, 'word/document.xml');

    const problem = external.problems.find(p => p.code === 'hyperlink/relationship-missing');
    expect(problem?.message).toContain('rId4');
    expect(problem?.silent).toBe(true);
    expect(hyperlinkResolves(external)).toBe(false);
  });

  it('catches a part with no relationship part at all', () => {
    const parts = wordPackage();
    delete parts['word/_rels/document.xml.rels'];
    const [external] = readHyperlinks(parts, 'word/document.xml');

    expect(external.problems[0].message).toContain('no relationship part at all');
  });

  it('catches an external relationship with an empty target', () => {
    // Checkable without leaving the package: the relationship promises a destination
    // outside the file and then names nothing at all.
    const parts = wordPackage({ 'word/_rels/document.xml.rels': rels(link('rId4', '')) });
    const [external] = readHyperlinks(parts, 'word/document.xml');

    expect(external.problems.map(p => p.code)).toEqual(['hyperlink/empty-external-target']);
    expect(hyperlinkResolves(external)).toBe(false);
  });

  it('flags a hyperlink relationship that forgot TargetMode="External"', () => {
    // OPC then resolves the Target as a package path, so every consumer follows
    // something other than the URI the author meant — and it renders identically.
    const parts = wordPackage({
      // Resolves, from word/document.xml, to word/document.xml — a part that is there.
      'word/_rels/document.xml.rels': rels(link('rId4', 'document.xml', false))
    });
    const [external] = readHyperlinks(parts, 'word/document.xml');

    const problem = external.problems.find(p => p.code === 'hyperlink/relationship-not-external');
    expect(problem?.severity).toBe('warning');
    expect(problem?.message).toContain('resolves it as a path inside the package');
  });

  it('escalates to an error when the non-external target is not in the package either', () => {
    const parts = wordPackage({
      'word/_rels/document.xml.rels': rels(link('rId4', 'nowhere/at-all.htm', false))
    });
    const [external] = readHyperlinks(parts, 'word/document.xml');

    const problem = external.problems.find(p => p.code === 'hyperlink/internal-target-missing');
    expect(problem?.severity).toBe('error');
    expect(problem?.subject?.target).toBe('word/nowhere/at-all.htm');
  });
});

describe('three formats, three ways to name a destination', () => {
  it('reads Word\'s internal destination as a bookmark name', () => {
    const [, internal] = readHyperlinks(wordPackage(), 'word/document.xml');

    expect(internal.destinationEvidence).toBe('w:hyperlink/@w:anchor');
    expect(internal.anchor).toBe('Chapter2');
  });

  it('reads Excel\'s internal destination as a location, and its position as @ref', () => {
    const links = readHyperlinks(excelPackage(), 'xl/worksheets/sheet1.xml');

    expect(links[1].destinationEvidence).toBe('x:hyperlink/@location');
    expect(links[1].anchor).toBe('Summary!A1');
    // The one format that states where the link lives, rather than wrapping it.
    expect(links[1].cellRange).toBe('B1');
    expect(links.flatMap(l => l.problems)).toEqual([]);
  });

  it('reads PowerPoint\'s internal destination as an action plus a relationship', () => {
    const [jump] = readHyperlinks(slidePackage(), 'ppt/slides/slide1.xml');

    expect(jump.label).toBe('a:hlinkClick');
    expect(jump.destinationEvidence).toBe('a:hlinkClick/@action');
    expect(jump.action).toBe('ppaction://hlinksldjump');
    expect(jump.internalTarget).toBe('ppt/slides/slide2.xml');
    expect(hyperlinkResolves(jump)).toBe(true);
  });

  it('reads the unprefixed attributes of x:hyperlink and a:hlinkClick', () => {
    // Verified against the SDK schema: on x:CT_Hyperlink and a:CT_Hyperlink every
    // attribute except r:id is in no namespace, so a namespaced lookup finds nothing.
    const [excel] = readHyperlinks(excelPackage(), 'xl/worksheets/sheet1.xml');
    const [ppt] = readHyperlinks(slidePackage(), 'ppt/slides/slide1.xml');

    expect(excel.tooltip).toBe('Docs');
    expect(ppt.action).toBe('ppaction://hlinksldjump');
  });

  it('reads the namespaced attributes of w:hyperlink', () => {
    const [external] = readHyperlinks(wordPackage(), 'word/document.xml');

    expect(external.tooltip).toBe('Our site');
    expect(external.targetFrame).toBe('_blank');
  });

  it('finds a:hlinkClick inside a Word document as well as a slide', () => {
    // DrawingML travels: a shape hyperlink in a .docx is the same element, resolved
    // against word/_rels/document.xml.rels.
    const parts: PackageParts = {
      'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${A} ${R}><w:body><w:p><w:r><w:drawing>
        <a:cNvPr id="1" name="Shape"><a:hlinkClick r:id="rId7"/></a:cNvPr>
      </w:drawing></w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(link('rId7', 'https://example.com/'))
    };
    const [shape] = readHyperlinks(parts, 'word/document.xml');

    expect(shape.format).toBe('word');
    expect(shape.label).toBe('a:hlinkClick');
    expect(shape.externalTarget).toBe('https://example.com/');
  });

  it('reads a:hlinkMouseOver and a:hlinkHover, which share a:CT_Hyperlink', () => {
    const parts = slidePackage({
      'ppt/slides/slide1.xml': `<?xml version="1.0"?><p:sld ${P} ${A} ${R}><p:cSld><p:spTree>
        <p:sp><p:nvSpPr><p:cNvPr id="2" name="s">
          <a:hlinkMouseOver r:id="rId3"/><a:hlinkHover r:id="rIdGone"/>
        </p:cNvPr></p:nvSpPr></p:sp>
      </p:spTree></p:cSld></p:sld>`
    });
    const links = readHyperlinks(parts, 'ppt/slides/slide1.xml');

    expect(links.map(l => l.label)).toEqual(['a:hlinkMouseOver', 'a:hlinkHover']);
    expect(links[1].problems.map(p => p.code)).toEqual(['hyperlink/relationship-missing']);
  });
});

describe('PowerPoint actions', () => {
  it('catches a slide jump whose slide is not in the package', () => {
    const parts = slidePackage();
    delete parts['ppt/slides/slide2.xml'];
    const [jump] = readHyperlinks(parts, 'ppt/slides/slide1.xml');

    const problem = jump.problems.find(p => p.code === 'hyperlink/internal-target-missing');
    expect(problem?.silent).toBe(true);
    expect(problem?.message).toContain('ppt/slides/slide2.xml');
    expect(hyperlinkResolves(jump)).toBe(false);
  });

  it('catches a slide jump that lands on something other than a slide', () => {
    const parts = slidePackage({
      'ppt/slides/_rels/slide1.xml.rels': rels(
        slideRel('rId2', '../slideLayouts/slideLayout1.xml') + link('rId3', 'https://example.com/')
      ),
      'ppt/slideLayouts/slideLayout1.xml': `<?xml version="1.0"?><p:sldLayout ${P}/>`
    });
    const [jump] = readHyperlinks(parts, 'ppt/slides/slide1.xml');

    expect(jump.problems.map(p => p.code)).toEqual(['hyperlink/slide-jump-off-target']);
  });

  it('catches a slide jump with no relationship to jump to', () => {
    const parts = slidePackage({
      'ppt/slides/slide1.xml': `<?xml version="1.0"?><p:sld ${P} ${A} ${R}><p:cSld><p:spTree>
        <p:sp><p:nvSpPr><p:cNvPr id="2" name="Jump">
          <a:hlinkClick action="ppaction://hlinksldjump"/>
        </p:cNvPr></p:nvSpPr></p:sp></p:spTree></p:cSld></p:sld>`
    });
    const [jump] = readHyperlinks(parts, 'ppt/slides/slide1.xml');

    expect(jump.problems.map(p => p.code)).toEqual(['hyperlink/action-needs-relationship']);
  });

  it('does not demand a relationship for a show-relative jump', () => {
    // ppaction://hlinkshowjump?jump=nextslide names its destination in the query string
    // and needs no r:id. Requiring one would report every "next slide" button as broken.
    const parts = slidePackage({
      'ppt/slides/slide1.xml': `<?xml version="1.0"?><p:sld ${P} ${A} ${R}><p:cSld><p:spTree>
        <p:sp><p:nvSpPr><p:cNvPr id="2" name="Next">
          <a:hlinkClick action="ppaction://hlinkshowjump?jump=nextslide"/>
        </p:cNvPr></p:nvSpPr></p:sp></p:spTree></p:cSld></p:sld>`
    });
    const [next] = readHyperlinks(parts, 'ppt/slides/slide1.xml');

    expect(next.problems).toEqual([]);
    expect(next.kind).toBe('action');
  });

  it('reports an unrecognised ppaction verb verbatim and judges nothing', () => {
    // The verb vocabulary is not in any schema, so an unknown one must not become a
    // finding. Reporting it and saying nothing more is the honest answer.
    const parts = slidePackage({
      'ppt/slides/slide1.xml': `<?xml version="1.0"?><p:sld ${P} ${A} ${R}><p:cSld><p:spTree>
        <p:sp><p:nvSpPr><p:cNvPr id="2" name="?">
          <a:hlinkClick action="ppaction://somethingnew?x=1"/>
        </p:cNvPr></p:nvSpPr></p:sp></p:spTree></p:cSld></p:sld>`
    });
    const [odd] = readHyperlinks(parts, 'ppt/slides/slide1.xml');

    expect(odd.problems).toEqual([]);
    expect(odd.action).toBe('ppaction://somethingnew?x=1');
  });

  it('treats an empty r:id beside @invalidUrl as "no relationship", not as a dangling one', () => {
    // The SDK documents this: a producer that cannot build a relationship for the URL
    // stores it in @invalidUrl and writes r:id="". Calling that a missing relationship
    // would report a second, wrong fault on top of the real one.
    const parts = slidePackage({
      'ppt/slides/slide1.xml': `<?xml version="1.0"?><p:sld ${P} ${A} ${R}><p:cSld><p:spTree>
        <p:sp><p:nvSpPr><p:cNvPr id="2" name="Bad">
          <a:hlinkClick r:id="" invalidUrl="htp:/broken url"/>
        </p:cNvPr></p:nvSpPr></p:sp></p:spTree></p:cSld></p:sld>`
    });
    const [bad] = readHyperlinks(parts, 'ppt/slides/slide1.xml');

    expect(bad.problems.map(p => p.code)).toEqual(['hyperlink/invalid-url']);
    expect(bad.problems[0].message).toContain('htp:/broken url');
  });
});

describe('Excel: the location, and the cells the link covers', () => {
  it('catches a location naming a sheet the workbook does not have', () => {
    const parts = excelPackage({
      'xl/workbook.xml': `<?xml version="1.0"?><x:workbook ${X}><x:sheets>
        <x:sheet name="Sheet1" sheetId="1"/></x:sheets></x:workbook>`
    });
    const [, location] = readHyperlinks(parts, 'xl/worksheets/sheet1.xml');

    const problem = location.problems.find(p => p.code === 'hyperlink/dangling-location');
    expect(problem?.subject).toEqual({ location: 'Summary!A1', sheet: 'Summary' });
    expect(hyperlinkResolves(location)).toBe(false);
  });

  it('unquotes a sheet name before looking it up', () => {
    const parts = excelPackage({
      'xl/workbook.xml': `<?xml version="1.0"?><x:workbook ${X}><x:sheets>
        <x:sheet name="Q1 Results" sheetId="1"/></x:sheets></x:workbook>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><x:worksheet ${X} ${R}><x:sheetData>
        <x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row></x:sheetData><x:hyperlinks>
        <x:hyperlink ref="A1" location="'Q1 Results'!$A$1"/></x:hyperlinks></x:worksheet>`
    });

    expect(readHyperlinks(parts, 'xl/worksheets/sheet1.xml')[0].problems).toEqual([]);
  });

  it('accepts a defined name and rejects one the workbook never defines', () => {
    const sheet = (location: string) => ({
      ...excelPackage(),
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><x:worksheet ${X} ${R}><x:sheetData>
        <x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row></x:sheetData><x:hyperlinks>
        <x:hyperlink ref="A1" location="${location}"/></x:hyperlinks></x:worksheet>`
    });

    expect(readHyperlinks(sheet('Total_Sales'), 'xl/worksheets/sheet1.xml')[0].problems).toEqual([]);
    expect(readHyperlinks(sheet('Total_Profit'), 'xl/worksheets/sheet1.xml')[0].problems.map(p => p.code)).toEqual([
      'hyperlink/dangling-location'
    ]);
  });

  it('accepts a bare cell address as a destination in the same sheet', () => {
    const parts = {
      ...excelPackage(),
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><x:worksheet ${X} ${R}><x:sheetData>
        <x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row></x:sheetData><x:hyperlinks>
        <x:hyperlink ref="A1" location="C9"/></x:hyperlinks></x:worksheet>`
    };

    expect(readHyperlinks(parts, 'xl/worksheets/sheet1.xml')[0].problems).toEqual([]);
  });

  it('says "cannot check" rather than "broken" when the workbook part is absent', () => {
    const parts = excelPackage();
    delete parts['xl/workbook.xml'];
    const [, location] = readHyperlinks(parts, 'xl/worksheets/sheet1.xml');

    expect(location.problems).toEqual([]);
    expect(hyperlinkResolves(location)).toBeNull();
  });

  it('catches a link whose range covers no cell with any content', () => {
    // What is left after someone deletes the text a link was attached to: the region is
    // still clickable and there is nothing on the sheet to click.
    const parts = excelPackage({
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><x:worksheet ${X} ${R}><x:sheetData>
        <x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row>
        <x:row r="5"><x:c r="A5" s="3"/></x:row>
      </x:sheetData><x:hyperlinks><x:hyperlink ref="A5" r:id="rId9"/></x:hyperlinks></x:worksheet>`
    });
    const [orphan] = readHyperlinks(parts, 'xl/worksheets/sheet1.xml');

    const problem = orphan.problems.find(p => p.code === 'hyperlink/empty-cell-range');
    expect(problem?.silent).toBe(true);
    expect(problem?.subject).toEqual({ ref: 'A5' });
  });

  it('counts a formula cell as content even with no cached value', () => {
    const parts = excelPackage({
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><x:worksheet ${X} ${R}><x:sheetData>
        <x:row r="5"><x:c r="A5"><x:f>SUM(B1:B9)</x:f></x:c></x:row>
      </x:sheetData><x:hyperlinks><x:hyperlink ref="A5" r:id="rId9"/></x:hyperlinks></x:worksheet>`
    });

    expect(readHyperlinks(parts, 'xl/worksheets/sheet1.xml')[0].problems).toEqual([]);
  });

  it('finds content anywhere inside a multi-cell range', () => {
    const parts = excelPackage({
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><x:worksheet ${X} ${R}><x:sheetData>
        <x:row r="3"><x:c r="C3"><x:v>7</x:v></x:c></x:row>
      </x:sheetData><x:hyperlinks><x:hyperlink ref="A1:E5" r:id="rId9"/></x:hyperlinks></x:worksheet>`
    });

    expect(readHyperlinks(parts, 'xl/worksheets/sheet1.xml')[0].problems).toEqual([]);
  });

  it('does not treat a cell just outside the range as content', () => {
    // The range check has to be a rectangle, not "is there any cell at all". F6 is one
    // column and one row past E5.
    const parts = excelPackage({
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><x:worksheet ${X} ${R}><x:sheetData>
        <x:row r="6"><x:c r="F6"><x:v>7</x:v></x:c></x:row>
      </x:sheetData><x:hyperlinks><x:hyperlink ref="A1:E5" r:id="rId9"/></x:hyperlinks></x:worksheet>`
    });

    expect(readHyperlinks(parts, 'xl/worksheets/sheet1.xml')[0].problems.map(p => p.code)).toEqual([
      'hyperlink/empty-cell-range'
    ]);
  });

  it('parses multi-letter columns rather than comparing letters', () => {
    // AA is column 27, which sorts before "B" as a string and after it as a column.
    const parts = excelPackage({
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><x:worksheet ${X} ${R}><x:sheetData>
        <x:row r="1"><x:c r="AA1"><x:v>7</x:v></x:c></x:row>
      </x:sheetData><x:hyperlinks>
        <x:hyperlink ref="Z1:AB1" r:id="rId9"/><x:hyperlink ref="A1:B1" location="C1"/>
      </x:hyperlinks></x:worksheet>`
    });
    const links = readHyperlinks(parts, 'xl/worksheets/sheet1.xml');

    expect(links[0].problems).toEqual([]);
    expect(links[1].problems.map(p => p.code)).toEqual(['hyperlink/empty-cell-range']);
  });

  it('reports a missing @ref, which the schema makes required', () => {
    const parts = excelPackage({
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><x:worksheet ${X} ${R}><x:sheetData/>
        <x:hyperlinks><x:hyperlink r:id="rId9"/></x:hyperlinks></x:worksheet>`
    });
    const problem = readHyperlinks(parts, 'xl/worksheets/sheet1.xml')[0].problems[0];

    expect(problem.code).toBe('hyperlink/missing-cell-range');
    // NOT marked silent: what Excel does with a ref-less hyperlink was never verified,
    // and `silent` claims the file still renders exactly as intended.
    expect(problem.silent).toBe(false);
  });

  it('does not claim an empty range when @ref cannot be parsed as one', () => {
    const parts = excelPackage({
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><x:worksheet ${X} ${R}><x:sheetData/>
        <x:hyperlinks><x:hyperlink ref="#REF!" r:id="rId9"/></x:hyperlinks></x:worksheet>`
    });

    expect(readHyperlinks(parts, 'xl/worksheets/sheet1.xml')[0].problems).toEqual([]);
  });
});

describe('ambiguity and absence', () => {
  it('flags a w:hyperlink carrying both r:id and w:anchor', () => {
    const parts = wordPackage({
      'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${R}><w:body>
        <w:p><w:bookmarkStart w:id="1" w:name="Chapter2"/><w:bookmarkEnd w:id="1"/></w:p>
        <w:p><w:hyperlink r:id="rId4" w:anchor="Chapter2"><w:r><w:t>both</w:t></w:r></w:hyperlink></w:p>
      </w:body></w:document>`
    });
    const [both] = readHyperlinks(parts, 'word/document.xml');

    const problem = both.problems.find(p => p.code === 'hyperlink/ambiguous-destination');
    expect(problem?.severity).toBe('warning');
    expect(problem?.message).toContain('w:docLocation');
    expect(problem?.subject).toEqual({ relationshipId: 'rId4', anchor: 'Chapter2' });
  });

  it('still checks the anchor when a link is ambiguous', () => {
    const parts = wordPackage({
      'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${R}><w:body>
        <w:p><w:hyperlink r:id="rId4" w:anchor="Gone"><w:r><w:t>both</w:t></w:r></w:hyperlink></w:p>
      </w:body></w:document>`
    });
    const [both] = readHyperlinks(parts, 'word/document.xml');

    expect(both.problems.map(p => p.code).sort()).toEqual([
      'hyperlink/ambiguous-destination',
      'hyperlink/dangling-anchor'
    ]);
  });

  it('reports a hyperlink element with no destination at all, per format', () => {
    const word = readHyperlinks(
      { 'word/document.xml': `<?xml version="1.0"?><w:document ${W}><w:body><w:p><w:hyperlink><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p></w:body></w:document>` },
      'word/document.xml'
    );
    const excel = readHyperlinks(
      { 'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><x:worksheet ${X}><x:hyperlinks><x:hyperlink ref="A1"/></x:hyperlinks></x:worksheet>` },
      'xl/worksheets/sheet1.xml'
    );
    const ppt = readHyperlinks(
      { 'ppt/slides/slide1.xml': `<?xml version="1.0"?><p:sld ${P} ${A}><p:cSld><a:hlinkClick/></p:cSld></p:sld>` },
      'ppt/slides/slide1.xml'
    );

    for (const [dead] of [word, excel, ppt]) {
      expect(dead.problems.map(p => p.code)).toContain('hyperlink/no-destination');
      expect(dead.kind).toBe('none');
      expect(hyperlinkResolves(dead)).toBe(false);
    }
  });
});

describe('hyperlinkFindings — the part-level view', () => {
  it('reports a hyperlink relationship nothing in the part references', () => {
    const parts = wordPackage({
      'word/_rels/document.xml.rels': rels(link('rId4', 'https://a.example/') + link('rId99', 'https://orphan.example/'))
    });
    const findings = hyperlinkFindings(parts, 'word/document.xml');

    const orphan = findings.find(f => f.code === 'hyperlink/unreferenced-relationship');
    expect(orphan?.severity).toBe('note');
    expect(orphan?.message).toContain('https://orphan.example/');
  });

  it('does not call a relationship orphaned when some other element uses it', () => {
    // References are gathered from every r: attribute in the part, not only from
    // hyperlink elements — otherwise anything else consuming the id reads as an orphan.
    const parts = wordPackage({
      'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${R}><w:body>
        <w:p><w:r><w:drawing><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rId99"/></w:drawing></w:r></w:p>
      </w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(link('rId99', 'https://a.example/'))
    });

    expect(hyperlinkFindings(parts, 'word/document.xml')).toEqual([]);
  });

  it('ignores relationships of other types', () => {
    const parts = wordPackage({
      'word/_rels/document.xml.rels': rels(link('rId4', 'https://a.example/') + slideRel('rId50', 'other.xml'))
    });

    expect(hyperlinkFindings(parts, 'word/document.xml').map(f => f.code)).toEqual([]);
  });

  it('collects the per-link problems too', () => {
    const parts = wordPackage({ 'word/_rels/document.xml.rels': rels('') });

    expect(hyperlinkFindings(parts, 'word/document.xml').map(f => f.code)).toContain('hyperlink/relationship-missing');
  });
});

describe('HYPERLINK_HOST_PART', () => {
  it('covers the body parts of all three formats', () => {
    for (const path of [
      'word/document.xml',
      'word/header1.xml',
      'word/footnotes.xml',
      'xl/worksheets/sheet1.xml',
      'xl/drawings/drawing1.xml',
      'ppt/slides/slide1.xml',
      'ppt/slideLayouts/slideLayout1.xml',
      'ppt/notesSlides/notesSlide1.xml'
    ]) {
      expect(HYPERLINK_HOST_PART.test(path)).toBe(true);
    }
  });

  it('excludes parts that cannot host one', () => {
    for (const path of [
      'word/styles.xml',
      'xl/workbook.xml',
      'ppt/presentation.xml',
      '[Content_Types].xml',
      'word/_rels/document.xml.rels'
    ]) {
      expect(HYPERLINK_HOST_PART.test(path)).toBe(false);
    }
  });
});

describe('tolerating input', () => {
  it('returns nothing for a part with no hyperlinks', () => {
    expect(readHyperlinks({ 'word/document.xml': `<?xml version="1.0"?><w:document ${W}><w:body/></w:document>` }, 'word/document.xml')).toEqual([]);
  });

  it('returns nothing for a part that is not in the package', () => {
    expect(readHyperlinks(wordPackage(), 'word/nope.xml')).toEqual([]);
  });

  it('returns nothing rather than throwing on malformed XML', () => {
    expect(readHyperlinks({ 'word/document.xml': '<w:document><unclosed>' }, 'word/document.xml')).toEqual([]);
    expect(hyperlinkFindings({ 'word/document.xml': '<w:document><unclosed>' }, 'word/document.xml')).toEqual([]);
  });

  it('survives a malformed relationship part', () => {
    const parts = wordPackage({ 'word/_rels/document.xml.rels': '<Relationships><unclosed>' });

    expect(readHyperlinks(parts, 'word/document.xml')[0].problems[0].code).toBe('hyperlink/relationship-missing');
  });

  it('does not mistake a same-named element in another namespace for a hyperlink', () => {
    // w:hyperlink and x:hyperlink share a local name. Matching on local name alone would
    // read a Word link as an Excel one and demand an @ref it has no business having.
    const parts: PackageParts = {
      'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${X} ${R}><w:body>
        <w:p><w:hyperlink r:id="rId4"><w:r><w:t>a</w:t></w:r></w:hyperlink></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(link('rId4', 'https://a.example/'))
    };
    const links = readHyperlinks(parts, 'word/document.xml');

    expect(links).toHaveLength(1);
    expect(links[0].label).toBe('w:hyperlink');
  });
});

describe('computeHyperlinkEvidenceForMarkup — panel wiring', () => {
  const wordXml = wordPackage()['word/document.xml'];

  it('returns null when no part in the bundle can host a hyperlink', () => {
    expect(computeHyperlinkEvidenceForMarkup({ 'word/styles.xml': '<w:styles/>' }, '<w:styles/>')).toBeNull();
  });

  it('returns null when the host parts contain no hyperlinks', () => {
    const xml = `<?xml version="1.0"?><w:document ${W}><w:body/></w:document>`;

    expect(computeHyperlinkEvidenceForMarkup({ 'word/document.xml': xml }, xml)).toBeNull();
  });

  it('leads with the reason the analyzer exists', () => {
    const evidence = computeHyperlinkEvidenceForMarkup(wordPackage(), wordXml);

    expect(evidence!.lines[0]).toContain('word/document.xml');
    expect(evidence!.lines[0]).toContain('renders identically to a working one');
  });

  it('prefers the part the user has open over the first host in the bundle', () => {
    // Key order is insertion order, so a bundle that lists the layout first would
    // otherwise report the layout's links for a question about the slide.
    const parts = slidePackage({
      'ppt/slideLayouts/slideLayout1.xml': `<?xml version="1.0"?><p:sldLayout ${P} ${A} ${R}><p:cSld><p:spTree>
        <p:sp><p:nvSpPr><p:cNvPr id="9" name="LayoutLink"><a:hlinkClick r:id="rIdL"/></p:cNvPr></p:nvSpPr></p:sp>
      </p:spTree></p:cSld></p:sldLayout>`,
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels': rels(link('rIdL', 'https://layout.example/'))
    });
    const ordered: PackageParts = {
      'ppt/slideLayouts/slideLayout1.xml': parts['ppt/slideLayouts/slideLayout1.xml'],
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels': parts['ppt/slideLayouts/_rels/slideLayout1.xml.rels'],
      ...parts
    };
    const evidence = computeHyperlinkEvidenceForMarkup(ordered, parts['ppt/slides/slide1.xml']);

    expect(evidence!.lines[0]).toContain('ppt/slides/slide1.xml');
  });

  it('sends an unverifiable external target to unresolved, not to problems', () => {
    const evidence = computeHyperlinkEvidenceForMarkup(wordPackage(), wordXml);

    expect(evidence!.unresolved.some(u => u.includes('no URL was fetched'))).toBe(true);
    expect(evidence!.lines.some(l => l.includes('nothing here was fetched'))).toBe(true);
  });

  it('does not claim silent breakage when everything that can be checked resolves', () => {
    const evidence = computeHyperlinkEvidenceForMarkup(wordPackage(), wordXml);

    expect(evidence!.lines.some(l => l.includes('broken anyway'))).toBe(false);
  });

  it('calls out links that render correctly and are broken anyway', () => {
    const xml = `<?xml version="1.0"?><w:document ${W} ${R}><w:body>
      <w:p><w:hyperlink w:anchor="Gone"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p></w:body></w:document>`;
    const evidence = computeHyperlinkEvidenceForMarkup({ 'word/document.xml': xml }, xml);

    expect(evidence!.lines.some(l => l.includes('render exactly as intended and are broken anyway'))).toBe(true);
  });

  it('caps the claim when both a relationship and an anchor are present', () => {
    const xml = `<?xml version="1.0"?><w:document ${W} ${R}><w:body>
      <w:p><w:bookmarkStart w:id="1" w:name="Ch"/><w:bookmarkEnd w:id="1"/></w:p>
      <w:p><w:hyperlink r:id="rId4" w:anchor="Ch"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p></w:body></w:document>`;
    const evidence = computeHyperlinkEvidenceForMarkup(
      { 'word/document.xml': xml, 'word/_rels/document.xml.rels': rels(link('rId4', 'https://a.example/')) },
      xml
    );

    expect(evidence!.unresolved.some(u => u.includes('was not verified against ECMA-376'))).toBe(true);
  });

  it('caps the claim when the workbook is missing and locations could not be checked', () => {
    const parts = excelPackage();
    delete parts['xl/workbook.xml'];
    const evidence = computeHyperlinkEvidenceForMarkup(parts, parts['xl/worksheets/sheet1.xml']);

    expect(evidence!.unresolved.some(u => u.includes('xl/workbook.xml is not in this bundle'))).toBe(true);
  });

  it('surfaces orphaned hyperlink relationships in the evidence', () => {
    const parts = wordPackage({
      'word/_rels/document.xml.rels': rels(link('rId4', 'https://a.example/') + link('rId99', 'https://orphan.example/'))
    });
    const evidence = computeHyperlinkEvidenceForMarkup(parts, wordXml);

    expect(evidence!.lines.some(l => l.includes('https://orphan.example/'))).toBe(true);
  });

  it('works for a sheet and for a slide as well as a document', () => {
    const excel = computeHyperlinkEvidenceForMarkup(excelPackage(), excelPackage()['xl/worksheets/sheet1.xml']);
    const ppt = computeHyperlinkEvidenceForMarkup(slidePackage(), slidePackage()['ppt/slides/slide1.xml']);

    expect(excel!.lines.some(l => l.includes('x:hyperlink') && l.includes('over B1'))).toBe(true);
    expect(ppt!.lines.some(l => l.includes('a:hlinkClick') && l.includes('ppaction://hlinksldjump'))).toBe(true);
  });
});
