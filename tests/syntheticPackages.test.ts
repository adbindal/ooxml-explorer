import { describe, it, expect } from 'vitest';
import { analyzePackage, capabilityLedger } from '../services/analyzers';
import { compareFindings, type Finding } from '../services/findings';
import type { PackageParts } from '../services/packageIntegrity';

/**
 * Minimal but REALISTICALLY SHAPED packages, built in memory, asserted to be clean.
 *
 * WHY THIS EXISTS, AND WHY IT IS NOT THE SAME AS `realFiles.test.ts`.
 *
 * The real-file suite is the only thing that can catch a false positive on genuine Office
 * output — and it is **gitignored and therefore absent from CI**, plus its smoke fixtures
 * only ever covered `.docx`. So the safety net was manual, local, and one-format.
 *
 * Both bugs the first real-file run exposed came from *structural* shapes rather than
 * document content, and structure can be built here:
 *
 *   - A Word package has FIVE body parts, not one. The comment analyzer read
 *     `matching(parts, WORD_BODY)[0]` and got `word/footnotes.xml`, so a comment anchored
 *     in the body looked orphaned. Every hand-written fixture in this repo has exactly one
 *     body part, which is why `[0]` was always right and nothing caught it.
 *   - A chart inside a WORKBOOK has no embedded workbook and needs none. The chart
 *     analyzer called that "cells that exist nowhere in this package" — seven warnings on
 *     a healthy file.
 *
 * Neither needed a real document. Both needed a fixture shaped like one. That is what this
 * file is: the shapes, in CI, on every commit, with no confidential data anywhere.
 *
 * It does NOT replace real files. Genuine Office output still carries content this cannot
 * anticipate — unusual field types, vendor extension namespaces, markup written by
 * something other than Office. Keep running `check-real-files` when fixtures appear.
 *
 * ⚠️ THE ASSERTION IS THE SAME ONE, AND THE SAME WAY ROUND.
 *
 * These packages are valid. The engine must claim **no error and nothing `silent`** about
 * them. A silent finding here is the engine telling someone their healthy file is broken
 * in a way they cannot check, which is the worst thing it can get wrong.
 */

/**
 * A real `[Content_Types].xml` declares every part, and the engine correctly reported all
 * of them as undeclared when this was an empty `<Types/>`. A `Default` for the `xml`
 * extension is what Office actually writes, and is how one line covers every part.
 */
const CONTENT_TYPES =
  '<?xml version="1.0"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '</Types>';

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const X = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';

/** Findings that claim breakage a reader cannot see. The only ones asserted on. */
const invisible = (findings: readonly Finding[]): Finding[] =>
  findings.filter(f => f.silent || f.severity === 'error');

const report = (name: string, findings: readonly Finding[]): string =>
  `${name} claims invisible breakage:\n` +
  [...findings].sort(compareFindings).map(f => `    [${f.code}] ${f.part} — ${f.message}`).join('\n');

// --- Word: the five-body-part shape ---------------------------------------

/**
 * A Word package with every story a real document carries.
 *
 * Key order deliberately puts `word/footnotes.xml` first, exactly as the archive of a real
 * document did. The comment is anchored in `word/document.xml`, which is LAST.
 */
const wordPackage = (): PackageParts => ({
  '[Content_Types].xml': CONTENT_TYPES,
  'word/footnotes.xml':
    `<?xml version="1.0"?><w:footnotes ${W}>` +
    `<w:footnote w:type="separator" w:id="-1"><w:p/></w:footnote>` +
    `<w:footnote w:id="1"><w:p><w:r><w:t>a note</w:t></w:r></w:p></w:footnote>` +
    `</w:footnotes>`,
  'word/header1.xml': `<?xml version="1.0"?><w:hdr ${W}><w:p><w:r><w:t>head</w:t></w:r></w:p></w:hdr>`,
  'word/footer1.xml': `<?xml version="1.0"?><w:ftr ${W}><w:p><w:r><w:t>foot</w:t></w:r></w:p></w:ftr>`,
  'word/endnotes.xml':
    `<?xml version="1.0"?><w:endnotes ${W}>` +
    `<w:endnote w:type="separator" w:id="-1"><w:p/></w:endnote></w:endnotes>`,
  'word/document.xml':
    `<?xml version="1.0"?><w:document ${W}><w:body>` +
    `<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr>` +
    `<w:commentRangeStart w:id="1"/><w:r><w:t>commented</w:t></w:r><w:commentRangeEnd w:id="1"/>` +
    `<w:r><w:commentReference w:id="1"/></w:r></w:p>` +
    `<w:p><w:bookmarkStart w:id="1" w:name="Intro"/><w:r><w:t>body</w:t></w:r>` +
    `<w:bookmarkEnd w:id="1"/></w:p>` +
    `<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`,
  'word/comments.xml':
    `<?xml version="1.0"?><w:comments ${W}>` +
    `<w:comment w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">` +
    `<w:p><w:r><w:t>a comment</w:t></w:r></w:p></w:comment></w:comments>`,
  'word/styles.xml':
    `<?xml version="1.0"?><w:styles ${W}>` +
    `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/></w:rPr></w:rPrDefault></w:docDefaults>` +
    `<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`
});

// --- Excel: the chart-inside-a-workbook shape ------------------------------

/** A workbook whose chart references one of its own sheets — no embedded workbook. */
const excelPackage = (): PackageParts => ({
  '[Content_Types].xml': CONTENT_TYPES,
  'xl/workbook.xml':
    `<?xml version="1.0"?><workbook ${X}><sheets><sheet name="Data" sheetId="1" r:id="rId1" ` +
    `xmlns:r="${OFFICE_REL}"/></sheets></workbook>`,
  // The sheet's r:id has to resolve, or package integrity reports it dangling — correctly.
  'xl/_rels/workbook.xml.rels':
    `<?xml version="1.0"?><Relationships xmlns="${RELS_NS}">` +
    `<Relationship Id="rId1" Type="${OFFICE_REL}/worksheet" Target="worksheets/sheet1.xml"/>` +
    `</Relationships>`,
  'xl/worksheets/sheet1.xml':
    `<?xml version="1.0"?><worksheet ${X}><sheetData>` +
    `<row r="1"><c r="A1" t="inlineStr"><is><t>label</t></is></c></row>` +
    `<row r="2"><c r="A2"><v>1</v></c></row>` +
    `</sheetData></worksheet>`,
  'xl/styles.xml':
    `<?xml version="1.0"?><styleSheet ${X}>` +
    `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
    `<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
    `<borders count="1"><border/></borders>` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
    `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>` +
    `</styleSheet>`,
  'xl/charts/chart1.xml':
    `<?xml version="1.0"?><c:chartSpace ${C} ${A}><c:chart><c:plotArea><c:barChart>` +
    `<c:ser><c:idx val="0"/><c:order val="0"/>` +
    `<c:val><c:numRef><c:f>Data!$A$2:$A$2</c:f>` +
    `<c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache>` +
    `</c:numRef></c:val></c:ser>` +
    `<c:axId val="1"/><c:axId val="2"/></c:barChart>` +
    `<c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx>` +
    `<c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>` +
    `</c:plotArea></c:chart></c:chartSpace>`
});

// --- PowerPoint: the slide/layout/master chain -----------------------------

const pptPackage = (): PackageParts => ({
  '[Content_Types].xml': CONTENT_TYPES,
  'ppt/presentation.xml':
    `<?xml version="1.0"?><p:presentation ${P}><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`,
  'ppt/slides/slide1.xml':
    `<?xml version="1.0"?><p:sld ${P} ${A}><p:cSld><p:spTree>` +
    `<p:sp><p:nvSpPr><p:nvPr><p:ph type="title" idx="0"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr/></p:sp>` +
    `<p:grpSp><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/>` +
    `<a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr>` +
    `<p:sp><p:spPr><a:xfrm><a:off x="10" y="10"/><a:ext cx="10" cy="10"/></a:xfrm></p:spPr></p:sp>` +
    `</p:grpSp></p:spTree></p:cSld></p:sld>`,
  'ppt/slideLayouts/slideLayout1.xml':
    `<?xml version="1.0"?><p:sldLayout ${P} ${A}><p:cSld><p:spTree>` +
    `<p:sp><p:nvSpPr><p:nvPr><p:ph type="title" idx="0"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="500" cy="500"/></a:xfrm></p:spPr></p:sp>` +
    `</p:spTree></p:cSld></p:sldLayout>`,
  'ppt/slideMasters/slideMaster1.xml':
    `<?xml version="1.0"?><p:sldMaster ${P} ${A}><p:cSld><p:spTree>` +
    `<p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="500" cy="500"/></a:xfrm></p:spPr></p:sp>` +
    `</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" ` +
    `accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" ` +
    `accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>`
});

describe('synthetic packages shaped like real Office output', () => {
  const cases: [string, () => PackageParts][] = [
    ['a Word document with five body parts', wordPackage],
    ['a workbook whose chart references its own sheet', excelPackage],
    ['a deck with a slide, layout and master', pptPackage]
  ];

  for (const [label, build] of cases) {
    it(`claims no invisible breakage about ${label}`, () => {
      const findings = analyzePackage(build()).findings;
      const bad = invisible(findings);

      expect(bad, bad.length > 0 ? report(label, bad) : '').toEqual([]);
    });
  }

  it('actually exercises a useful number of analyzers', () => {
    // Guards the way this file could rot into nothing: a package so minimal that every
    // analyzer skips would pass the assertions above while checking nothing at all.
    for (const [label, build] of cases) {
      const ledger = capabilityLedger(analyzePackage(build()));
      expect(ledger.ran.length, `${label} ran only ${ledger.ran.map(a => a.id).join(', ')}`)
        .toBeGreaterThanOrEqual(4);
    }
  });

  it('would have caught the comment analyzer reading the wrong body part', () => {
    // The specific regression, asserted specifically rather than relying on the sweep
    // above: the comment is anchored in `word/document.xml`, which is the LAST body part
    // in the package, and four other stories precede it.
    const findings = analyzePackage(wordPackage()).findings;

    expect(findings.map(f => f.code).filter(c => c.startsWith('comment/'))).toEqual([]);
  });

  it('would have caught every chart in a workbook being called data-less', () => {
    const findings = analyzePackage(excelPackage()).findings;

    expect(findings.map(f => f.code).filter(c => c.startsWith('chart/'))).toEqual([]);
  });
});
