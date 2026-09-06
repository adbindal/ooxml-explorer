import { describe, it, expect } from 'vitest';
import { powerpointInheritanceFindings } from '../services/powerpointInheritance';
import { analyzePackage } from '../services/analyzers';
import type { PackageParts } from '../services/packageIntegrity';

/**
 * The three PowerPoint faults that render and are broken anyway.
 *
 * Every describe block below asserts in both directions. The negative case is the one
 * that matters: an analyzer that catches every real fault and also fires on healthy
 * decks is worse than no analyzer, because a report nobody trusts is a report nobody
 * reads. Each "stays quiet" test is a deck PowerPoint is perfectly happy with.
 */

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

const parse = (xml: string): Document =>
  new DOMParser().parseFromString(xml, 'application/xml');

/** A theme whose fill list has `fills` entries and line list has `lines`. */
const theme = (fills: number, lines = 3): string =>
  `<?xml version="1.0"?><a:theme xmlns:a="${A}"><a:themeElements><a:fmtScheme>` +
  `<a:fillStyleLst>${'<a:solidFill/>'.repeat(fills)}</a:fillStyleLst>` +
  `<a:lnStyleLst>${'<a:ln/>'.repeat(lines)}</a:lnStyleLst>` +
  `<a:bgFillStyleLst><a:solidFill/></a:bgFillStyleLst>` +
  `</a:fmtScheme></a:themeElements></a:theme>`;

const themeFor = (xml: string) => () => parse(xml);
const noTheme = () => null;

const codes = (f: { code: string }[]) => f.map(x => x.code);

// --- notes -----------------------------------------------------------------

const notesMaster = (types: string[]) =>
  `<?xml version="1.0"?><p:notesMaster xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree>` +
  types.map(t => `<p:sp><p:nvSpPr><p:nvPr><p:ph type="${t}"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>`).join('') +
  `</p:spTree></p:cSld></p:notesMaster>`;

const notesSlide = (types: string[]) =>
  `<?xml version="1.0"?><p:notes xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree>` +
  types.map(t => `<p:sp><p:nvSpPr><p:nvPr><p:ph type="${t}"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>`).join('') +
  `</p:spTree></p:cSld></p:notes>`;

describe('notes placeholders that inherit nothing', () => {
  it('reports a notes placeholder whose type is absent from the master', () => {
    const parts: PackageParts = {
      'ppt/notesMasters/notesMaster1.xml': notesMaster(['body']),
      'ppt/notesSlides/notesSlide1.xml': notesSlide(['sldImg', 'body'])
    };

    const found = powerpointInheritanceFindings(parts, noTheme);
    expect(codes(found)).toEqual(['pptInheritance/notes-placeholder-unmatched']);
    expect(found[0].message).toContain('sldImg');
    expect(found[0].silent).toBe(true);
  });

  it('stays quiet when every type has a counterpart', () => {
    const parts: PackageParts = {
      'ppt/notesMasters/notesMaster1.xml': notesMaster(['sldImg', 'body']),
      'ppt/notesSlides/notesSlide1.xml': notesSlide(['sldImg', 'body'])
    };
    expect(powerpointInheritanceFindings(parts, noTheme)).toEqual([]);
  });

  it('matches on type and not on idx, which is the whole reason notes have their own matcher', () => {
    // Same types, deliberately mismatched idx. Applying the slide rule here reports a
    // fault on a correct deck.
    const parts: PackageParts = {
      'ppt/notesMasters/notesMaster1.xml':
        `<?xml version="1.0"?><p:notesMaster xmlns:p="${P}"><p:cSld><p:spTree>` +
        `<p:sp><p:nvSpPr><p:nvPr><p:ph type="body" idx="12"/></p:nvPr></p:nvSpPr></p:sp>` +
        `</p:spTree></p:cSld></p:notesMaster>`,
      'ppt/notesSlides/notesSlide1.xml':
        `<?xml version="1.0"?><p:notes xmlns:p="${P}"><p:cSld><p:spTree>` +
        `<p:sp><p:nvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr></p:sp>` +
        `</p:spTree></p:cSld></p:notes>`
    };
    expect(powerpointInheritanceFindings(parts, noTheme)).toEqual([]);
  });

  it('says nothing when the deck carries no notes master at all', () => {
    // A deck without speaker notes has no reason to have one; that is not a fault.
    const parts: PackageParts = { 'ppt/notesSlides/notesSlide1.xml': notesSlide(['body']) };
    expect(powerpointInheritanceFindings(parts, noTheme)).toEqual([]);
  });
});

// --- style references ------------------------------------------------------

const slideWithRef = (kind: string, idx: number) =>
  `<?xml version="1.0"?><p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree><p:sp>` +
  `<p:style><a:${kind} idx="${idx}"><a:schemeClr val="accent1"/></a:${kind}></p:style>` +
  `</p:sp></p:spTree></p:cSld></p:sld>`;

describe('style references that index past the theme', () => {
  it('reports a fillRef beyond the end of fillStyleLst', () => {
    const parts: PackageParts = { 'ppt/slides/slide1.xml': slideWithRef('fillRef', 5) };

    const found = powerpointInheritanceFindings(parts, themeFor(theme(3)));
    expect(codes(found)).toEqual(['pptInheritance/style-ref-out-of-range']);
    expect(found[0].subject).toMatchObject({ requested: '5', available: '3', list: 'fillStyleLst' });
  });

  it('stays quiet for an index the theme actually defines', () => {
    const parts: PackageParts = { 'ppt/slides/slide1.xml': slideWithRef('fillRef', 3) };
    expect(powerpointInheritanceFindings(parts, themeFor(theme(3)))).toEqual([]);
  });

  it('treats 0 and 1000 as "no fill" rather than out of range', () => {
    for (const idx of [0, 1000]) {
      const parts: PackageParts = { 'ppt/slides/slide1.xml': slideWithRef('fillRef', idx) };
      expect(powerpointInheritanceFindings(parts, themeFor(theme(3))), `idx=${idx}`).toEqual([]);
    }
  });

  it('applies the 1001 offset to fillRef but not to lnRef', () => {
    // 1001 is the FIRST background fill, and bgFillStyleLst has one entry, so this is
    // in range. The same number on lnRef has no offset and is far past lnStyleLst.
    const bg: PackageParts = { 'ppt/slides/slide1.xml': slideWithRef('fillRef', 1001) };
    expect(powerpointInheritanceFindings(bg, themeFor(theme(3)))).toEqual([]);

    const ln: PackageParts = { 'ppt/slides/slide1.xml': slideWithRef('lnRef', 1001) };
    const found = powerpointInheritanceFindings(ln, themeFor(theme(3)));
    expect(codes(found)).toEqual(['pptInheritance/style-ref-out-of-range']);
    expect(found[0].subject).toMatchObject({ list: 'lnStyleLst', requested: '1001' });
  });

  it('says nothing when there is no theme to check against', () => {
    // Not knowing and being wrong are different answers. Without a theme the engine
    // cannot tell whether index 5 exists, so it must not claim it does not.
    const parts: PackageParts = { 'ppt/slides/slide1.xml': slideWithRef('fillRef', 5) };
    expect(powerpointInheritanceFindings(parts, noTheme)).toEqual([]);
  });

  it('says nothing when the theme omits the list entirely', () => {
    const bare = `<?xml version="1.0"?><a:theme xmlns:a="${A}"><a:themeElements><a:fmtScheme/></a:themeElements></a:theme>`;
    const parts: PackageParts = { 'ppt/slides/slide1.xml': slideWithRef('fillRef', 5) };
    expect(powerpointInheritanceFindings(parts, themeFor(bare))).toEqual([]);
  });
});

// --- group child space -----------------------------------------------------

/** A group at (1000,1000) sized 500x500, with one child at (4000,4000). */
const groupSlide = (groupXfrmInner: string) =>
  `<?xml version="1.0"?><p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree>` +
  `<p:grpSp><p:grpSpPr><a:xfrm>${groupXfrmInner}</a:xfrm></p:grpSpPr>` +
  `<p:sp><p:spPr><a:xfrm><a:off x="4000" y="4000"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:sp>` +
  `</p:grpSp></p:spTree></p:cSld></p:sld>`;

describe('groups with no child coordinate space', () => {
  it('reports a positioned group that declares no chOff/chExt', () => {
    const parts: PackageParts = {
      'ppt/slides/slide1.xml': groupSlide('<a:off x="1000" y="1000"/><a:ext cx="500" cy="500"/>')
    };

    const found = powerpointInheritanceFindings(parts, noTheme);
    expect(codes(found)).toEqual(['pptInheritance/group-child-space-missing']);
    // Child at 4000 maps to 1000 + (4000 - 0) * 1 = 5000 on each axis: 1000 EMU of
    // drift per axis, so 2000 in total.
    expect(found[0].subject).toMatchObject({ driftEmu: '2000' });
  });

  it('stays quiet when the group declares both', () => {
    const parts: PackageParts = {
      'ppt/slides/slide1.xml': groupSlide(
        '<a:off x="1000" y="1000"/><a:ext cx="500" cy="500"/><a:chOff x="0" y="0"/><a:chExt cx="500" cy="500"/>'
      )
    };
    expect(powerpointInheritanceFindings(parts, noTheme)).toEqual([]);
  });

  it('reports a group that declares chOff but not chExt', () => {
    // Half a child space is not half a fault: the missing axis is unmapped exactly as
    // if neither were present, and the attribute that IS there makes it look intended.
    const parts: PackageParts = {
      'ppt/slides/slide1.xml': groupSlide(
        '<a:off x="1000" y="1000"/><a:ext cx="500" cy="500"/><a:chOff x="0" y="0"/>'
      )
    };
    expect(codes(powerpointInheritanceFindings(parts, noTheme)))
      .toEqual(['pptInheritance/group-child-space-missing']);
  });

  it('reports a group that declares chExt but not chOff', () => {
    const parts: PackageParts = {
      'ppt/slides/slide1.xml': groupSlide(
        '<a:off x="1000" y="1000"/><a:ext cx="500" cy="500"/><a:chExt cx="500" cy="500"/>'
      )
    };
    expect(codes(powerpointInheritanceFindings(parts, noTheme)))
      .toEqual(['pptInheritance/group-child-space-missing']);
  });

  it('stays quiet for a group that inherits its geometry', () => {
    const parts: PackageParts = {
      'ppt/slides/slide1.xml':
        `<?xml version="1.0"?><p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree>` +
        `<p:grpSp><p:grpSpPr/>` +
        `<p:sp><p:spPr><a:xfrm><a:off x="4000" y="4000"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:sp>` +
        `</p:grpSp></p:spTree></p:cSld></p:sld>`
    };
    expect(powerpointInheritanceFindings(parts, noTheme)).toEqual([]);
  });

  it('stays quiet when the identity mapping is already correct', () => {
    // chOff/chExt absent, but the group sits at the origin at unit scale, so children
    // land exactly where their own coordinates say. Nothing moves; nothing to report.
    const parts: PackageParts = {
      'ppt/slides/slide1.xml': groupSlide('<a:off x="0" y="0"/><a:ext cx="500" cy="500"/>')
    };
    expect(powerpointInheritanceFindings(parts, noTheme)).toEqual([]);
  });
});

// --- integration -----------------------------------------------------------

describe('reached through the registry', () => {
  const layoutParts = (extra: PackageParts): PackageParts => ({
    '[Content_Types].xml': '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
    'ppt/slideLayouts/slideLayout1.xml': `<?xml version="1.0"?><p:sldLayout xmlns:p="${P}"><p:cSld><p:spTree/></p:cSld></p:sldLayout>`,
    ...extra
  });

  it('the powerpoint analyzer now contributes findings, not only explanations', () => {
    // The regression this guards: these checks existed and were tested for weeks while
    // nothing in production called them, so the engine never found one on its own.
    const run = analyzePackage(layoutParts({
      'ppt/notesMasters/notesMaster1.xml': notesMaster(['body']),
      'ppt/notesSlides/notesSlide1.xml': notesSlide(['sldImg'])
    }));

    expect(run.ran).toContain('powerpoint');
    expect(codes(run.findings)).toContain('pptInheritance/notes-placeholder-unmatched');
  });

  it('a healthy deck produces no inheritance findings', () => {
    const run = analyzePackage(layoutParts({
      'ppt/notesMasters/notesMaster1.xml': notesMaster(['sldImg', 'body']),
      'ppt/notesSlides/notesSlide1.xml': notesSlide(['sldImg', 'body'])
    }));

    expect(run.findings.filter(f => f.code.startsWith('pptInheritance/'))).toEqual([]);
  });
});
