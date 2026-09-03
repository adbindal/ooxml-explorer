import { describe, it, expect } from 'vitest';
import {
  loadPowerpointContext,
  resolveSlideChain,
  analyzeShape,
  locateShapeByMarkup,
  computePowerpointEvidenceForMarkup
} from '../services/powerpointFormattingAnalysis';
import { selectEvidenceTier } from '../services/aiService';
import { P_NAMESPACE, A_NAMESPACE } from '../services/powerpointResolver';
import type { PackageParts } from '../services/packageIntegrity';

const NS = `xmlns:p="${P_NAMESPACE}" xmlns:a="${A_NAMESPACE}"`;
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const rels = (entries: { id: string; type: string; target: string }[]) =>
  `<?xml version="1.0"?><Relationships xmlns="${REL}">${
    entries.map(e => `<Relationship Id="${e.id}" Type="${e.type}" Target="${e.target}"/>`).join('')
  }</Relationships>`;

const sp = (phAttrs: string | null, spPr = '') =>
  `<p:sp><p:nvSpPr><p:nvPr>${phAttrs === null ? '' : `<p:ph ${phAttrs}/>`}</p:nvPr></p:nvSpPr><p:spPr>${spPr}</p:spPr></p:sp>`;

const slide = (body: string) => `<?xml version="1.0"?><p:sld ${NS}><p:cSld><p:spTree>${body}</p:spTree></p:cSld></p:sld>`;
const layout = (body: string) => `<?xml version="1.0"?><p:sldLayout ${NS}><p:cSld><p:spTree>${body}</p:spTree></p:cSld></p:sldLayout>`;
const master = (body: string) => `<?xml version="1.0"?><p:sldMaster ${NS}><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:cSld><p:spTree>${body}</p:spTree></p:cSld></p:sldMaster>`;
const theme = () => `<?xml version="1.0"?><a:theme ${NS}><a:themeElements><a:clrScheme name="Office"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme></a:themeElements></a:theme>`;

const XFRM = '<a:xfrm><a:off x="100" y="200"/><a:ext cx="3000" cy="1000"/></a:xfrm>';

/** A minimal but complete deck: slide → layout → master → theme, all wired. */
const deck = (over: Partial<PackageParts> = {}): PackageParts => ({
  'ppt/slides/slide1.xml': slide(sp('type="title" idx="0"')),
  'ppt/slides/_rels/slide1.xml.rels': rels([
    { id: 'rId1', type: `${OFFICE_REL}/slideLayout`, target: '../slideLayouts/slideLayout1.xml' }
  ]),
  'ppt/slideLayouts/slideLayout1.xml': layout(sp('type="title" idx="0"', XFRM)),
  'ppt/slideLayouts/_rels/slideLayout1.xml.rels': rels([
    { id: 'rId1', type: `${OFFICE_REL}/slideMaster`, target: '../slideMasters/slideMaster1.xml' }
  ]),
  'ppt/slideMasters/slideMaster1.xml': master(sp('type="title"', XFRM)),
  'ppt/slideMasters/_rels/slideMaster1.xml.rels': rels([
    { id: 'rId1', type: `${OFFICE_REL}/theme`, target: '../theme/theme1.xml' }
  ]),
  'ppt/theme/theme1.xml': theme(),
  ...over
});

describe('walking the implicit chain', () => {
  it('resolves slide to layout to master to theme', () => {
    const chain = resolveSlideChain(deck(), 'ppt/slides/slide1.xml')!;
    expect(chain.layoutPath).toBe('ppt/slideLayouts/slideLayout1.xml');
    expect(chain.masterPath).toBe('ppt/slideMasters/slideMaster1.xml');
    expect(chain.themePath).toBe('ppt/theme/theme1.xml');
    expect(chain.problems).toEqual([]);
  });

  it('resolves parent-relative targets', () => {
    // Every hop in a real deck climbs out of its own directory.
    expect(resolveSlideChain(deck(), 'ppt/slides/slide1.xml')!.layout).not.toBeNull();
  });

  it('names the broken hop when a slide has no layout relationship', () => {
    // The failure this module exists for: nothing in slide1.xml looks wrong, because
    // there was never a reference in it to break.
    const parts = deck();
    parts['ppt/slides/_rels/slide1.xml.rels'] = rels([]);
    const chain = resolveSlideChain(parts, 'ppt/slides/slide1.xml')!;
    expect(chain.layoutPath).toBeNull();
    expect(chain.problems.join(' ')).toContain('no reference in the XML');
  });

  it('reports a missing relationship part entirely', () => {
    const parts = deck();
    delete parts['ppt/slides/_rels/slide1.xml.rels'];
    expect(resolveSlideChain(parts, 'ppt/slides/slide1.xml')!.problems.join(' '))
      .toContain('no relationship part');
  });

  it('rejects an ambiguous chain rather than picking one', () => {
    const parts = deck();
    parts['ppt/slides/_rels/slide1.xml.rels'] = rels([
      { id: 'rId1', type: `${OFFICE_REL}/slideLayout`, target: '../slideLayouts/slideLayout1.xml' },
      { id: 'rId2', type: `${OFFICE_REL}/slideLayout`, target: '../slideLayouts/slideLayout1.xml' }
    ]);
    expect(resolveSlideChain(parts, 'ppt/slides/slide1.xml')!.problems.join(' '))
      .toContain('exactly one is expected');
  });

  it('accepts Strict relationship URIs as well as Transitional', () => {
    // Matching the whole URI would report every Strict package as broken.
    const parts = deck();
    parts['ppt/slides/_rels/slide1.xml.rels'] = rels([
      { id: 'rId1', type: 'http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout', target: '../slideLayouts/slideLayout1.xml' }
    ]);
    expect(resolveSlideChain(parts, 'ppt/slides/slide1.xml')!.layoutPath).toBe('ppt/slideLayouts/slideLayout1.xml');
  });

  it('reports a layout that is referenced but absent', () => {
    const parts = deck();
    delete parts['ppt/slideLayouts/slideLayout1.xml'];
    expect(resolveSlideChain(parts, 'ppt/slides/slide1.xml')!.problems.join(' '))
      .toContain('referenced but missing');
  });
});

describe('geometry inheritance', () => {
  const analyze = (parts: PackageParts) => {
    const chain = resolveSlideChain(parts, 'ppt/slides/slide1.xml')!;
    const shape = chain.slide.getElementsByTagNameNS(P_NAMESPACE, 'sp').item(0)!;
    return analyzeShape(chain, shape);
  };

  it('inherits from the layout when the slide shape has no xfrm', () => {
    const a = analyze(deck());
    expect(a.geometrySource).toBe('layout');
    expect(a.explanation.join('\n')).toContain('which is correct');
  });

  it('uses the shape\'s own transform when it has one', () => {
    const a = analyze(deck({ 'ppt/slides/slide1.xml': slide(sp('type="title" idx="0"', XFRM)) }));
    expect(a.geometrySource).toBe('own');
    expect(a.explanation.join('\n')).toContain('overriding any inherited transform');
  });

  it('falls through to the master when the layout also inherits', () => {
    const a = analyze(deck({
      'ppt/slideLayouts/slideLayout1.xml': layout(sp('type="title" idx="0"'))
    }));
    expect(a.geometrySource).toBe('master');
  });

  it('reports unresolved geometry when nothing supplies a transform', () => {
    const a = analyze(deck({
      'ppt/slideLayouts/slideLayout1.xml': layout(sp('type="title" idx="0"')),
      'ppt/slideMasters/slideMaster1.xml': master(sp('type="title"'))
    }));
    expect(a.geometrySource).toBe('none');
    expect(a.explanation.join('\n')).toContain('no ancestor placeholder supplied a transform');
  });

  it('reports an idx with no layout counterpart', () => {
    const a = analyze(deck({ 'ppt/slides/slide1.xml': slide(sp('type="title" idx="9"')) }));
    expect(a.unresolved.join(' ')).toContain('no layout placeholder has idx=9');
  });

  it('says a non-placeholder inherits nothing', () => {
    const a = analyze(deck({ 'ppt/slides/slide1.xml': slide(sp(null, XFRM)) }));
    expect(a.placeholder).toBeNull();
    expect(a.explanation.join('\n')).toContain('not a placeholder');
  });
});

describe('scheme colours through the active map', () => {
  it('resolves a colour reference and names the map source', () => {
    const parts = deck({
      'ppt/slides/slide1.xml': slide(sp('type="title" idx="0"', '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'))
    });
    const chain = resolveSlideChain(parts, 'ppt/slides/slide1.xml')!;
    const shape = chain.slide.getElementsByTagNameNS(P_NAMESPACE, 'sp').item(0)!;
    const a = analyzeShape(chain, shape);
    expect(a.explanation.join('\n')).toContain('map from the master');
    expect(a.explanation.join('\n')).toContain('tx1 → dk1');
  });

  it('follows an inverted map rather than assuming tx1 is dark', () => {
    const inverted = master(sp('type="title"', XFRM)).replace('bg1="lt1" tx1="dk1"', 'bg1="dk1" tx1="lt1"');
    const parts = deck({
      'ppt/slideMasters/slideMaster1.xml': inverted,
      'ppt/slides/slide1.xml': slide(sp('type="title" idx="0"', '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'))
    });
    const chain = resolveSlideChain(parts, 'ppt/slides/slide1.xml')!;
    const shape = chain.slide.getElementsByTagNameNS(P_NAMESPACE, 'sp').item(0)!;
    expect(analyzeShape(chain, shape).explanation.join('\n')).toContain('tx1 → lt1');
  });
});

describe('locating a shape, and refusing to guess', () => {
  it('matches a shape by its markup', () => {
    const ctx = loadPowerpointContext(deck());
    expect(locateShapeByMarkup(ctx, sp('type="title" idx="0"'))?.slidePath).toBe('ppt/slides/slide1.xml');
  });

  it('REFUSES when two shapes share identical markup', () => {
    const ctx = loadPowerpointContext(deck({
      'ppt/slides/slide1.xml': slide(sp('type="body" idx="1"') + sp('type="body" idx="1"'))
    }));
    expect(locateShapeByMarkup(ctx, sp('type="body" idx="1"'))).toBeNull();
  });

  it('searches every slide in the deck', () => {
    const parts = deck({ 'ppt/slides/slide2.xml': slide(sp('type="body" idx="4"', XFRM)) });
    const ctx = loadPowerpointContext(parts);
    expect(locateShapeByMarkup(ctx, sp('type="body" idx="4"', XFRM))?.slidePath).toBe('ppt/slides/slide2.xml');
  });

  it('reports a deck with no slides', () => {
    expect(loadPowerpointContext({}).unresolved.join(' ')).toContain('no slide parts');
  });
});

describe('the full chain', () => {
  it('produces evidence that selects the verified tier', () => {
    const evidence = computePowerpointEvidenceForMarkup(deck(), sp('type="title" idx="0"'))!;
    expect(evidence.lines.join('\n')).toContain('Slide: ppt/slides/slide1.xml');
    expect(evidence.lines.join('\n')).toContain('inherited from the layout placeholder');
    expect(selectEvidenceTier(false, evidence)).toBe('verified');
  });

  it('caps at grounded when the layout relationship is missing', () => {
    const parts = deck();
    parts['ppt/slides/_rels/slide1.xml.rels'] = rels([]);
    const evidence = computePowerpointEvidenceForMarkup(parts, sp('type="title" idx="0"'))!;
    expect(evidence.unresolved.length).toBeGreaterThan(0);
    expect(selectEvidenceTier(false, evidence)).toBe('grounded');
  });

  it('returns null when the shape cannot be located', () => {
    expect(computePowerpointEvidenceForMarkup(deck(), sp('type="ftr" idx="88"'))).toBeNull();
  });
});
