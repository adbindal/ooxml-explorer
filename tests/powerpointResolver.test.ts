import { describe, it, expect } from 'vitest';
import {
  readPlaceholders,
  matchSlideToLayout,
  matchNotesToMaster,
  matchLayoutToMaster,
  readTransform,
  rotationDegrees,
  applyGroupTransform,
  readColourMap,
  resolveSchemeColour,
  resolveColourMap,
  resolveStyleReference,
  COLOUR_SLOTS,
  COLOUR_MAP_KEYS,
  MAP_BYPASSING_VALUES,
  NO_CORRESPONDENCE_IDX,
  ROTATION_UNITS_PER_DEGREE,
  P_NAMESPACE,
  A_NAMESPACE
} from '../services/powerpointResolver';

const parse = (xml: string) => new DOMParser().parseFromString(xml, 'application/xml');
const NS = `xmlns:p="${P_NAMESPACE}" xmlns:a="${A_NAMESPACE}"`;

const part = (body: string) => parse(`<?xml version="1.0"?><p:sld ${NS}>${body}</p:sld>`);
const frag = (xml: string): Element =>
  parse(`<?xml version="1.0"?><root ${NS}>${xml}</root>`).documentElement.firstElementChild!;

const shape = (phAttrs: string, extra = '') =>
  `<p:sp><p:nvSpPr><p:nvPr><p:ph ${phAttrs}/></p:nvPr></p:nvSpPr><p:spPr>${extra}</p:spPr></p:sp>`;

describe('reading placeholders', () => {
  it('defaults type to obj and idx to 0, as the schema specifies', () => {
    const [ph] = readPlaceholders(part(shape('')));
    expect(ph).toMatchObject({ type: 'obj', idx: 0 });
  });

  it('reads explicit type and idx', () => {
    const [ph] = readPlaceholders(part(shape('type="title" idx="3"')));
    expect(ph).toMatchObject({ type: 'title', idx: 3 });
  });

  it('finds the owning shape rather than assuming a nesting depth', () => {
    const [ph] = readPlaceholders(part(shape('type="body"')));
    expect(ph.shape.localName).toBe('sp');
  });
});

describe('slide to layout — matches on idx, never on type', () => {
  const layout = readPlaceholders(part(
    shape('type="title" idx="0"') + shape('type="body" idx="1"') + shape('type="body" idx="2"')
  ));

  it('matches by idx', () => {
    const slide = readPlaceholders(part(shape('type="body" idx="2"')))[0];
    expect(matchSlideToLayout(slide, layout).layoutPlaceholder?.idx).toBe(2);
  });

  it('does NOT fall back to matching on type', () => {
    // Two layout placeholders share type="body". Matching on type would pick one of
    // them; the correct answer is that idx=7 corresponds to nothing.
    const slide = readPlaceholders(part(shape('type="body" idx="7"')))[0];
    const match = matchSlideToLayout(slide, layout);
    expect(match.layoutPlaceholder).toBeNull();
    expect(match.problems.join(' ')).toContain('inherits no position, size or text style');
  });

  it('treats an omitted idx as a real index of 0, not as a wildcard', () => {
    const slide = readPlaceholders(part(shape('type="title"')))[0];
    expect(matchSlideToLayout(slide, layout).layoutPlaceholder?.idx).toBe(0);
  });

  it('honours the no-correspondence sentinel', () => {
    const slide = readPlaceholders(part(shape(`idx="${NO_CORRESPONDENCE_IDX}"`)))[0];
    const match = matchSlideToLayout(slide, layout);
    expect(match.layoutPlaceholder).toBeNull();
    expect(match.problems).toEqual([]);
    expect(match.trace.join(' ')).toContain('by design');
  });

  it('flags a layout with duplicate idx values', () => {
    const dupes = readPlaceholders(part(shape('idx="1"') + shape('idx="1"')));
    const slide = readPlaceholders(part(shape('idx="1"')))[0];
    expect(matchSlideToLayout(slide, dupes).problems.join(' ')).toContain('must be unique');
  });
});

describe('notes slide to notes master — a different rule', () => {
  it('matches on type, not idx', () => {
    // Deliberately mismatched idx values; the match must still succeed.
    const master = readPlaceholders(part(shape('type="body" idx="9"')));
    const notes = readPlaceholders(part(shape('type="body" idx="2"')))[0];
    expect(matchNotesToMaster(notes, master).layoutPlaceholder?.idx).toBe(9);
    expect(matchNotesToMaster(notes, master).trace.join(' ')).toContain('not idx');
  });

  it('reports a missing type on the notes master', () => {
    const master = readPlaceholders(part(shape('type="body"')));
    const notes = readPlaceholders(part(shape('type="sldImg"')))[0];
    expect(matchNotesToMaster(notes, master).problems.join(' ')).toContain('no placeholder of type "sldImg"');
  });
});

describe('layout to master — undocumented, and labelled as such', () => {
  const master = readPlaceholders(part(shape('type="title"') + shape('type="body"')));

  it('flags every match as observed practice rather than specification', () => {
    const layout = readPlaceholders(part(shape('type="title"')))[0];
    const match = matchLayoutToMaster(layout, master);
    expect(match.layoutPlaceholder).not.toBeNull();
    expect(match.problems.join(' ')).toContain('undocumented');
  });

  it('folds ctrTitle onto title, as the ecosystem does', () => {
    const layout = readPlaceholders(part(shape('type="ctrTitle"')))[0];
    expect(matchLayoutToMaster(layout, master).layoutPlaceholder?.type).toBe('title');
  });

  it('folds subTitle onto body', () => {
    const layout = readPlaceholders(part(shape('type="subTitle"')))[0];
    expect(matchLayoutToMaster(layout, master).layoutPlaceholder?.type).toBe('body');
  });

  it('reports no ancestor for types a master cannot carry', () => {
    // A master may only hold title, body, dt, ftr and sldNum, so most layout
    // placeholders have no shape-level ancestor at all.
    const layout = readPlaceholders(part(shape('type="pic"')))[0];
    const match = matchLayoutToMaster(layout, master);
    expect(match.layoutPlaceholder).toBeNull();
    expect(match.trace.join(' ')).toContain('cannot exist on a master');
  });
});

describe('transforms — absent means inherit, not zero', () => {
  it('reports inheritance when there is no xfrm', () => {
    const t = readTransform(frag('<p:spPr/>'));
    expect(t.inherits).toBe(true);
    expect(t.offset).toBeNull();
  });

  it('does not treat an explicit zero offset as inheritance', () => {
    // The distinction that matters: a generator writing <a:off x="0" y="0"/> as a
    // "default" has pinned the shape to the corner and severed inheritance.
    const t = readTransform(frag('<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="50"/></a:xfrm></p:spPr>'));
    expect(t.inherits).toBe(false);
    expect(t.offset).toEqual({ x: 0, y: 0 });
  });

  it('reads flips and rotation', () => {
    const t = readTransform(frag('<p:spPr><a:xfrm rot="5400000" flipH="1"><a:off x="1" y="2"/></a:xfrm></p:spPr>'));
    expect(t.flipH).toBe(true);
    expect(t.flipV).toBe(false);
    expect(t.rotation).toBe(5400000);
  });

  it('converts rotation using 1/60000 of a degree, not 1/64000', () => {
    // ECMA-376's prose for xfrm says 1/64000; its own ST_Angle type, the schema and
    // Office all say 1/60000. The difference reads as a rendering artefact.
    expect(ROTATION_UNITS_PER_DEGREE).toBe(60000);
    expect(rotationDegrees(5400000)).toBe(90);
    expect(rotationDegrees(null)).toBeNull();
  });
});

describe('group transform', () => {
  const group = readTransform(frag(`<p:grpSpPr><a:xfrm>
    <a:off x="1000" y="2000"/><a:ext cx="4000" cy="2000"/>
    <a:chOff x="0" y="0"/><a:chExt cx="2000" cy="1000"/>
  </a:xfrm></p:grpSpPr>`));

  it('scales children by the ratio of extent to child extent', () => {
    const r = applyGroupTransform(group, { offset: { x: 500, y: 500 }, extent: { cx: 200, cy: 100 } });
    expect(r.scale).toEqual({ x: 2, y: 2 });
    expect(r.extent).toEqual({ cx: 400, cy: 200 });
  });

  it('translates out of child space into the parent', () => {
    const r = applyGroupTransform(group, { offset: { x: 500, y: 500 }, extent: null });
    // 1000 + (500 - 0) * 2
    expect(r.offset).toEqual({ x: 2000, y: 3000 });
  });

  it('subtracts the child offset before scaling', () => {
    const shifted = readTransform(frag(`<p:grpSpPr><a:xfrm>
      <a:off x="0" y="0"/><a:ext cx="100" cy="100"/>
      <a:chOff x="50" y="50"/><a:chExt cx="100" cy="100"/>
    </a:xfrm></p:grpSpPr>`));
    expect(applyGroupTransform(shifted, { offset: { x: 50, y: 50 }, extent: null }).offset).toEqual({ x: 0, y: 0 });
  });

  it('disables scaling when the child extent is zero rather than dividing by it', () => {
    const zero = readTransform(frag(`<p:grpSpPr><a:xfrm>
      <a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chExt cx="0" cy="0"/>
    </a:xfrm></p:grpSpPr>`));
    expect(applyGroupTransform(zero, { offset: { x: 7, y: 7 }, extent: { cx: 3, cy: 3 } }).scale)
      .toEqual({ x: 1, y: 1 });
  });

  it('is the identity when extent equals child extent', () => {
    const identity = readTransform(frag(`<p:grpSpPr><a:xfrm>
      <a:off x="10" y="10"/><a:ext cx="100" cy="100"/>
      <a:chOff x="10" y="10"/><a:chExt cx="100" cy="100"/>
    </a:xfrm></p:grpSpPr>`));
    const r = applyGroupTransform(identity, { offset: { x: 40, y: 40 }, extent: { cx: 5, cy: 5 } });
    expect(r.offset).toEqual({ x: 40, y: 40 });
    expect(r.extent).toEqual({ cx: 5, cy: 5 });
  });
});

describe('colours — three vocabularies that are not the same', () => {
  it('keeps the slot names and the map keys distinct', () => {
    // They overlap on the accents, which is exactly why the bug hides.
    expect(COLOUR_SLOTS).toContain('dk1');
    expect(COLOUR_MAP_KEYS).not.toContain('dk1');
    expect(COLOUR_MAP_KEYS).toContain('bg1');
    expect(COLOUR_SLOTS).not.toContain('bg1');
    for (const accent of ['accent1', 'accent6', 'hlink', 'folHlink']) {
      expect(COLOUR_SLOTS).toContain(accent);
      expect(COLOUR_MAP_KEYS).toContain(accent);
    }
  });

  const clrScheme = frag(`<a:clrScheme name="Office">
    <a:dk1><a:srgbClr val="000000"/></a:dk1>
    <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
    <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
  </a:clrScheme>`);

  it('resolves tx1 through the map, NOT to dk1 by assumption', () => {
    // The specification's own example inverts these on a dark master.
    const map = readColourMap(frag('<p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'));
    const r = resolveSchemeColour('tx1', map, clrScheme);
    expect(r.slot).toBe('lt1');
    expect(r.value?.localName).toBe('lt1');
  });

  it('bypasses the map for dk1/lt1/dk2/lt2', () => {
    for (const val of MAP_BYPASSING_VALUES) {
      const r = resolveSchemeColour(val, new Map(), clrScheme);
      expect(r.slot).toBe(val);
    }
    expect(resolveSchemeColour('dk1', new Map(), clrScheme).trace.join(' ')).toContain('bypassing');
  });

  it('reports phClr as supplied by the referencing style', () => {
    expect(resolveSchemeColour('phClr', new Map(), clrScheme).trace.join(' ')).toContain('style reference');
  });

  it('reports a map key with no entry rather than guessing', () => {
    expect(resolveSchemeColour('bg2', new Map(), clrScheme).problems.join(' ')).toContain('no entry in the active map');
  });

  it('reports a slot missing from the theme', () => {
    const map = readColourMap(frag('<p:clrMap bg1="dk2" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'));
    expect(resolveSchemeColour('bg1', map, clrScheme).problems.join(' ')).toContain('no "dk2" entry');
  });
});

describe('colour map override', () => {
  const master = parse(`<?xml version="1.0"?><p:sldMaster ${NS}><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>`);

  it('uses the master map when the slide defers to it', () => {
    const slide = parse(`<?xml version="1.0"?><p:sld ${NS}><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
    const r = resolveColourMap(master, slide);
    expect(r.source).toBe('master');
    expect(r.map.get('tx1')).toBe('dk1');
  });

  it('uses a complete override when the slide supplies one', () => {
    const slide = parse(`<?xml version="1.0"?><p:sld ${NS}><p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr></p:sld>`);
    const r = resolveColourMap(master, slide);
    expect(r.source).toBe('override');
    expect(r.map.get('tx1')).toBe('lt1');
  });
});

describe('style reference indexing is not uniform', () => {
  it('sends fillRef 1001 to the FIRST background fill, 1-based', () => {
    expect(resolveStyleReference('fillRef', 1001)).toMatchObject({ list: 'bgFillStyleLst', position: 1 });
    expect(resolveStyleReference('bgRef', 1002)).toMatchObject({ list: 'bgFillStyleLst', position: 2 });
  });

  it('sends fillRef 1-999 to the fill style list', () => {
    expect(resolveStyleReference('fillRef', 2)).toMatchObject({ list: 'fillStyleLst', position: 2 });
  });

  it('treats 0 and 1000 as no fill', () => {
    expect(resolveStyleReference('fillRef', 0).list).toBeNull();
    expect(resolveStyleReference('bgRef', 1000).list).toBeNull();
  });

  it('does NOT apply the 1001 offset to lnRef or effectRef', () => {
    // Each addresses a single list, so there is no discriminator to encode.
    expect(resolveStyleReference('lnRef', 1001)).toMatchObject({ list: 'lnStyleLst', position: 1001 });
    expect(resolveStyleReference('effectRef', 1001)).toMatchObject({ list: 'effectStyleLst', position: 1001 });
  });

  it('sends lnRef to lnStyleLst despite the specification naming fillStyleLst', () => {
    const r = resolveStyleReference('lnRef', 2);
    expect(r.list).toBe('lnStyleLst');
    expect(r.note).toContain('ECMA-376 names fillStyleLst here');
  });
});
