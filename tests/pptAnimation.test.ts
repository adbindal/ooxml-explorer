import { describe, it, expect } from 'vitest';
import {
  readAnimations,
  deadAnimations,
  animationFindings,
  computeAnimationEvidenceForMarkup,
  ANIMATION_HOST_PART
} from '../services/pptAnimation';

const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const P14 = 'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"';

const parse = (xml: string): Document => new DOMParser().parseFromString(xml, 'application/xml');

const SLIDE = 'ppt/slides/slide1.xml';

/** A text shape. `paragraphs` is how many `a:p` its `p:txBody` holds — what `p:pRg` indexes. */
const sp = (id: string, name: string, paragraphs = 3) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>` +
  `<p:txBody><a:bodyPr/><a:lstStyle/>${'<a:p><a:r><a:t>line</a:t></a:r></a:p>'.repeat(paragraphs)}</p:txBody></p:sp>`;

/** A shape with no text body at all — a rectangle, a line. `p:pRg` has nothing to index. */
const emptySp = (id: string, name: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>`;

/** A table frame. It has text, but not at an index `p:pRg` counts — so it is never judged. */
const graphicFrame = (id: string, name: string) =>
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${name}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
  `<a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:p/><a:p/></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;

/**
 * A slide. The `p:cNvPr id="1"` on the tree's own `p:nvGrpSpPr` is what PowerPoint really
 * writes, so it is here rather than idealised away.
 */
const slide = (shapes: string, timing = '', extraCSld = '') =>
  `<?xml version="1.0"?><p:sld ${P} ${A} ${P14}><p:cSld><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
  `${shapes}</p:spTree>${extraCSld}</p:cSld>${timing}</p:sld>`;

const timing = (nodes: string, bldLst = '') =>
  `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">` +
  `<p:childTnLst>${nodes}</p:childTnLst></p:cTn></p:par></p:tnLst>${bldLst}</p:timing>`;

/** One behaviour node, targeting through `p:cBhvr/p:tgtEl` — not through `p:cTn`. */
const behaviour = (name: string, spid: string, inner = '', id = '9') =>
  `<p:${name}><p:cBhvr><p:cTn id="${id}" dur="500"/>` +
  `<p:tgtEl><p:spTgt spid="${spid}">${inner}</p:spTgt></p:tgtEl></p:cBhvr></p:${name}>`;

/** A `p:cond` in whichever of the four condition lists is named. There is no `p:condLst`. */
const cond = (list: string, spid: string, evt = 'onClick') =>
  `<p:seq><p:cTn id="3"><p:${list}><p:cond evt="${evt}" delay="0">` +
  `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cond></p:${list}></p:cTn></p:seq>`;

const bld = (name: string, spid: string) => `<p:${name} spid="${spid}" grpId="0"/>`;

/** Three text shapes, three entrance effects, one per shape. Everything resolves. */
const healthySlide = () =>
  slide(
    sp('2', 'Title') + sp('3', 'Bullets') + sp('4', 'Footnote'),
    timing(behaviour('animEffect', '2') + behaviour('animEffect', '3') + behaviour('animEffect', '4'))
  );

const index = (xml: string) => readAnimations(parse(xml), SLIDE);
const codes = (xml: string) => index(xml).problems.map(f => f.code);

describe('the animation that points at a shape that is gone', () => {
  it('reports nothing when every behaviour names a shape that is there', () => {
    const found = index(healthySlide());

    expect(found.animations).toHaveLength(3);
    expect(found.animations.map(a => a.target.exists)).toEqual([true, true, true]);
    expect(found.animations.map(a => a.timeNodeId)).toEqual(['9', '9', '9']);
    expect(deadAnimations(found)).toEqual([]);
    expect(found.problems).toEqual([]);
  });

  it('catches the behaviour whose shape was deleted, and leaves its siblings alone', () => {
    // The headline case. Deleting "Bullets" leaves the animation behind: PowerPoint opens
    // the deck without complaint and the effect simply never happens.
    const found = index(
      slide(
        sp('2', 'Title') + sp('4', 'Footnote'),
        timing(behaviour('animEffect', '2') + behaviour('animEffect', '3') + behaviour('animEffect', '4'))
      )
    );

    expect(deadAnimations(found)).toHaveLength(1);
    expect(found.animations.map(a => a.target.exists)).toEqual([true, false, true]);

    const [problem] = found.problems;
    expect(problem.code).toBe('animation/dead-target');
    expect(problem.severity).toBe('error');
    // Nothing about a timing tree is drawn, so no render, thumbnail or export shows this.
    expect(problem.silent).toBe(true);
    expect(problem.part).toBe(SLIDE);
    expect(problem.subject).toEqual({ spid: '3', node: 'p:animEffect' });
    expect(problem.message).toContain('shape id 3');
  });

  it('scopes shapes to the shape tree, so a shape elsewhere in the part is not a target', () => {
    // A p:cNvPr outside p:spTree is not a shape an animation can reach. Counting it would
    // silence a genuinely dead animation, and would invent duplicate ids besides.
    //
    // The fixture puts a whole p:sp outside the tree rather than a bare p:cNvPr: a loose
    // p:cNvPr is discarded anyway by the walk out to its owning shape, so a fixture built
    // that way passes whether the scoping happens or not.
    const found = index(
      slide(sp('2', 'Title'), timing(behaviour('anim', '42')), sp('42', 'Outside the tree'))
    );

    expect(found.shapes.has('42')).toBe(false);
    expect(found.problems.map(f => f.code)).toEqual(['animation/dead-target']);
  });

  it('reads the shapes it did find, with their names and paragraph counts', () => {
    const found = index(healthySlide());

    expect([...found.shapes.keys()]).toEqual(['1', '2', '3', '4']);
    expect(found.shapes.get('3')?.name).toBe('Bullets');
    expect(found.shapes.get('3')?.paragraphCount).toBe(3);
    // Not a p:sp: "how many paragraphs" has no answer here rather than the answer zero.
    expect(index(slide(graphicFrame('7', 'Table'), timing(''))).shapes.get('7')?.paragraphCount).toBeNull();
  });
});

describe('where a behaviour actually names its target', () => {
  it('reads p:cBhvr/p:tgtEl and not p:cTn/p:tgtEl', () => {
    // CT_TLCommonTimeNodeData has exactly six children and p:tgtEl is not among them, so
    // the commonly written p:cTn/p:tgtEl/p:spTgt path does not exist. A module that read
    // it would find nothing on every real deck — and would report this fixture, which is
    // the invalid arrangement, as a dead target.
    const found = index(
      slide(
        sp('2', 'Title'),
        timing('<p:anim><p:cBhvr><p:cTn id="9"><p:tgtEl><p:spTgt spid="99"/></p:tgtEl></p:cTn></p:cBhvr></p:anim>')
      )
    );

    expect(found.animations).toHaveLength(1);
    expect(found.animations[0].target.kind).toBe('none');
    expect(found.problems).toEqual([]);
  });

  it('reads all eight behaviour elements, each of which wraps a p:cBhvr', () => {
    const names = ['anim', 'animClr', 'animEffect', 'animMotion', 'animRot', 'animScale', 'set', 'cmd'];
    const found = index(slide(sp('2', 'Title'), timing(names.map(n => behaviour(n, '99')).join(''))));

    expect(found.animations.map(a => a.label)).toEqual(names.map(n => `p:${n}`));
    expect(found.problems.filter(f => f.code === 'animation/dead-target')).toHaveLength(8);
  });

  it('ignores a behaviour that is not inside p:timing', () => {
    // Only the timing tree animates. An element called p:anim anywhere else is not one.
    expect(index(slide(sp('2', 'Title') + behaviour('anim', '99'))).animations).toEqual([]);
  });

  it('ignores an element with the right name in the wrong namespace', () => {
    // conformance.ts maps Strict to Transitional before any analyzer runs, so exact
    // equality is the whole comparison and a foreign namespace is simply not ours.
    const found = index(
      slide(
        sp('2', 'Title'),
        timing('<q:anim xmlns:q="urn:example:other"><p:cBhvr><p:tgtEl><p:spTgt spid="99"/></p:tgtEl></p:cBhvr></q:anim>')
      )
    );

    expect(found.animations).toEqual([]);
    expect(found.problems).toEqual([]);
  });

  it('leaves a slide, sound or ink target unjudged', () => {
    const tgt = (inner: string) => `<p:anim><p:cBhvr><p:cTn id="9"/><p:tgtEl>${inner}</p:tgtEl></p:cBhvr></p:anim>`;
    const found = index(
      slide(
        sp('2', 'Title'),
        timing(tgt('<p:sldTgt/>') + tgt('<p:sndTgt><a:snd r:embed="rId9" xmlns:r="urn:r"/></p:sndTgt>') + tgt('<p:inkTgt spid="x"/>'))
      )
    );

    expect(found.animations.map(a => a.target.kind)).toEqual(['slide', 'sound', 'ink']);
    expect(found.animations.every(a => a.target.exists === null)).toBe(true);
    expect(found.problems).toEqual([]);
  });
});

describe('conditions — the click that waits for a shape that is not there', () => {
  it('finds p:cond in every one of the four condition lists', () => {
    // There is no p:condLst element: conditions live in p:stCondLst and p:endCondLst on
    // p:cTn, and p:prevCondLst / p:nextCondLst on p:seq. Finding p:cond by name at any
    // depth is what covers all four without relying on an invented parent name.
    for (const list of ['stCondLst', 'endCondLst', 'prevCondLst', 'nextCondLst']) {
      const found = index(slide(sp('2', 'Title'), timing(cond(list, '99'))));

      expect(found.triggers).toHaveLength(1);
      expect(found.triggers[0].list).toBe(`p:${list}`);
      expect(found.triggers[0].event).toBe('onClick');
      expect(found.triggers[0].problems.map(f => f.code)).toEqual(['animation/dead-trigger']);
    }
  });

  it('marks a dead trigger silent, and says the click does nothing', () => {
    const [problem] = index(slide(sp('2', 'Title'), timing(cond('stCondLst', '99')))).problems;

    expect(problem.code).toBe('animation/dead-trigger');
    expect(problem.severity).toBe('error');
    expect(problem.silent).toBe(true);
    expect(problem.subject).toEqual({ spid: '99', evt: 'onClick' });
    expect(problem.message).toContain('p:stCondLst');
  });

  it('reports nothing when the triggering shape is still on the slide', () => {
    expect(codes(slide(sp('2', 'Title'), timing(cond('stCondLst', '2'))))).toEqual([]);
  });

  it('reads p:endSync, which is the same type as p:cond', () => {
    const found = index(
      slide(
        sp('2', 'Title'),
        timing('<p:excl><p:cTn id="3"><p:endSync evt="end"><p:tgtEl><p:spTgt spid="99"/></p:tgtEl></p:endSync></p:cTn></p:excl>')
      )
    );

    expect(found.triggers.map(t => t.label)).toEqual(['p:endSync']);
    expect(found.triggers[0].problems.map(f => f.code)).toEqual(['animation/dead-trigger']);
  });

  it('leaves a condition that waits on a time node rather than a shape alone', () => {
    // CT_TLTimeCondition is a choice: p:tgtEl, p:tn, or p:rtn. Only the first names a shape.
    const found = index(
      slide(
        sp('2', 'Title'),
        timing('<p:seq><p:cTn id="3"><p:stCondLst><p:cond delay="0"><p:tn val="5"/></p:cond></p:stCondLst></p:cTn></p:seq>')
      )
    );

    expect(found.triggers).toHaveLength(1);
    expect(found.triggers[0].event).toBeNull();
    expect(found.triggers[0].target.kind).toBe('none');
    expect(found.problems).toEqual([]);
  });
});

describe('build entries — p:bldLst', () => {
  it('reads all four build elements and reports the ones whose shape is gone', () => {
    const names = ['bldP', 'bldDgm', 'bldOleChart', 'bldGraphic'];
    const found = index(slide(sp('2', 'Title'), timing('', `<p:bldLst>${names.map(n => bld(n, '99')).join('')}</p:bldLst>`)));

    expect(found.builds.map(b => b.label)).toEqual(names.map(n => `p:${n}`));
    expect(found.problems.map(f => f.code)).toEqual(names.map(() => 'animation/dead-build'));
    expect(found.problems[0].severity).toBe('error');
    expect(found.problems[0].silent).toBe(true);
    expect(found.problems[0].subject).toEqual({ spid: '99', node: 'p:bldP' });
  });

  it('reports nothing for a build entry whose shape is still there', () => {
    const found = index(slide(sp('2', 'Title'), timing('', `<p:bldLst>${bld('bldP', '2')}</p:bldLst>`)));

    expect(found.builds[0].exists).toBe(true);
    expect(found.problems).toEqual([]);
  });
});

describe('paragraph ranges — the text the shape no longer has', () => {
  const withRange = (attrs: string, paragraphs = 3, id = '2') =>
    slide(sp(id, 'Bullets', paragraphs), timing(behaviour('animEffect', id, `<p:txEl><p:pRg ${attrs}/></p:txEl>`)));

  it('accepts a range that ends on the last paragraph, because indices are zero-based and inclusive', () => {
    // st="0" end="2" is the spec's own example for "the first 3 text paragraphs". Under a
    // one-based reading this fixture would be over-reported.
    const found = index(withRange('st="0" end="2"', 3));

    expect(found.animations[0].target.paragraphRange).toEqual({ start: 0, end: 2 });
    expect(found.problems).toEqual([]);
  });

  it('reports the range that runs one paragraph past the end', () => {
    const [problem] = index(withRange('st="0" end="3"', 3)).problems;

    expect(problem.code).toBe('animation/paragraph-out-of-range');
    expect(problem.severity).toBe('error');
    expect(problem.silent).toBe(true);
    expect(problem.subject).toEqual({ spid: '2', st: '0', end: '3', paragraphs: '3' });
    expect(problem.message).toContain('numbered 0 to 2');
  });

  it('says a shape with no text body at all has no text, rather than reporting nothing', () => {
    const found = index(
      slide(emptySp('2', 'Rectangle'), timing(behaviour('animEffect', '2', '<p:txEl><p:pRg st="0" end="0"/></p:txEl>')))
    );

    expect(found.shapes.get('2')?.paragraphCount).toBe(0);
    expect(found.problems[0].code).toBe('animation/paragraph-out-of-range');
    expect(found.problems[0].message).toContain('no text at all');
  });

  it('calls an inverted range a warning, not an error, and does not also range-check it', () => {
    // ECMA does not define what a consumer does with st > end, so the effect is undefined
    // rather than provably absent.
    const found = index(withRange('st="2" end="1"', 3));

    expect(found.problems.map(f => f.code)).toEqual(['animation/inverted-range']);
    expect(found.problems[0].severity).toBe('warning');
  });

  it('does not range-check a target whose shape is missing — that is one defect, not two', () => {
    const found = index(
      slide(sp('2', 'Title'), timing(behaviour('animEffect', '99', '<p:txEl><p:pRg st="0" end="9"/></p:txEl>')))
    );

    expect(found.problems.map(f => f.code)).toEqual(['animation/dead-target']);
  });

  it('does not report an inverted range on a shape that is not there either', () => {
    // The inverted-range check needs no shape to run against, so it is the one range
    // problem that can be reported about a shape the slide does not have. The dead target
    // is the finding worth reading; a second one about the range it does not have is noise.
    const found = index(
      slide(sp('2', 'Title'), timing(behaviour('animEffect', '99', '<p:txEl><p:pRg st="2" end="1"/></p:txEl>')))
    );

    expect(found.problems.map(f => f.code)).toEqual(['animation/dead-target']);
  });

  it('does not range-check a table or chart frame, whose text p:pRg does not index', () => {
    const found = index(
      slide(graphicFrame('7', 'Table'), timing(behaviour('animEffect', '7', '<p:txEl><p:pRg st="0" end="9"/></p:txEl>')))
    );

    expect(found.animations[0].target.exists).toBe(true);
    expect(found.problems).toEqual([]);
  });

  it('leaves a range with a missing or non-numeric index unjudged rather than reading it as 0', () => {
    // @st and @end are both REQUIRED unsignedInt, so this is a schema-validity problem and
    // not an animation problem. Reading the absent @st with Number() would make it 0 and
    // produce a confident claim about a range the markup never states.
    for (const attrs of ['end="4"', 'st="1"', 'st="x" end="2"', 'st="1.5" end="2"']) {
      const found = index(withRange(attrs, 3));

      expect(found.animations[0].target.paragraphRange).toBeNull();
      expect(found.problems).toEqual([]);
    }
  });

  it('range-checks the target of a condition as well as that of a behaviour', () => {
    const found = index(
      slide(
        sp('2', 'Bullets', 2),
        timing(
          '<p:seq><p:cTn id="3"><p:stCondLst><p:cond evt="onClick">' +
            '<p:tgtEl><p:spTgt spid="2"><p:txEl><p:pRg st="0" end="5"/></p:txEl></p:spTgt></p:tgtEl>' +
            '</p:cond></p:stCondLst></p:cTn></p:seq>'
        )
      )
    );

    expect(found.triggers[0].problems.map(f => f.code)).toEqual(['animation/paragraph-out-of-range']);
  });

  it('records a character range without judging it', () => {
    const found = index(
      slide(sp('2', 'Bullets'), timing(behaviour('animEffect', '2', '<p:txEl><p:charRg st="0" end="99"/></p:txEl>')))
    );

    expect(found.animations[0].target.characterRange).toBe(true);
    expect(found.animations[0].target.paragraphRange).toBeNull();
    expect(found.problems).toEqual([]);
  });
});

describe('@spid is not always a shape id', () => {
  // The SDK records a StringValidator for Office2007 and a numeric ST_DrawingElementId
  // from Office2010 on, while p:cNvPr/@id is UInt32 in every version. A 2007-era deck can
  // legally write spid="_x0000_s1026" — a VML shape name matching no p:cNvPr/@id BY
  // DESIGN. Reporting those would be a confident wrong answer on every deck of that age.
  it('declines to judge the string form on a behaviour', () => {
    const found = index(slide(sp('2', 'Title'), timing(behaviour('anim', '_x0000_s1026'))));

    expect(found.animations[0].target.shapeId).toBe('_x0000_s1026');
    expect(found.animations[0].target.exists).toBeNull();
    expect(found.problems).toEqual([]);
    // And it is not counted among the dead either: "1 of 1 animations will never run" is
    // the same wrong answer as the finding, in the sentence people actually read.
    expect(deadAnimations(found)).toEqual([]);
  });

  it('declines to judge the string form on a condition and on a build entry', () => {
    const found = index(
      slide(
        sp('2', 'Title'),
        timing(cond('stCondLst', '_x0000_s1027'), `<p:bldLst>${bld('bldP', '_x0000_s1028')}</p:bldLst>`)
      )
    );

    expect(found.triggers[0].target.exists).toBeNull();
    expect(found.builds[0].exists).toBeNull();
    expect(found.problems).toEqual([]);
  });
});

describe('two shapes with one id', () => {
  const duplicated = (targeted: string) =>
    slide(sp('3', 'Bullets') + sp('3', 'Bullets copy') + sp('4', 'Footnote'), timing(behaviour('animEffect', targeted)));

  it('reports an ambiguous id as a warning, because the animation still runs — on one of them', () => {
    const found = index(duplicated('3'));

    expect(found.duplicateShapeIds.get('3')).toBe(2);
    const [problem] = found.problems;
    expect(problem.code).toBe('animation/duplicate-shape-id');
    expect(problem.severity).toBe('warning');
    expect(problem.silent).toBe(true);
    expect(problem.subject).toEqual({ spid: '3', shapes: '2' });
    expect(problem.message).toContain('2 shapes');
  });

  it('says nothing when the duplicated id is not what the timing tree targets', () => {
    // A duplicate id nothing animates is a different and much quieter problem.
    const found = index(duplicated('4'));

    expect(found.duplicateShapeIds.get('3')).toBe(2);
    expect(found.problems).toEqual([]);
  });

  it('keeps the first shape with the id, and reports the duplicate once however many target it', () => {
    const found = index(
      slide(
        sp('3', 'Bullets') + sp('3', 'Bullets copy'),
        timing(behaviour('animEffect', '3') + behaviour('anim', '3') + cond('stCondLst', '3'), `<p:bldLst>${bld('bldP', '3')}</p:bldLst>`)
      )
    );

    expect(found.shapes.get('3')?.name).toBe('Bullets');
    expect(found.problems.filter(f => f.code === 'animation/duplicate-shape-id')).toHaveLength(1);
  });
});

describe('what belongs to pptMedia rather than here', () => {
  it('does not re-report a p:video or p:audio timing node with a dangling target', () => {
    // pptMedia.ts owns these as media/dangling-trigger. Reporting them again would say the
    // same thing twice in two different vocabularies.
    const found = index(
      slide(
        sp('2', 'Title'),
        timing(
          '<p:video><p:cMediaNode><p:cTn id="9"/><p:tgtEl><p:spTgt spid="99"/></p:tgtEl></p:cMediaNode></p:video>' +
            '<p:audio><p:cMediaNode><p:cTn id="10"/><p:tgtEl><p:spTgt spid="98"/></p:tgtEl></p:cMediaNode></p:audio>'
        )
      )
    );

    expect(found.animations).toEqual([]);
    expect(found.problems).toEqual([]);
  });
});

describe('tolerating input', () => {
  it('returns an empty index for a slide with no timing tree', () => {
    const found = index(slide(sp('2', 'Title')));

    expect(found).toEqual({
      part: SLIDE,
      animations: [],
      triggers: [],
      builds: [],
      shapes: new Map(),
      duplicateShapeIds: new Map(),
      problems: []
    });
  });

  it('returns an empty index for a timing tree with nothing in it', () => {
    expect(index(slide(sp('2', 'Title'), '<p:timing><p:tnLst/></p:timing>')).problems).toEqual([]);
  });

  it('reads an element root as well as a document', () => {
    const doc = parse(healthySlide());
    expect(readAnimations(doc.documentElement, SLIDE).animations).toHaveLength(3);
  });

  it('tolerates a behaviour with no p:cBhvr at all', () => {
    const found = index(slide(sp('2', 'Title'), timing('<p:anim/>')));

    expect(found.animations[0].target.kind).toBe('none');
    expect(found.animations[0].timeNodeId).toBeNull();
    expect(found.problems).toEqual([]);
  });

  it('unions every finding for the registry', () => {
    const found = animationFindings(
      parse(
        slide(
          sp('2', 'Title'),
          timing(behaviour('animEffect', '99') + cond('stCondLst', '98'), `<p:bldLst>${bld('bldP', '97')}</p:bldLst>`)
        )
      ),
      SLIDE
    );

    expect(found.map(f => f.code).sort()).toEqual([
      'animation/dead-build',
      'animation/dead-target',
      'animation/dead-trigger'
    ]);
    expect(found.every(f => f.part === SLIDE)).toBe(true);
  });

  it('matches the parts whose schema permits p:timing, and no others', () => {
    expect(ANIMATION_HOST_PART.test('ppt/slides/slide1.xml')).toBe(true);
    expect(ANIMATION_HOST_PART.test('ppt/slideLayouts/slideLayout2.xml')).toBe(true);
    expect(ANIMATION_HOST_PART.test('ppt/slideMasters/slideMaster1.xml')).toBe(true);
    // CT_NotesSlide has no p:timing child, so a notes slide can never hold anything to find.
    expect(ANIMATION_HOST_PART.test('ppt/notesSlides/notesSlide1.xml')).toBe(false);
    expect(ANIMATION_HOST_PART.test('ppt/slides/_rels/slide1.xml.rels')).toBe(false);
    expect(ANIMATION_HOST_PART.test('ppt/presentation.xml')).toBe(false);
  });
});

describe('computeAnimationEvidenceForMarkup — panel wiring', () => {
  const deadDeck = () =>
    slide(
      sp('2', 'Title') + sp('4', 'Footnote'),
      timing(
        behaviour('animEffect', '2') + behaviour('animEffect', '3') + behaviour('animEffect', '5'),
        `<p:bldLst>${bld('bldP', '2')}</p:bldLst>`
      )
    );

  it('returns null when no part in the bundle can host a timing tree', () => {
    expect(computeAnimationEvidenceForMarkup({ 'ppt/presentation.xml': '<p:presentation/>' }, '')).toBeNull();
  });

  it('returns null when the only slide has no timing tree', () => {
    expect(computeAnimationEvidenceForMarkup({ [SLIDE]: slide(sp('2', 'Title')) }, '')).toBeNull();
  });

  it('returns null rather than throwing when the slide does not parse', () => {
    expect(computeAnimationEvidenceForMarkup({ [SLIDE]: '<p:sld><unclosed>' }, '')).toBeNull();
  });

  it('skips past a host part with no animations to one that has them', () => {
    // Key order is insertion order, so taking the first host part blind would report "no
    // animations" for this deck purely because the layout was bundled first.
    const evidence = computeAnimationEvidenceForMarkup(
      { 'ppt/slideLayouts/slideLayout1.xml': slide(sp('2', 'Title')), [SLIDE]: healthySlide() },
      ''
    );

    expect(evidence!.lines[0]).toContain(SLIDE);
    expect(evidence!.lines[0]).toContain('3 animation behaviour(s)');
  });

  it('leads with the count that someone can act on', () => {
    const evidence = computeAnimationEvidenceForMarkup({ [SLIDE]: deadDeck() }, '');

    expect(evidence!.lines[1]).toContain('2 of 3 animations will never run');
    // The ids are matched with their brackets on purpose: "3, 5" is a substring of the
    // list every targeted id would produce, so the loose form passes either way.
    expect(evidence!.lines[1]).toContain('(3, 5)');
  });

  it('says so plainly when every animation resolves', () => {
    const evidence = computeAnimationEvidenceForMarkup({ [SLIDE]: healthySlide() }, '');

    expect(evidence!.lines[1]).toContain('Every animation names a shape id that exists');
    expect(evidence!.lines.some(l => l.includes('will never run'))).toBe(false);
  });

  it('counts dead conditions and dead build entries separately from behaviours', () => {
    const evidence = computeAnimationEvidenceForMarkup(
      {
        [SLIDE]: slide(
          sp('2', 'Title'),
          timing(cond('stCondLst', '99'), `<p:bldLst>${bld('bldP', '98')}${bld('bldDgm', '97')}</p:bldLst>`)
        )
      },
      ''
    );

    expect(evidence!.lines.some(l => l.includes('1 timing condition(s) wait on a shape'))).toBe(true);
    expect(evidence!.lines.some(l => l.includes('2 build entry(ies) describe how to reveal'))).toBe(true);
  });

  it('renders every finding into the lines', () => {
    const evidence = computeAnimationEvidenceForMarkup({ [SLIDE]: deadDeck() }, '');

    expect(evidence!.lines.filter(l => l.includes('no shape in ppt/slides/slide1.xml has that id'))).toHaveLength(2);
    expect(evidence!.lines.some(l => l.includes('renders correctly and is broken anyway'))).toBe(true);
  });

  it('describes the shape the selected markup names', () => {
    const evidence = computeAnimationEvidenceForMarkup(
      { [SLIDE]: healthySlide() },
      '<p:tgtEl><p:spTgt spid="3"/></p:tgtEl>'
    );

    expect(evidence!.lines.some(l => l.includes('"Bullets"') && l.includes('3 paragraph(s)'))).toBe(true);
  });

  it('says when the selected markup names a shape that is not there', () => {
    const evidence = computeAnimationEvidenceForMarkup(
      { [SLIDE]: deadDeck() },
      '<p:tgtEl><p:spTgt spid="3"/></p:tgtEl>'
    );

    expect(evidence!.lines.some(l => l.includes('shape id 3, which no shape in ppt/slides/slide1.xml carries'))).toBe(
      true
    );
  });

  it('says nothing about a selected shape when the markup names none', () => {
    const evidence = computeAnimationEvidenceForMarkup({ [SLIDE]: healthySlide() }, '<p:sp><p:spPr/></p:sp>');

    expect(evidence!.lines.some(l => l.includes('The selected markup targets'))).toBe(false);
  });

  it('caps the claim: a resolving animation is not a correct one', () => {
    const evidence = computeAnimationEvidenceForMarkup({ [SLIDE]: healthySlide() }, '');

    expect(evidence!.unresolved.some(u => u.includes('only whether the shape it names still exists'))).toBe(true);
  });

  it('states the Office 2007 @spid limit rather than silently skipping it', () => {
    const evidence = computeAnimationEvidenceForMarkup(
      { [SLIDE]: slide(sp('2', 'Title'), timing(behaviour('anim', '_x0000_s1026'))) },
      ''
    );

    expect(evidence!.unresolved.some(u => u.includes('Office 2007 string form'))).toBe(true);
    expect(computeAnimationEvidenceForMarkup({ [SLIDE]: healthySlide() }, '')!.unresolved).toHaveLength(1);
  });

  it('states the Office 2007 @spid limit when only a build entry uses it', () => {
    const evidence = computeAnimationEvidenceForMarkup(
      { [SLIDE]: slide(sp('2', 'Title'), timing('', `<p:bldLst>${bld('bldP', '_x0000_s1026')}</p:bldLst>`)) },
      ''
    );

    expect(evidence!.unresolved.some(u => u.includes('Office 2007 string form'))).toBe(true);
  });

  it('says a sub-shape target was not resolved', () => {
    const evidence = computeAnimationEvidenceForMarkup(
      { [SLIDE]: slide(sp('2', 'Group'), timing(behaviour('anim', '2', '<p:subSp spid="5"/>'))) },
      ''
    );

    expect(evidence!.unresolved.some(u => u.includes('p:subSp'))).toBe(true);
  });

  it('says a character range was not checked', () => {
    const evidence = computeAnimationEvidenceForMarkup(
      { [SLIDE]: slide(sp('2', 'Bullets'), timing(behaviour('anim', '2', '<p:txEl><p:charRg st="0" end="99"/></p:txEl>'))) },
      ''
    );

    expect(evidence!.unresolved.some(u => u.includes('p:charRg'))).toBe(true);
  });

  it('says a paragraph range on a table or chart frame was not checked', () => {
    const evidence = computeAnimationEvidenceForMarkup(
      {
        [SLIDE]: slide(graphicFrame('7', 'Table'), timing(behaviour('anim', '7', '<p:txEl><p:pRg st="0" end="9"/></p:txEl>')))
      },
      ''
    );

    expect(evidence!.unresolved.some(u => u.includes('not a p:sp'))).toBe(true);
  });

  it('does not claim a paragraph range was unchecked when the shape simply does not exist', () => {
    const evidence = computeAnimationEvidenceForMarkup(
      { [SLIDE]: slide(sp('2', 'Title'), timing(behaviour('anim', '99', '<p:txEl><p:pRg st="0" end="9"/></p:txEl>'))) },
      ''
    );

    expect(evidence!.unresolved.some(u => u.includes('not a p:sp'))).toBe(false);
  });
});
