import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

/**
 * The corpus, checked against ECMA-376 itself.
 *
 * `public/rag-data.json` comes from the Open XML SDK — Microsoft's account of what
 * Office implements. That is the right source, because a user's question is about a file
 * Office wrote. But it leaves the whole dictionary resting on one vendor's description
 * with nothing to check it against, which for a tool whose entire value proposition is
 * *"this answer is grounded"* is the wrong place to have no second opinion.
 *
 * `tests/spec-facts.json` is that second opinion: ECMA-376's own normative XSDs,
 * distilled by `scripts/extractSpecFacts.ts` (run `pnpm run spec:facts`). Two
 * independent derivations of the same format.
 *
 * WHAT THIS TEST IS FOR.
 *
 * Not to force agreement. The two sources genuinely differ, and **both are right about
 * different questions** — ECMA-376 says what a conformant file may contain, the SDK says
 * what Office actually accepts, and the gap between them is the entire subject of
 * [MS-OI29500]. Erasing that gap would make the corpus less true, not more.
 *
 * It is a **ratchet**. Every divergence is enumerated below with a reason. A change to
 * the ingest script that introduces a new one fails here and has to be justified; a
 * change that fixes one also fails, so the record cannot quietly go stale. That is what
 * makes "compare the changed part against the spec" a property of the build rather than
 * a thing someone remembers to do.
 */

const corpus = JSON.parse(
  readFileSync(join(__dirname, '..', 'public', 'rag-data.json'), 'utf8')
) as {
  namespace: string;
  tag: string;
  children?: string[];
  attributes: { name: string; values?: string[]; required?: true }[];
}[];

const spec = JSON.parse(
  readFileSync(join(__dirname, 'spec-facts.json'), 'utf8')
) as Record<string, {
  attributes: Record<string, { valueSets?: string[][]; required?: true }>;
  childSets?: string[][];
}>;

/** The corpus merges an element declared under several complex types; so must this. */
const corpusByKey = new Map<string, Map<string, { values?: string[]; required?: true }>>();
for (const record of corpus) {
  const key = `${record.namespace}:${record.tag}`;
  const attrs = corpusByKey.get(key) ?? new Map();
  for (const attr of record.attributes) {
    // The corpus keeps the qualified name (`w:val`); the schema declares the local one.
    attrs.set(attr.name.split(':').pop()!, attr);
  }
  corpusByKey.set(key, attrs);
}

const corpusChildren = new Map<string, string[]>(
  corpus.map(record => [`${record.namespace}:${record.tag}`, record.children ?? []])
);

const sorted = (values: readonly string[]): string[] => [...values].sort();

/** Values in either list but not both. */
const difference = (a: readonly string[], b: readonly string[]): string[] =>
  [...a.filter(v => !b.includes(v)), ...b.filter(v => !a.includes(v))];

/**
 * Enumerations where the SDK and ECMA-376 disagree, with the reason.
 *
 * Two directions, and they mean opposite things:
 *
 *   `officeExtra`   — the SDK permits values ECMA-376 does not. These are real: Office
 *                     writes them, files in the wild contain them, and a corpus that
 *                     omitted them would call a working document invalid.
 *   `specOnly`      — ECMA-376 permits values the SDK omits. A conformant file may carry
 *                     them and Office may still not round-trip them.
 *
 * Keyed by `element/@attribute`. The values listed are the difference, not the whole set.
 */
const KNOWN_DIVERGENCES: Record<string, { officeExtra?: string[]; specOnly?: string[] }> = {

  // Page-border art. Office ships six `tribal` styles ECMA-376 never adopted, and omits
  // `earth3` and `custom`, which it does not draw. Seventeen attributes share ST_Border.
  "w:bar/@val": { officeExtra: ["tribal1", "tribal2", "tribal3", "tribal4", "tribal5", "tribal6"], specOnly: ["custom", "earth3"] },
  "w:bdr/@val": { officeExtra: ["tribal1", "tribal2", "tribal3", "tribal4", "tribal5", "tribal6"], specOnly: ["custom", "earth3"] },
  "w:between/@val": { officeExtra: ["tribal1", "tribal2", "tribal3", "tribal4", "tribal5", "tribal6"], specOnly: ["custom", "earth3"] },
  "w:bottom/@val": { officeExtra: ["tribal1", "tribal2", "tribal3", "tribal4", "tribal5", "tribal6"], specOnly: ["custom", "earth3"] },
  "w:end/@val": { officeExtra: ["tribal1", "tribal2", "tribal3", "tribal4", "tribal5", "tribal6"], specOnly: ["custom", "earth3"] },
  "w:insideH/@val": { officeExtra: ["tribal1", "tribal2", "tribal3", "tribal4", "tribal5", "tribal6"], specOnly: ["custom", "earth3"] },
  "w:insideV/@val": { officeExtra: ["tribal1", "tribal2", "tribal3", "tribal4", "tribal5", "tribal6"], specOnly: ["custom", "earth3"] },
  "w:left/@val": { officeExtra: ["tribal1", "tribal2", "tribal3", "tribal4", "tribal5", "tribal6"], specOnly: ["custom", "earth3"] },
  "w:right/@val": { officeExtra: ["tribal1", "tribal2", "tribal3", "tribal4", "tribal5", "tribal6"], specOnly: ["custom", "earth3"] },
  "w:start/@val": { officeExtra: ["tribal1", "tribal2", "tribal3", "tribal4", "tribal5", "tribal6"], specOnly: ["custom", "earth3"] },
  "w:tl2br/@val": { officeExtra: ["tribal1", "tribal2", "tribal3", "tribal4", "tribal5", "tribal6"], specOnly: ["custom", "earth3"] },
  "w:top/@val": { officeExtra: ["tribal1", "tribal2", "tribal3", "tribal4", "tribal5", "tribal6"], specOnly: ["custom", "earth3"] },
  "w:tr2bl/@val": { officeExtra: ["tribal1", "tribal2", "tribal3", "tribal4", "tribal5", "tribal6"], specOnly: ["custom", "earth3"] },

  // Excel's date cell type. ST_CellType in Transitional is b|n|e|s|str|inlineStr — `d` is
  // absent. Excel writes t="d" for ISO dates anyway, so a strict Transitional validator
  // rejects a file Excel itself produced. One of the more consequential gaps here.
  "x:cell/@t": { officeExtra: ["d"] },
  "x:nc/@t": { officeExtra: ["d"] },
  "x:oc/@t": { officeExtra: ["d"] },

  // PowerPoint animation triggers added after the standard: onMediaBookmark is Office 2010
  // media, and `none` is the SDK's explicit no-trigger.
  "p:cTn/@syncBehavior": { officeExtra: ["none"] },
  "p:cond/@evt": { officeExtra: ["none", "onMediaBookmark"] },
  "p:endSync/@evt": { officeExtra: ["none", "onMediaBookmark"] },

  // The SDK's `invalid` sentinel for a value it does not recognise, against the spec's
  // `custom`. Not a format difference — two vocabularies for the same idea.
  "p:modifyVerifier/@cryptAlgorithmClass": { officeExtra: ["invalid"], specOnly: ["custom"] },
  "p:modifyVerifier/@cryptAlgorithmType": { officeExtra: ["invalid"], specOnly: ["custom"] },
  "p:modifyVerifier/@cryptProviderType": { officeExtra: ["invalid"], specOnly: ["custom"] },

  // Values the specification permits and Office does not implement.
  "a:bodyPr/@anchor": { specOnly: ["dist", "just"] },
  "a:tcPr/@anchor": { specOnly: ["dist", "just"] },
  "c:scatterStyle/@val": { specOnly: ["none"] },
  "c:splitType/@val": { specOnly: ["auto"] },
  "p:cTn/@masterRel": { specOnly: ["lastClick"] },
  "w:caption/@pos": { specOnly: ["left", "right"] },
  "w:lvlJc/@val": { specOnly: ["both", "distribute", "end", "highKashida", "lowKashida", "mediumKashida", "numTab", "start", "thaiDistribute"] },
  "w:pos/@val": { specOnly: ["docEnd"] },
  "w:vAlign/@val": { specOnly: ["both"] },
  "x:connection/@credentials": { specOnly: ["prompt"] },
  "x:filters/@calendarType": { specOnly: ["saka"] },
  "x:format/@action": { specOnly: ["drill", "formula"] },
  "x:textPr/@fileType": { specOnly: ["lin", "other"] },

  // Values Office implements that the specification does not list.
  "w:calendar/@val": { officeExtra: ["umalqura"] },
  "w:rFonts/@hint": { officeExtra: ["cs"] },

  // Office Math. ST_YAlign in the shared simple types carries the full set for any
  // vertically aligned object; the SDK models the subset that means something for an
  // equation base, plus `bot` as a legacy spelling of `bottom`.
  "m:baseJc/@val": { officeExtra: ["bot"], specOnly: ["inline", "inside", "outside"] },
  "m:mcJc/@val": { specOnly: ["inside", "outside"] },

  // VML shadows. Word writes shadow types positioned relative to the drawing or the
  // shape; ECMA-376's VML annex lists only the four classical kinds.
  "v:shadow/@type": { officeExtra: ["drawingRelative", "shapeRelative"] },

  // Document property variants. `cf` is the clipboard-format variant, which OLE uses and
  // the standard's variant list omits.
  "vt:vector/@baseType": { officeExtra: ["cf"] }
};

const divergenceKey = (element: string, attribute: string): string => `${element}/@${attribute}`;

interface Divergence {
  key: string;
  officeExtra: string[];
  specOnly: string[];
}

const findDivergences = (): Divergence[] => {
  const out: Divergence[] = [];
  for (const [element, facts] of Object.entries(spec)) {
    const ours = corpusByKey.get(element);
    if (!ours) continue;
    for (const [attribute, specAttr] of Object.entries(facts.attributes)) {
      const mine = ours.get(attribute);
      if (!mine || !specAttr.valueSets?.length || !mine.values) continue;

      const ourValues = sorted(mine.values);
      // An element name is not a type. If the corpus matches ANY value set the schema
      // permits for this name, it has simply kept a different context, which is not a
      // divergence — see SpecAttribute.valueSets.
      if (specAttr.valueSets.some(set => set.join('\u0000') === ourValues.join('\u0000'))) continue;

      // Compare against the closest permitted set, so the report names the smallest
      // real difference rather than an arbitrary one.
      const closest = [...specAttr.valueSets].sort(
        (a, b) => difference(a, ourValues).length - difference(b, ourValues).length
      )[0];
      const officeExtra = ourValues.filter(v => !closest.includes(v));
      const specOnly = closest.filter(v => !ourValues.includes(v));
      if (officeExtra.length || specOnly.length) {
        out.push({ key: divergenceKey(element, attribute), officeExtra, specOnly });
      }
    }
  }
  return out.sort((a, b) => a.key.localeCompare(b.key));
};

describe('the corpus against ECMA-376', () => {
  it('checks a substantial part of the corpus, not a corner of it', () => {
    // A conformance suite that silently matched nothing would pass every other test in
    // this file. This is the guard against that.
    let compared = 0;
    for (const [element, facts] of Object.entries(spec)) {
      const ours = corpusByKey.get(element);
      if (!ours) continue;
      for (const attribute of Object.keys(facts.attributes)) if (ours.has(attribute)) compared += 1;
    }
    expect(compared, 'attributes cross-checked against the specification').toBeGreaterThan(3000);
  });

  it('agrees with the specification on every enumeration except the recorded ones', () => {
    const unexpected = findDivergences().filter(d => {
      const known = KNOWN_DIVERGENCES[d.key];
      if (!known) return true;
      return (
        JSON.stringify(sorted(known.officeExtra ?? [])) !== JSON.stringify(d.officeExtra) ||
        JSON.stringify(sorted(known.specOnly ?? [])) !== JSON.stringify(d.specOnly)
      );
    });

    expect(
      unexpected.map(d => `${d.key} office-only=[${d.officeExtra}] spec-only=[${d.specOnly}]`),
      'new or changed divergence from ECMA-376 — justify it and record it in KNOWN_DIVERGENCES'
    ).toEqual([]);
  });

  it('has no stale entries in the divergence record', () => {
    // A divergence that has been resolved must be removed from the list, or the record
    // slowly becomes a list of things that used to be true.
    const live = new Set(findDivergences().map(d => d.key));
    const stale = Object.keys(KNOWN_DIVERGENCES).filter(key => !live.has(key));

    expect(stale, 'recorded divergences that no longer occur — delete them').toEqual([]);
  });
});


/* ------------------------------------------------------------------------- *
 * Content models
 * ------------------------------------------------------------------------- */

/**
 * ON THE REPRESENTATION, SINCE IT IS THE OBVIOUS QUESTION.
 *
 * `children: string[]` is a flat, unordered set. The schema says considerably more:
 * 836 of 913 content models are an `xsd:sequence`, 401 of them with more than one child,
 * so **order is enforced**; 880 children are required and 303 are repeatable.
 *
 * Storing that faithfully is not a matter of adding fields. A content model belongs to a
 * *type*, not to an element name, and 97 element names are declared with genuinely
 * different models depending on context — `w:jc` is `CT_Jc` in a paragraph and
 * `CT_JcTable` in a table. Carrying order would mean keying records by type and having
 * elements reference them, which is a different corpus shape, not an extra column.
 *
 * That work is not done here, because nothing needs it yet. The failure it would catch —
 * children in the wrong order — makes Word show a repair prompt, which is *visible*, and
 * this engine exists for the faults that are not. When an analyzer genuinely needs
 * ordering, this comment is the argument for changing the shape rather than bolting on a
 * field that cannot hold the answer.
 *
 * REQUIRED AND REPEATABLE WERE TRIED, MEASURED, AND DROPPED.
 *
 * Order is the expensive one, but `required` and `repeatable` looked cheap — only 7
 * element names disagree about them across contexts, against 97 for the full model. So
 * they were built from the SDK's `Particle` trees and measured. The result killed it:
 *
 *   repeatable   marked 6,738 of 6,762 child slots — 100%. Anything reachable through a
 *                repeating group repeats, and almost everything is. A flag that is true
 *                for every row carries no information and costs ~200 KB to say so.
 *   required     marked 140 slots, 2%, where the same count taken from the ECMA XSDs is
 *                433. `w:tr` inside `w:tbl` came out NOT required, though a table must
 *                have a row. Sparse and provably incomplete.
 *
 * Restricting `repeatable` to an element's own bound would reproduce the XSD's 303, but
 * that number describes how the schema was written rather than what a document may
 * contain, which is not the question anyone asks.
 *
 * Neither flag survived contact with its own numbers, so neither is stored. The failure
 * they would catch — a missing mandatory child — is a Word repair prompt, which is
 * visible, and this engine is for the faults that are not.
 *
 * So the check below compares SETS: may this element contain that one. It still catches
 * a whole class of ingest defect, and it caught three while being written.
 */

/** Prefixes whose schemas the corpus ingests. Anything else is out of scope both ways. */
const IN_SCOPE_PREFIXES = new Set([
  'w', 'x', 'p', 'a', 'c', 'dgm', 'cdr', 'pic', 'lc', 'wp', 'xdr', 'v', 'm', 'b', 'ap', 'op', 'vt'
]);

/**
 * Children that differ everywhere they appear, recorded once rather than per element.
 *
 * `w:contentPart` is an Office extension the standard has no equivalent for.
 * `w:smartTag` is the reverse — the standard has it and Office withdrew the feature.
 * `a:schemeClr` inside a theme colour slot would be a scheme colour defined in terms of
 * itself, which the schema permits structurally and Office rejects.
 */
const ALWAYS_CORPUS_ONLY = new Set(['w:contentPart']);
const ALWAYS_SPEC_ONLY = new Set(['w:smartTag', 'a:schemeClr', 'm:oMath', 'm:oMathPara', 'w:customXml']);

/**
 * Wherever the specification reaches equations through `m:oMath`, the SDK lists every
 * Office Math element directly instead. Same set of documents, two ways of writing the
 * model down — an indirection difference, not a disagreement about what may appear.
 */
const inlinesOfficeMath = (specChildren: readonly string[], ours: readonly string[]): boolean =>
  specChildren.includes('m:oMath') || ours.includes('m:oMath');

interface ChildDivergence {
  key: string;
  corpusOnly: string[];
  specOnly: string[];
}

const findChildDivergences = (): ChildDivergence[] => {
  const out: ChildDivergence[] = [];

  for (const [element, facts] of Object.entries(spec)) {
    const record = corpusChildren.get(element);
    if (!facts.childSets?.length || !record) continue;

    const ours = [...record].filter(c => IN_SCOPE_PREFIXES.has(c.split(':')[0])).sort();
    if (facts.childSets.some(set => set.join('\u0000') === ours.join('\u0000'))) continue;

    const closest = [...facts.childSets].sort(
      (a, b) => difference(a, ours).length - difference(b, ours).length
    )[0];

    let corpusOnly = ours.filter(c => !closest.includes(c) && !ALWAYS_CORPUS_ONLY.has(c));
    const specOnly = closest.filter(c => !ours.includes(c) && !ALWAYS_SPEC_ONLY.has(c));
    if (inlinesOfficeMath(closest, ours)) corpusOnly = corpusOnly.filter(c => !c.startsWith('m:'));

    if (corpusOnly.length || specOnly.length) out.push({ key: element, corpusOnly, specOnly });
  }
  return out.sort((a, b) => a.key.localeCompare(b.key));
};

/** Per-element content-model differences that the rules above do not explain. */
const KNOWN_CHILD_DIVERGENCES: Record<string, { corpusOnly?: string[]; specOnly?: string[] }> = {
  "a:path": { corpusOnly: ["a:fillToRect"] },
  "a:tcPr": { specOnly: ["a:headers"] },
  "b:Author": { corpusOnly: ["b:Corporate", "b:NameList"] },
  "c:pivotFmt": { specOnly: ["c:txPr"] },
  "c:ser": { corpusOnly: ["c:bubble3D", "c:bubbleSize", "c:explosion", "c:marker", "c:smooth", "c:xVal", "c:yVal"] },
  "c:surface3DChart": { corpusOnly: ["c:varyColors"] },
  "c:tx": { corpusOnly: ["c:strLit", "c:v"] },
  "dgm:extLst": { corpusOnly: ["dgm:ext"] },
  "dgm:styleLbl": { corpusOnly: ["dgm:scene3d", "dgm:sp3d", "dgm:style", "dgm:txPr"] },
  "m:ctrlPr": { corpusOnly: ["w:moveFrom", "w:moveTo"] },
  "m:oMathPara": { corpusOnly: ["w:bookmarkEnd", "w:bookmarkStart", "w:commentRangeEnd", "w:commentRangeStart", "w:customXmlDelRangeEnd", "w:customXmlDelRangeStart", "w:customXmlInsRangeEnd", "w:customXmlInsRangeStart", "w:customXmlMoveFromRangeEnd", "w:customXmlMoveFromRangeStart", "w:customXmlMoveToRangeEnd", "w:customXmlMoveToRangeStart", "w:del", "w:ins", "w:moveFrom", "w:moveFromRangeEnd", "w:moveFromRangeStart", "w:moveTo", "w:moveToRangeEnd", "w:moveToRangeStart", "w:permEnd", "w:permStart", "w:proofErr", "w:r"] },
  "m:r": { specOnly: ["w:contentPart"] },
  "op:property": { corpusOnly: ["vt:cf"] },
  "p:bgPr": { specOnly: ["a:grpFill"] },
  "p:presentation": { specOnly: ["p:smartTags"] },
  "p:progress": { specOnly: ["p:boolVal", "p:clrVal", "p:intVal", "p:strVal"] },
  "p:tnLst": { specOnly: ["p:anim", "p:animClr", "p:animEffect", "p:animMotion", "p:animRot", "p:animScale", "p:audio", "p:cmd", "p:excl", "p:seq", "p:set", "p:video"] },
  "p:to": { corpusOnly: ["p:boolVal", "p:clrVal", "p:fltVal", "p:intVal", "p:strVal"] },
  "v:group": { specOnly: ["v:fill", "v:formulas", "v:handles", "v:imagedata", "v:path", "v:shadow", "v:stroke", "v:textbox", "v:textpath"] },
  "vt:variant": { corpusOnly: ["vt:cf"] },
  "vt:vector": { corpusOnly: ["vt:cf"] },
  "w:background": { corpusOnly: ["v:background"], specOnly: ["w:drawing"] },
  "w:comment": { specOnly: ["w:customXmlDelRangeEnd", "w:customXmlDelRangeStart", "w:customXmlInsRangeEnd", "w:customXmlInsRangeStart", "w:customXmlMoveFromRangeEnd", "w:customXmlMoveFromRangeStart", "w:customXmlMoveToRangeEnd", "w:customXmlMoveToRangeStart", "w:del", "w:ins", "w:moveFrom", "w:moveFromRangeEnd", "w:moveFromRangeStart", "w:moveTo", "w:moveToRangeEnd", "w:moveToRangeStart"] },
  "w:customXml": { corpusOnly: ["w:p", "w:tbl", "w:tc", "w:tr"], specOnly: ["w:customXmlPr"] },
  "w:del": { corpusOnly: ["w:rPr"] },
  "w:ffData": { specOnly: ["w:label", "w:tabIndex"] },
  "w:frame": { specOnly: ["w:longDesc", "w:title"] },
  "w:frameset": { specOnly: ["w:title"] },
  "w:ins": { corpusOnly: ["w:rPr"] },
  "w:moveFrom": { corpusOnly: ["w:rPr"] },
  "w:moveTo": { corpusOnly: ["w:rPr"] },
  "w:object": { corpusOnly: ["v:arc", "v:curve", "v:group", "v:image", "v:line", "v:oval", "v:polyline", "v:rect", "v:roundrect", "v:shape", "v:shapetype"], specOnly: ["w:movie"] },
  "w:pict": { corpusOnly: ["v:arc", "v:curve", "v:group", "v:image", "v:line", "v:oval", "v:polyline", "v:rect", "v:roundrect", "v:shape", "v:shapetype"] },
  "w:r": { specOnly: ["w:contentPart"] },
  "w:rt": { corpusOnly: ["w:customXml", "w:fldSimple", "w:hyperlink", "w:sdt"] },
  "w:rubyBase": { corpusOnly: ["w:customXml", "w:fldSimple", "w:hyperlink", "w:sdt"] },
  "w:sdt": { corpusOnly: ["w:bookmarkEnd", "w:bookmarkStart", "w:commentRangeEnd", "w:commentRangeStart", "w:customXmlDelRangeEnd", "w:customXmlDelRangeStart", "w:customXmlInsRangeEnd", "w:customXmlInsRangeStart", "w:customXmlMoveFromRangeEnd", "w:customXmlMoveFromRangeStart", "w:customXmlMoveToRangeEnd", "w:customXmlMoveToRangeStart", "w:moveFromRangeEnd", "w:moveFromRangeStart", "w:moveToRangeEnd", "w:moveToRangeStart"], specOnly: ["w:sdtEndPr", "w:sdtPr"] },
  "w:sdtContent": { corpusOnly: ["w:p", "w:tbl", "w:tc", "w:tr"] },
  "w:sdtPr": { specOnly: ["w:label", "w:tabIndex"] },
  "w:settings": { corpusOnly: ["w:uiCompat97To2003"], specOnly: ["w:doNotEmbedSmartTags", "w:smartTagType"] },
  "w:tcPr": { specOnly: ["w:headers"] },
  "w:webSettings": { specOnly: ["w:saveSmartTagsAsXml"] },
  "x:anchor": { corpusOnly: ["x:from", "x:to"], specOnly: ["xdr:from", "xdr:to"] },
  "x:bk": { corpusOnly: ["x:extLst"] },
  "x:protectedRange": { specOnly: ["x:securityDescriptor"] },
  "x:r": { corpusOnly: ["x:rPr", "x:t"] },
  "x:row": { corpusOnly: ["x:cell"] },
  "x:workbook": { specOnly: ["x:smartTagPr", "x:smartTagTypes"] },
  "x:worksheet": { specOnly: ["x:smartTags"] }
};

describe('content models', () => {
  it('compares a substantial number of them', () => {
    let compared = 0;
    for (const [element, facts] of Object.entries(spec)) {
      if (facts.childSets?.length && corpusChildren.has(element)) compared += 1;
    }
    expect(compared, 'elements whose content model was checked').toBeGreaterThan(900);
  });

  it('permits only the children the specification permits, except where recorded', () => {
    const unexpected = findChildDivergences().filter(d => {
      const known = KNOWN_CHILD_DIVERGENCES[d.key];
      if (!known) return true;
      return (
        JSON.stringify(known.corpusOnly ?? []) !== JSON.stringify(d.corpusOnly) ||
        JSON.stringify(known.specOnly ?? []) !== JSON.stringify(d.specOnly)
      );
    });

    expect(
      unexpected.map(d => `${d.key} corpus-only=[${d.corpusOnly}] spec-only=[${d.specOnly}]`),
      'new or changed content-model divergence — justify it and record it'
    ).toEqual([]);
  });

  it('has no stale entries in the content-model record', () => {
    const live = new Set(findChildDivergences().map(d => d.key));
    expect(
      Object.keys(KNOWN_CHILD_DIVERGENCES).filter(key => !live.has(key)),
      'recorded content-model divergences that no longer occur — delete them'
    ).toEqual([]);
  });
});
