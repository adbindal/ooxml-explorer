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
) as { namespace: string; tag: string; attributes: { name: string; values?: string[]; required?: true }[] }[];

const spec = JSON.parse(
  readFileSync(join(__dirname, 'spec-facts.json'), 'utf8')
) as Record<string, Record<string, { valueSets?: string[][]; required?: true }>>;

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
  for (const [element, attributes] of Object.entries(spec)) {
    const ours = corpusByKey.get(element);
    if (!ours) continue;
    for (const [attribute, specAttr] of Object.entries(attributes)) {
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
    for (const [element, attributes] of Object.entries(spec)) {
      const ours = corpusByKey.get(element);
      if (!ours) continue;
      for (const attribute of Object.keys(attributes)) if (ours.has(attribute)) compared += 1;
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
