import { describe, it, expect } from 'vitest';
import { parseStyles } from '../services/wordStyleResolver';
import {
  parseNumbering,
  resolveNumbering,
  readNumberingReference,
  MAX_ILVL
} from '../services/wordNumbering';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const numbering = (body: string) =>
  parseNumbering(`<?xml version="1.0"?><w:numbering ${W}>${body}</w:numbering>`);
const styles = (body: string) =>
  parseStyles(`<?xml version="1.0"?><w:styles ${W}>${body}</w:styles>`);
const frag = (xml: string): Element => {
  const doc = new DOMParser().parseFromString(`<?xml version="1.0"?><root ${W}>${xml}</root>`, 'application/xml');
  return doc.documentElement.firstElementChild!;
};

const noStyles = styles('');

/** Pattern 1: numId → abstractNumId → abstractNum → lvl. */
const simpleList = numbering(`
  <w:abstractNum w:abstractNumId="7">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
      <w:suff w:val="space"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="3"><w:abstractNumId w:val="7"/></w:num>
  <w:num w:numId="4"><w:abstractNumId w:val="7"/></w:num>`);

describe('the direct chain', () => {
  it('resolves numId through abstractNumId to the level', () => {
    const r = resolveNumbering(simpleList, noStyles, { numId: '3', ilvl: 0 })!;
    expect(r.abstractNumId).toBe('7');
    expect(r.numFmt).toBe('decimal');
    expect(r.lvlText).toBe('%1.');
    expect(r.start).toBe('1');
  });

  it('selects the level by ilvl, not by position', () => {
    const r = resolveNumbering(simpleList, noStyles, { numId: '3', ilvl: 1 })!;
    expect(r.numFmt).toBe('lowerLetter');
  });

  it('defaults ilvl to 0 when absent', () => {
    expect(resolveNumbering(simpleList, noStyles, { numId: '3' })!.numFmt).toBe('decimal');
  });

  it('exposes the level pPr, which overrides the paragraph style indentation', () => {
    const r = resolveNumbering(simpleList, noStyles, { numId: '3', ilvl: 0 })!;
    expect(r.pPr).not.toBeNull();
  });

  it('defaults suff to tab when unspecified', () => {
    expect(resolveNumbering(simpleList, noStyles, { numId: '3', ilvl: 1 })!.suff).toBe('tab');
  });

  it('reads an explicit suff', () => {
    expect(resolveNumbering(simpleList, noStyles, { numId: '3', ilvl: 0 })!.suff).toBe('space');
  });
});

describe('numId 0 means remove numbering, not "look up num 0"', () => {
  it('returns null rather than attempting a lookup', () => {
    // The classic bug: resolvers look up num 0, find nothing, fall back to the
    // style's numbering, and render bullets on a paragraph that asked for none.
    expect(resolveNumbering(simpleList, noStyles, { numId: '0', ilvl: 0 })).toBeNull();
  });

  it('returns null even when a num 0 mistakenly exists in the file', () => {
    const withZero = numbering(`
      <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>
      <w:num w:numId="0"><w:abstractNumId w:val="1"/></w:num>`);
    expect(resolveNumbering(withZero, noStyles, { numId: '0', ilvl: 0 })).toBeNull();
  });

  it('distinguishes removal from absence', () => {
    expect(resolveNumbering(simpleList, noStyles, { numId: null })).toBeNull();
  });
});

describe('counters vs definitions', () => {
  it('treats two numIds sharing an abstractNumId as separate counters', () => {
    // Same definition, independent counters. Confusing the two produces either
    // "the list restarts mid-document" or "the second list continues the first".
    const a = resolveNumbering(simpleList, noStyles, { numId: '3', ilvl: 0 })!;
    const b = resolveNumbering(simpleList, noStyles, { numId: '4', ilvl: 0 })!;
    expect(a.abstractNumId).toBe(b.abstractNumId);
    expect(a.numId).not.toBe(b.numId);
  });
});

describe('lvlOverride', () => {
  const overridden = numbering(`
    <w:abstractNum w:abstractNumId="2">
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
    </w:abstractNum>
    <w:num w:numId="5">
      <w:abstractNumId w:val="2"/>
      <w:lvlOverride w:ilvl="0"><w:startOverride w:val="7"/></w:lvlOverride>
    </w:num>
    <w:num w:numId="6">
      <w:abstractNumId w:val="2"/>
      <w:lvlOverride w:ilvl="0">
        <w:lvl w:ilvl="0"><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1"/></w:lvl>
      </w:lvlOverride>
    </w:num>`);

  it('applies a startOverride while keeping the abstract formatting', () => {
    const r = resolveNumbering(overridden, noStyles, { numId: '5', ilvl: 0 })!;
    expect(r.start).toBe('7');
    expect(r.numFmt).toBe('decimal');
  });

  it('lets a full lvl in the override replace the abstract level', () => {
    const r = resolveNumbering(overridden, noStyles, { numId: '6', ilvl: 0 })!;
    expect(r.numFmt).toBe('upperRoman');
  });

  it('ignores an override aimed at a different ilvl', () => {
    const r = resolveNumbering(overridden, noStyles, { numId: '5', ilvl: 1 })!;
    expect(r.start).toBeNull();
  });
});

describe('numStyleLink — the double hop that silently returns nothing', () => {
  // An abstractNum with NO lvl children, only a link to a style, whose numPr points
  // at a different num reaching a different abstractNum that holds the levels.
  const linked = numbering(`
    <w:abstractNum w:abstractNumId="10">
      <w:numStyleLink w:val="MyListStyle"/>
    </w:abstractNum>
    <w:abstractNum w:abstractNumId="11">
      <w:styleLink w:val="MyListStyle"/>
      <w:lvl w:ilvl="0"><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1."/></w:lvl>
    </w:abstractNum>
    <w:num w:numId="20"><w:abstractNumId w:val="10"/></w:num>
    <w:num w:numId="21"><w:abstractNumId w:val="11"/></w:num>`);
  const listStyles = styles(`
    <w:style w:type="numbering" w:styleId="MyListStyle">
      <w:pPr><w:numPr><w:numId w:val="21"/></w:numPr></w:pPr>
    </w:style>`);

  it('follows the link out to styles.xml and back', () => {
    const r = resolveNumbering(linked, listStyles, { numId: '20', ilvl: 0 })!;
    expect(r.abstractNumId).toBe('11');
    expect(r.numFmt).toBe('upperLetter');
  });

  it('records the hop in the trace', () => {
    const r = resolveNumbering(linked, listStyles, { numId: '20', ilvl: 0 })!;
    expect(r.trace.join(' ')).toContain('numStyleLink');
  });

  it('degrades to no level when the linked style is missing', () => {
    const r = resolveNumbering(linked, noStyles, { numId: '20', ilvl: 0 })!;
    expect(r.lvl).toBeNull();
    expect(r.trace.join(' ')).toContain('not found');
  });

  it('survives a numStyleLink cycle instead of hanging', () => {
    const cyclic = numbering(`
      <w:abstractNum w:abstractNumId="30"><w:numStyleLink w:val="Loop"/></w:abstractNum>
      <w:num w:numId="40"><w:abstractNumId w:val="30"/></w:num>`);
    const loopStyles = styles(`
      <w:style w:type="numbering" w:styleId="Loop">
        <w:pPr><w:numPr><w:numId w:val="40"/></w:numPr></w:pPr>
      </w:style>`);
    const r = resolveNumbering(cyclic, loopStyles, { numId: '40', ilvl: 0 })!;
    expect(r.trace.join(' ')).toContain('cycle');
  });
});

describe('problems that schema validation cannot catch', () => {
  it('flags an ilvl Word will refuse to open', () => {
    const r = resolveNumbering(simpleList, noStyles, { numId: '3', ilvl: MAX_ILVL + 1 })!;
    expect(r.problems.join(' ')).toContain('will not open');
  });

  it('accepts the boundary levels', () => {
    expect(resolveNumbering(simpleList, noStyles, { numId: '3', ilvl: 0 })!.problems).toEqual([]);
    expect(resolveNumbering(simpleList, noStyles, { numId: '3', ilvl: MAX_ILVL })!.problems).toEqual([]);
  });

  it('flags a %N token on a bullet level, where tokens are not substituted', () => {
    const bulletWithToken = numbering(`
      <w:abstractNum w:abstractNumId="1">
        <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="%1."/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>`);
    const r = resolveNumbering(bulletWithToken, noStyles, { numId: '1', ilvl: 0 })!;
    expect(r.problems.join(' ')).toContain('bullet');
  });

  it('reports a dangling numId without throwing', () => {
    const r = resolveNumbering(simpleList, noStyles, { numId: '999', ilvl: 0 })!;
    expect(r.lvl).toBeNull();
    expect(r.trace.join(' ')).toContain('not found');
  });
});

describe('reading the reference off a paragraph', () => {
  it('extracts numId and ilvl from pPr/numPr', () => {
    const ref = readNumberingReference(frag('<w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="5"/></w:numPr></w:pPr>'));
    expect(ref).toEqual({ numId: '5', ilvl: 2 });
  });

  it('returns nulls for a paragraph with no numbering', () => {
    expect(readNumberingReference(frag('<w:pPr><w:jc w:val="center"/></w:pPr>'))).toEqual({ numId: null, ilvl: null });
  });

  it('handles a missing pPr', () => {
    expect(readNumberingReference(null)).toEqual({ numId: null, ilvl: null });
  });

  it('reports numId 0 rather than swallowing it', () => {
    // The caller needs to see 0 to distinguish "remove numbering" from "no numbering
    // specified" — they behave differently against an inherited style.
    expect(readNumberingReference(frag('<w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr>')).numId).toBe('0');
  });
});
