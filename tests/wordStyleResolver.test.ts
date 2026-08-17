import { describe, it, expect } from 'vitest';
import {
  parseStyles,
  styleChain,
  resolveRunProperties,
  resolveParagraphProperties,
  explainResolution,
  TOGGLE_PROPERTIES,
  W_NAMESPACE
} from '../services/wordStyleResolver';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const styles = (body: string) =>
  parseStyles(`<?xml version="1.0"?><w:styles ${W}>${body}</w:styles>`);

/** Parses a fragment such as `<w:rPr><w:b/></w:rPr>` into an Element. */
const frag = (xml: string): Element => {
  const doc = new DOMParser().parseFromString(`<?xml version="1.0"?><root ${W}>${xml}</root>`, 'application/xml');
  return doc.documentElement.firstElementChild!;
};

const DOC_DEFAULTS = `
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259"/></w:pPr></w:pPrDefault>
  </w:docDefaults>`;

describe('parsing', () => {
  it('reads docDefaults for both runs and paragraphs', () => {
    const sheet = styles(DOC_DEFAULTS);
    expect(sheet.docDefaults.rPr).toBeDefined();
    expect(sheet.docDefaults.pPr).toBeDefined();
  });

  it('reads style id, type, name and basedOn', () => {
    const sheet = styles(`
      <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
        <w:rPr><w:b/></w:rPr>
      </w:style>`);
    const style = sheet.styles.get('Heading1')!;
    expect(style.type).toBe('paragraph');
    expect(style.name).toBe('heading 1');
    expect(style.basedOn).toBe('Normal');
  });

  it('does not mistake a tblStylePr child for the style\'s own pPr', () => {
    // A table style's conditional formatting nests pPr one level deeper. Picking it
    // up as the style's own would apply first-row formatting to every paragraph.
    const sheet = styles(`
      <w:style w:type="table" w:styleId="Grid">
        <w:tblStylePr w:type="firstRow"><w:pPr><w:jc w:val="center"/></w:pPr></w:tblStylePr>
      </w:style>`);
    expect(sheet.styles.get('Grid')!.pPr).toBeUndefined();
  });
});

describe('basedOn chains', () => {
  const sheet = styles(`
    <w:style w:type="paragraph" w:styleId="Normal"><w:rPr><w:sz w:val="22"/></w:rPr></w:style>
    <w:style w:type="paragraph" w:styleId="Body"><w:basedOn w:val="Normal"/></w:style>
    <w:style w:type="paragraph" w:styleId="Quote"><w:basedOn w:val="Body"/></w:style>`);

  it('returns the chain root-first, which is application order', () => {
    expect(styleChain(sheet, 'Quote').map(s => s.styleId)).toEqual(['Normal', 'Body', 'Quote']);
  });

  it('survives a cycle instead of hanging', () => {
    const cyclic = styles(`
      <w:style w:type="paragraph" w:styleId="A"><w:basedOn w:val="B"/></w:style>
      <w:style w:type="paragraph" w:styleId="B"><w:basedOn w:val="A"/></w:style>`);
    expect(styleChain(cyclic, 'A').map(s => s.styleId).sort()).toEqual(['A', 'B']);
  });

  it('stops at a missing base rather than throwing', () => {
    const dangling = styles(`<w:style w:type="paragraph" w:styleId="X"><w:basedOn w:val="Ghost"/></w:style>`);
    expect(styleChain(dangling, 'X').map(s => s.styleId)).toEqual(['X']);
  });
});

describe('merge semantics along a chain', () => {
  it('merges spacing attribute by attribute', () => {
    // Base sets before+after; derived sets only after. before must survive.
    const sheet = styles(`${DOC_DEFAULTS}
      <w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:spacing w:before="200" w:after="0"/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="Derived"><w:basedOn w:val="Base"/><w:pPr><w:spacing w:after="200"/></w:pPr></w:style>`);
    const spacing = resolveParagraphProperties(sheet, { paragraphStyleId: 'Derived' }).properties.get('spacing')!;
    expect(spacing.element!.getAttributeNS(W_NAMESPACE, 'before')).toBe('200');
    expect(spacing.element!.getAttributeNS(W_NAMESPACE, 'after')).toBe('200');
  });

  it('replaces an individual border side wholesale, dropping omitted attributes', () => {
    // The surprising rule. The derived w:top omits w:color, so the base's red is
    // gone entirely rather than inherited. Assuming attribute-merge here would
    // report a border colour Word does not draw.
    const sheet = styles(`
      <w:style w:type="paragraph" w:styleId="Base">
        <w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:color="FF0000"/></w:pBdr></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Derived">
        <w:basedOn w:val="Base"/>
        <w:pPr><w:pBdr><w:top w:val="single" w:sz="18"/></w:pBdr></w:pPr>
      </w:style>`);
    const pBdr = resolveParagraphProperties(sheet, { paragraphStyleId: 'Derived' }).properties.get('pBdr')!;
    const top = pBdr.element!.getElementsByTagNameNS(W_NAMESPACE, 'top').item(0)!;
    expect(top.getAttributeNS(W_NAMESPACE, 'sz')).toBe('18');
    expect(top.getAttributeNS(W_NAMESPACE, 'color')).toBeNull();
  });

  it('keeps a border side the derived style did not mention', () => {
    // The container merges side by side, so setting only a top border must not
    // silently delete the bottom border the base style set.
    const sheet = styles(`
      <w:style w:type="paragraph" w:styleId="Base">
        <w:pPr><w:pBdr><w:top w:val="single" w:sz="4"/><w:bottom w:val="double" w:sz="8"/></w:pBdr></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Derived">
        <w:basedOn w:val="Base"/>
        <w:pPr><w:pBdr><w:top w:val="dashed" w:sz="12"/></w:pBdr></w:pPr>
      </w:style>`);
    const pBdr = resolveParagraphProperties(sheet, { paragraphStyleId: 'Derived' }).properties.get('pBdr')!;
    const top = pBdr.element!.getElementsByTagNameNS(W_NAMESPACE, 'top').item(0)!;
    const bottom = pBdr.element!.getElementsByTagNameNS(W_NAMESPACE, 'bottom').item(0);
    expect(top.getAttributeNS(W_NAMESPACE, 'val')).toBe('dashed');
    expect(bottom).not.toBeNull();
    expect(bottom!.getAttributeNS(W_NAMESPACE, 'val')).toBe('double');
  });

  it('lets a later style override a simple val property', () => {
    const sheet = styles(`
      <w:style w:type="paragraph" w:styleId="Base"><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
      <w:style w:type="paragraph" w:styleId="Derived"><w:basedOn w:val="Base"/><w:rPr><w:sz w:val="28"/></w:rPr></w:style>`);
    expect(resolveRunProperties(sheet, { paragraphStyleId: 'Derived' }).get('sz')).toBe('28');
  });
});

describe('the cascade', () => {
  const sheet = styles(`${DOC_DEFAULTS}
    <w:style w:type="paragraph" w:styleId="Heading1"><w:rPr><w:sz w:val="32"/><w:b/></w:rPr></w:style>
    <w:style w:type="character" w:styleId="Emphasis"><w:rPr><w:i/></w:rPr></w:style>`);

  it('falls back to docDefaults when nothing else speaks', () => {
    const resolved = resolveRunProperties(sheet, {});
    expect(resolved.get('sz')).toBe('22');
  });

  it('lets a paragraph style beat docDefaults', () => {
    expect(resolveRunProperties(sheet, { paragraphStyleId: 'Heading1' }).get('sz')).toBe('32');
  });

  it('lets direct formatting beat everything', () => {
    const resolved = resolveRunProperties(sheet, {
      paragraphStyleId: 'Heading1',
      directRPr: frag('<w:rPr><w:sz w:val="48"/></w:rPr>')
    });
    expect(resolved.get('sz')).toBe('48');
    expect(resolved.properties.get('sz')!.source).toBe('direct');
  });

  it('applies character styles over paragraph styles', () => {
    const resolved = resolveRunProperties(sheet, {
      paragraphStyleId: 'Heading1',
      characterStyleId: 'Emphasis'
    });
    expect(resolved.isOn('i')).toBe(true);
    expect(resolved.isOn('b')).toBe(true);
  });

  it('records where each property came from', () => {
    const resolved = resolveRunProperties(sheet, { paragraphStyleId: 'Heading1' });
    expect(resolved.properties.get('sz')!.source).toBe('style:Heading1');
    expect(resolved.properties.get('rFonts')!.source).toBe('docDefaults');
  });
});

describe('toggle properties', () => {
  it('covers the twelve defined toggles', () => {
    expect(TOGGLE_PROPERTIES.size).toBe(12);
    for (const name of ['b', 'i', 'caps', 'smallCaps', 'strike', 'vanish', 'outline', 'shadow', 'emboss', 'imprint', 'bCs', 'iCs']) {
      expect(TOGGLE_PROPERTIES.has(name)).toBe(true);
    }
  });

  it('toggles OFF when two style layers both assert bold', () => {
    // The classic surprise: bold from a paragraph style plus bold from a character
    // style yields NOT bold, because toggles XOR across style types.
    const sheet = styles(`
      <w:style w:type="paragraph" w:styleId="P"><w:rPr><w:b/></w:rPr></w:style>
      <w:style w:type="character" w:styleId="C"><w:rPr><w:b/></w:rPr></w:style>`);
    const resolved = resolveRunProperties(sheet, { paragraphStyleId: 'P', characterStyleId: 'C' });
    expect(resolved.isOn('b')).toBe(false);
  });

  it('does NOT toggle for direct formatting, which is absolute', () => {
    const sheet = styles(`<w:style w:type="paragraph" w:styleId="P"><w:rPr><w:b/></w:rPr></w:style>`);
    const resolved = resolveRunProperties(sheet, {
      paragraphStyleId: 'P',
      directRPr: frag('<w:rPr><w:b/></w:rPr>')
    });
    expect(resolved.isOn('b')).toBe(true);
  });

  it('honours direct b w:val="0" as an absolute off', () => {
    const sheet = styles(`<w:style w:type="paragraph" w:styleId="P"><w:rPr><w:b/></w:rPr></w:style>`);
    const resolved = resolveRunProperties(sheet, {
      paragraphStyleId: 'P',
      directRPr: frag('<w:rPr><w:b w:val="0"/></w:rPr>')
    });
    expect(resolved.isOn('b')).toBe(false);
  });

  it('reads the Transitional on/off spellings, not just 0/1', () => {
    const sheet = styles('');
    expect(resolveRunProperties(sheet, { directRPr: frag('<w:rPr><w:b w:val="off"/></w:rPr>') }).isOn('b')).toBe(false);
    expect(resolveRunProperties(sheet, { directRPr: frag('<w:rPr><w:b w:val="on"/></w:rPr>') }).isOn('b')).toBe(true);
    expect(resolveRunProperties(sheet, { directRPr: frag('<w:rPr><w:b w:val="false"/></w:rPr>') }).isOn('b')).toBe(false);
  });

  it('flags the one case where the spec and Word\'s reference implementation disagree', () => {
    // A style turning bold off over an inherited bold. Rather than silently picking
    // a winner, the resolver marks the result uncertain.
    const sheet = styles(`
      <w:style w:type="paragraph" w:styleId="P"><w:rPr><w:b/></w:rPr></w:style>
      <w:style w:type="character" w:styleId="C"><w:rPr><w:b w:val="0"/></w:rPr></w:style>`);
    const resolved = resolveRunProperties(sheet, { paragraphStyleId: 'P', characterStyleId: 'C' });
    expect(resolved.properties.get('b')!.uncertain).toBe(true);
  });

  it('does not flag uncertainty when nothing was inherited', () => {
    const sheet = styles(`<w:style w:type="character" w:styleId="C"><w:rPr><w:b w:val="0"/></w:rPr></w:style>`);
    const resolved = resolveRunProperties(sheet, { characterStyleId: 'C' });
    expect(resolved.properties.get('b')!.uncertain).toBeUndefined();
  });
});

describe('unimplemented layers are reported, not hidden', () => {
  it('warns that table styles were not considered', () => {
    const sheet = styles(DOC_DEFAULTS);
    const resolved = resolveRunProperties(sheet, { insideTable: true });
    const note = resolved.trace.find(t => t.layer === 'tableStyle')?.note;
    expect(note).toContain('not resolved yet');
  });

  it('says nothing about table styles outside a table', () => {
    const sheet = styles(DOC_DEFAULTS);
    expect(resolveRunProperties(sheet, {}).trace.find(t => t.layer === 'tableStyle')).toBeUndefined();
  });
});

describe('explanation output', () => {
  it('explains why a property has its value, not just the value', () => {
    const sheet = styles(`${DOC_DEFAULTS}
      <w:style w:type="paragraph" w:styleId="Heading1"><w:rPr><w:sz w:val="32"/></w:rPr></w:style>`);
    const lines = explainResolution(resolveRunProperties(sheet, { paragraphStyleId: 'Heading1' }));
    expect(lines.join('\n')).toContain('sz = 32 (from style:Heading1)');
  });

  it('surfaces the uncertainty marker in prose', () => {
    const sheet = styles(`
      <w:style w:type="paragraph" w:styleId="P"><w:rPr><w:b/></w:rPr></w:style>
      <w:style w:type="character" w:styleId="C"><w:rPr><w:b w:val="0"/></w:rPr></w:style>`);
    const lines = explainResolution(resolveRunProperties(sheet, { paragraphStyleId: 'P', characterStyleId: 'C' }));
    expect(lines.join('\n')).toContain('uncertain');
  });
});
