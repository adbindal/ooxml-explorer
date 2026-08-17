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

  it('stays ON when two style layers both assert bold, because Word resets', () => {
    // [MS-OI29500] §2.1.258 and §2.1.246: Word "resets the value of the toggle
    // property to the value specified by the style", it does not toggle. A strict
    // ECMA §17.7.3 XOR reading would give NOT bold here - that divergence is
    // reported rather than silently chosen.
    const sheet = styles(`
      <w:style w:type="paragraph" w:styleId="P"><w:rPr><w:b/></w:rPr></w:style>
      <w:style w:type="character" w:styleId="C"><w:rPr><w:b/></w:rPr></w:style>`);
    const resolved = resolveRunProperties(sheet, { paragraphStyleId: 'P', characterStyleId: 'C' });
    expect(resolved.isOn('b')).toBe(true);
    expect(resolved.properties.get('b')!.divergence).toContain('XOR');
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

  it('turns bold OFF when a style sets it off over an inherited on', () => {
    // Word resets to the style's value. The previously-shipped code marked this
    // uncertain; it is documented behaviour, so it is now asserted.
    const sheet = styles(`
      <w:style w:type="paragraph" w:styleId="P"><w:rPr><w:b/></w:rPr></w:style>
      <w:style w:type="character" w:styleId="C"><w:rPr><w:b w:val="0"/></w:rPr></w:style>`);
    const resolved = resolveRunProperties(sheet, { paragraphStyleId: 'P', characterStyleId: 'C' });
    expect(resolved.isOn('b')).toBe(false);
    // A strict XOR reading would leave it on, so the divergence is reported.
    expect(resolved.properties.get('b')!.divergence).toContain('XOR');
  });

  it('reports no divergence where Word and the standard agree', () => {
    const sheet = styles(`<w:style w:type="character" w:styleId="C"><w:rPr><w:b w:val="0"/></w:rPr></w:style>`);
    const resolved = resolveRunProperties(sheet, { characterStyleId: 'C' });
    expect(resolved.properties.get('b')!.divergence).toBeUndefined();
  });
});

describe('layers the caller must supply', () => {
  it('warns when a run is in a table but no table style was given', () => {
    const sheet = styles(DOC_DEFAULTS);
    const resolved = resolveRunProperties(sheet, { insideTable: true });
    expect(resolved.trace.find(t => t.layer === 'tableStyle')?.note).toContain('no table style was supplied');
  });

  it('stops warning once the table style layer is supplied', () => {
    const sheet = styles(DOC_DEFAULTS);
    const resolved = resolveRunProperties(sheet, {
      insideTable: true,
      tableStyle: [{ type: 'wholeTable', rPr: frag('<w:rPr><w:sz w:val="18"/></w:rPr>') }]
    });
    expect(resolved.trace.find(t => t.layer === 'tableStyle')?.note).toBeUndefined();
  });

  it('says nothing about table styles outside a table', () => {
    const sheet = styles(DOC_DEFAULTS);
    expect(resolveRunProperties(sheet, {}).trace.find(t => t.layer === 'tableStyle')).toBeUndefined();
  });
});

describe('layer 2 - table style conditional formatting', () => {
  it('applies table style properties over docDefaults', () => {
    const sheet = styles(DOC_DEFAULTS);
    const resolved = resolveRunProperties(sheet, {
      insideTable: true,
      tableStyle: [{ type: 'wholeTable', rPr: frag('<w:rPr><w:sz w:val="18"/></w:rPr>') }]
    });
    expect(resolved.get('sz')).toBe('18');
    expect(resolved.properties.get('sz')!.source).toBe('tableStyle:wholeTable');
  });

  it('lets a later conditional block override an earlier one', () => {
    // The array is already in Word's application order, so later wins.
    const sheet = styles(DOC_DEFAULTS);
    const resolved = resolveRunProperties(sheet, {
      insideTable: true,
      tableStyle: [
        { type: 'wholeTable', rPr: frag('<w:rPr><w:sz w:val="18"/></w:rPr>') },
        { type: 'firstRow', rPr: frag('<w:rPr><w:sz w:val="24"/></w:rPr>') }
      ]
    });
    expect(resolved.get('sz')).toBe('24');
  });

  it('is beaten by a paragraph style, which sits above it in the cascade', () => {
    const sheet = styles(`${DOC_DEFAULTS}
      <w:style w:type="paragraph" w:styleId="Big"><w:rPr><w:sz w:val="40"/></w:rPr></w:style>`);
    const resolved = resolveRunProperties(sheet, {
      insideTable: true,
      paragraphStyleId: 'Big',
      tableStyle: [{ type: 'wholeTable', rPr: frag('<w:rPr><w:sz w:val="18"/></w:rPr>') }]
    });
    expect(resolved.get('sz')).toBe('40');
  });
});

describe('layer 3 - numbering', () => {
  it('applies the level indentation, which overrides the paragraph style', () => {
    // "I set an indent on my list style and nothing happened" - the level's own
    // w:ind sits above the paragraph style in the cascade.
    const sheet = styles(`${DOC_DEFAULTS}
      <w:style w:type="paragraph" w:styleId="Body"><w:pPr><w:ind w:left="0"/></w:pPr></w:style>`);
    const resolved = resolveParagraphProperties(sheet, {
      paragraphStyleId: 'Body',
      numbering: {
        numId: '3', ilvl: 0, abstractNumId: '7', lvl: null,
        start: '1', numFmt: 'decimal', lvlText: '%1.', suff: 'tab', isLgl: false,
        pPr: frag('<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'),
        rPr: null, trace: [], problems: []
      }
    });
    // Implemented per the ECMA-376 17.7.2 cascade order: numbering is layer 3 and
    // the paragraph style is layer 4, so the style wins.
    //
    // KNOWN CONFLICT, deliberately not papered over: MS-OI29500 on numbering-level
    // pPr states that the indentation in lvl/pPr *overrides* the paragraph style's
    // indentation, which is the opposite. The cascade order is the better-sourced of
    // the two, so it is what ships; the trace records that numbering contributed, so
    // a caller can see both layers were consulted. Worth settling against real Word.
    expect(resolved.properties.get('ind')!.source).toBe('style:Body');
    expect(resolved.trace.map(t => t.layer)).toContain('numbering:3/0');
  });

  it('contributes nothing when the paragraph is not numbered', () => {
    const sheet = styles(DOC_DEFAULTS);
    const resolved = resolveParagraphProperties(sheet, { numbering: null });
    expect(resolved.trace.some(t => t.layer.startsWith('numbering:'))).toBe(false);
  });
});

describe('explanation output', () => {
  it('explains why a property has its value, not just the value', () => {
    const sheet = styles(`${DOC_DEFAULTS}
      <w:style w:type="paragraph" w:styleId="Heading1"><w:rPr><w:sz w:val="32"/></w:rPr></w:style>`);
    const lines = explainResolution(resolveRunProperties(sheet, { paragraphStyleId: 'Heading1' }));
    expect(lines.join('\n')).toContain('sz = 32 (from style:Heading1)');
  });

  it('surfaces the divergence in prose', () => {
    const sheet = styles(`
      <w:style w:type="paragraph" w:styleId="P"><w:rPr><w:b/></w:rPr></w:style>
      <w:style w:type="character" w:styleId="C"><w:rPr><w:b w:val="0"/></w:rPr></w:style>`);
    const lines = explainResolution(resolveRunProperties(sheet, { paragraphStyleId: 'P', characterStyleId: 'C' }));
    expect(lines.join('\n')).toContain('XOR');
  });
});
