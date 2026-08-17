import { describe, it, expect } from 'vitest';
import {
  resolveAlternateContent,
  readIgnorableNamespaces,
  countAlternateContent,
  MODERN_CONSUMER_NAMESPACES,
  LEGACY_CONSUMER_NAMESPACES
} from '../services/markupCompatibility';

const parse = (xml: string): Document =>
  new DOMParser().parseFromString(xml, 'application/xml');

/**
 * The shape Word actually writes for a text box: the same content twice, modern
 * DrawingML in the Choice and legacy VML in the Fallback.
 */
const shapeWrittenTwice = () => parse(`<?xml version="1.0"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:v="urn:schemas-microsoft-com:vml"
  mc:Ignorable="wps">
  <w:body>
    <w:p><w:r>
      <mc:AlternateContent>
        <mc:Choice Requires="wps">
          <w:drawing><wps:wsp><w:t>Modern</w:t></wps:wsp></w:drawing>
        </mc:Choice>
        <mc:Fallback>
          <w:pict><v:shape><w:t>Legacy</w:t></v:shape></w:pict>
        </mc:Fallback>
      </mc:AlternateContent>
    </w:r></w:p>
  </w:body>
</w:document>`);

const textOf = (doc: Document) =>
  Array.from(doc.getElementsByTagName('w:t')).map(el => el.textContent);

describe('the double-counting problem', () => {
  it('a naive reader sees the same shape twice', () => {
    // This is the bug, demonstrated. Both branches are present in the raw file.
    expect(textOf(shapeWrittenTwice())).toEqual(['Modern', 'Legacy']);
  });

  it('resolving leaves exactly one branch', () => {
    const { document } = resolveAlternateContent(shapeWrittenTwice());
    expect(textOf(document)).toEqual(['Modern']);
  });

  it('a legacy consumer gets the fallback instead', () => {
    // Same file, different consumer. This is how you answer "why does this look
    // different in an older Word?"
    const { document } = resolveAlternateContent(shapeWrittenTwice(), LEGACY_CONSUMER_NAMESPACES);
    expect(textOf(document)).toEqual(['Legacy']);
  });

  it('removes the wrapper entirely, not just its contents', () => {
    const { document } = resolveAlternateContent(shapeWrittenTwice());
    expect(countAlternateContent(document)).toBe(0);
  });
});

describe('choice selection', () => {
  it('reports which branch was taken and what was passed over', () => {
    const { selections } = resolveAlternateContent(shapeWrittenTwice());
    expect(selections).toHaveLength(1);
    expect(selections[0].chose).toBe('choice');
    expect(selections[0].requires).toBe('wps');
    expect(selections[0].path).toContain('w:r');
  });

  it('records rejected choices when falling back', () => {
    const { selections } = resolveAlternateContent(shapeWrittenTwice(), LEGACY_CONSUMER_NAMESPACES);
    expect(selections[0].chose).toBe('fallback');
    expect(selections[0].rejected).toEqual(['wps']);
  });

  it('takes the FIRST matching choice, since order is the author\'s preference', () => {
    const doc = parse(`<?xml version="1.0"?>
<root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <mc:AlternateContent>
    <mc:Choice Requires="a"><w:t>first</w:t></mc:Choice>
    <mc:Choice Requires="w"><w:t>second</w:t></mc:Choice>
    <mc:Fallback><w:t>fallback</w:t></mc:Fallback>
  </mc:AlternateContent>
</root>`);
    const { document } = resolveAlternateContent(doc);
    expect(textOf(document)).toEqual(['first']);
  });

  it('requires ALL listed namespaces, not any of them', () => {
    const doc = parse(`<?xml version="1.0"?>
<root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
      xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:unknown="http://example.invalid/nope">
  <mc:AlternateContent>
    <mc:Choice Requires="w unknown"><w:t>choice</w:t></mc:Choice>
    <mc:Fallback><w:t>fallback</w:t></mc:Fallback>
  </mc:AlternateContent>
</root>`);
    const { document } = resolveAlternateContent(doc);
    expect(textOf(document)).toEqual(['fallback']);
  });

  it('drops the element when nothing matches and there is no fallback', () => {
    const doc = parse(`<?xml version="1.0"?>
<root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
      xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:unknown="http://example.invalid/nope">
  <mc:AlternateContent>
    <mc:Choice Requires="unknown"><w:t>choice</w:t></mc:Choice>
  </mc:AlternateContent>
</root>`);
    const { document, selections } = resolveAlternateContent(doc);
    expect(textOf(document)).toEqual([]);
    expect(selections[0].chose).toBe('nothing');
  });

  it('resolves a prefix using the declaration in scope, not a global guess', () => {
    // The same prefix `x` is bound to different namespaces on the two branches.
    // Resolving it globally would pick the wrong one.
    const doc = parse(`<?xml version="1.0"?>
<root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
      xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <mc:AlternateContent>
    <mc:Choice Requires="x" xmlns:x="http://example.invalid/nope"><w:t>wrong</w:t></mc:Choice>
    <mc:Choice Requires="x" xmlns:x="http://schemas.openxmlformats.org/drawingml/2006/main"><w:t>right</w:t></mc:Choice>
  </mc:AlternateContent>
</root>`);
    const { document } = resolveAlternateContent(doc);
    expect(textOf(document)).toEqual(['right']);
  });
});

describe('nesting and structure', () => {
  it('resolves alternate content nested inside a chosen branch', () => {
    const doc = parse(`<?xml version="1.0"?>
<root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
      xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <mc:AlternateContent>
    <mc:Choice Requires="w">
      <mc:AlternateContent>
        <mc:Choice Requires="a"><w:t>inner</w:t></mc:Choice>
        <mc:Fallback><w:t>inner-fallback</w:t></mc:Fallback>
      </mc:AlternateContent>
    </mc:Choice>
  </mc:AlternateContent>
</root>`);
    const { document, selections } = resolveAlternateContent(doc);
    expect(textOf(document)).toEqual(['inner']);
    expect(countAlternateContent(document)).toBe(0);
    expect(selections).toHaveLength(2);
  });

  it('preserves sibling order when splicing a branch in', () => {
    const doc = parse(`<?xml version="1.0"?>
<root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
      xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:t>before</w:t>
  <mc:AlternateContent>
    <mc:Choice Requires="w"><w:t>middle-a</w:t><w:t>middle-b</w:t></mc:Choice>
  </mc:AlternateContent>
  <w:t>after</w:t>
</root>`);
    const { document } = resolveAlternateContent(doc);
    expect(textOf(document)).toEqual(['before', 'middle-a', 'middle-b', 'after']);
  });

  it('leaves a document with no alternate content untouched', () => {
    const doc = parse(`<?xml version="1.0"?>
<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:t>only</w:t>
</root>`);
    const { document, selections } = resolveAlternateContent(doc);
    expect(textOf(document)).toEqual(['only']);
    expect(selections).toEqual([]);
  });
});

describe('mc:Ignorable', () => {
  it('reports the ignorable namespaces a document declares', () => {
    const ignorable = readIgnorableNamespaces(shapeWrittenTwice());
    expect(ignorable.has('http://schemas.microsoft.com/office/word/2010/wordprocessingShape')).toBe(true);
  });

  it('returns an empty set when none are declared', () => {
    const doc = parse('<?xml version="1.0"?><root/>');
    expect(readIgnorableNamespaces(doc).size).toBe(0);
  });
});

describe('consumer presets', () => {
  it('modern is a superset of legacy', () => {
    for (const ns of LEGACY_CONSUMER_NAMESPACES) {
      expect(MODERN_CONSUMER_NAMESPACES).toContain(ns);
    }
    expect(MODERN_CONSUMER_NAMESPACES.length).toBeGreaterThan(LEGACY_CONSUMER_NAMESPACES.length);
  });

  it('legacy excludes the Office 2010+ extension namespaces', () => {
    expect(LEGACY_CONSUMER_NAMESPACES).not.toContain('http://schemas.microsoft.com/office/word/2010/wordml');
  });

  it('uses the 2012 URI for w15, which ships in Office 2013', () => {
    // The year in an extension URI is not the product year. Encoding the wrong one
    // would silently mis-resolve every Choice that requires w15.
    expect(MODERN_CONSUMER_NAMESPACES).toContain('http://schemas.microsoft.com/office/word/2012/wordml');
  });
});
