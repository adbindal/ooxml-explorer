import { describe, it, expect } from 'vitest';
import {
  STRICT_TO_TRANSITIONAL,
  isStrictXml,
  toTransitionalXml,
  detectConformance,
  normaliseParts,
  conformanceFindings,
  orderedStrictKeys
} from '../services/conformance';
import { analyzePackage } from '../services/analyzers';
import { readBookmarks } from '../services/wordBookmarks';
import { W_NAMESPACE } from '../services/wordStyleResolver';
import { S_NAMESPACE } from '../services/excelStyleResolver';
import type { PackageParts } from '../services/packageIntegrity';

const strictDoc = `<?xml version="1.0"?><w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:body><w:p>
  <w:bookmarkStart w:id="1" w:name="Ref"/><w:r><w:t>hello</w:t></w:r></w:p></w:body></w:document>`;

describe('the mapping table', () => {
  it('targets the Transitional wordprocessingml namespace WITH the year', () => {
    // The mapping this table was cross-checked against gets this one wrong: it maps to
    // '.../wordprocessingml/main', dropping /2006. Copying it wholesale would leave
    // every Strict Word document unreadable, so the value is pinned here.
    expect(STRICT_TO_TRANSITIONAL['http://purl.oclc.org/ooxml/wordprocessingml/main']).toBe(W_NAMESPACE);
    expect(W_NAMESPACE).toContain('/2006/');
  });

  it('agrees with the namespace constants the analyzers actually compare against', () => {
    // A mapping that produces a URI no analyzer recognises is worse than none: the
    // markup parses, matches nothing, and reports a clean file.
    expect(STRICT_TO_TRANSITIONAL['http://purl.oclc.org/ooxml/spreadsheetml/main']).toBe(S_NAMESPACE);
  });

  it('maps every Strict key to a schemas.openxmlformats.org URI', () => {
    for (const [strict, transitional] of Object.entries(STRICT_TO_TRANSITIONAL)) {
      expect(strict, strict).toContain('purl.oclc.org/ooxml/');
      expect(transitional, strict).toContain('schemas.openxmlformats.org/');
    }
  });
});

describe('detection', () => {
  it('recognises Strict markup', () => {
    expect(isStrictXml(strictDoc)).toBe(true);
    expect(isStrictXml('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')).toBe(false);
  });

  it('calls a package Strict when any part is', () => {
    expect(detectConformance({ 'word/document.xml': strictDoc })).toBe('strict');
    expect(detectConformance({ 'word/document.xml': '<w:document/>' })).toBe('transitional');
  });
});

describe('rewriting', () => {
  it('produces markup the existing analyzers can read', () => {
    // The whole point. Before this, a Strict document reported zero bookmarks - not an
    // error, just silence, which is the worst possible failure mode.
    const parsed = new DOMParser().parseFromString(toTransitionalXml(strictDoc), 'application/xml');
    const index = readBookmarks(parsed, 'word/document.xml');

    expect(index.bookmarks.map(b => b.name)).toEqual(['Ref']);
    expect(parsed.documentElement.namespaceURI).toBe(W_NAMESPACE);
  });

  it('rewrites relationship Type attributes too, since they share the base URI', () => {
    const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`;

    expect(toTransitionalXml(rels)).toContain(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
    );
  });

  it('rewrites both members of a prefix pair correctly', () => {
    // '.../drawingml/chart' is a prefix of '.../drawingml/chartDrawing'. Note this
    // passes under EITHER replacement order, because both stems gain the same /2006 —
    // see the ordering tests below for why that is luck rather than design.
    const xml = `<x xmlns:a="http://purl.oclc.org/ooxml/drawingml/chartDrawing" xmlns:b="http://purl.oclc.org/ooxml/drawingml/chart"/>`;
    const out = toTransitionalXml(xml);

    expect(out).toContain('http://schemas.openxmlformats.org/drawingml/2006/chartDrawing');
    expect(out).toContain('http://schemas.openxmlformats.org/drawingml/2006/chart"');
    expect(out).not.toContain('purl.oclc.org');
  });

  it('leaves Transitional markup byte-identical', () => {
    const transitional = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>';

    expect(toTransitionalXml(transitional)).toBe(transitional);
  });

  it('returns the same object for a Transitional package, doing no work', () => {
    const parts: PackageParts = { 'word/document.xml': '<w:document/>' };

    expect(normaliseParts(parts)).toBe(parts);
  });
});

describe('what the reader is told', () => {
  it('reports Strict as a note, not as a fault', () => {
    // Strict is a valid conformance class and usually a deliberate choice. Calling it
    // an error would train a reader to ignore the list.
    const [note] = conformanceFindings({ 'word/document.xml': strictDoc });

    expect(note.severity).toBe('note');
    expect(note.code).toBe('conformance/iso-strict');
  });

  it('says what the namespace mapping does NOT cover', () => {
    // A clean report on a Strict file would otherwise imply coverage this engine does
    // not have. The VML/DrawingML preview difference is the concrete case.
    const [note] = conformanceFindings({ 'word/document.xml': strictDoc });

    expect(note.message).toContain('not a conversion');
    expect(note.message).toContain('VML');
  });

  it('says nothing at all about a Transitional package', () => {
    expect(conformanceFindings({ 'word/document.xml': '<w:document/>' })).toEqual([]);
  });
});

describe('end to end through the registry', () => {
  it('analyses a Strict package instead of reporting silence', () => {
    const run = analyzePackage({ 'word/document.xml': strictDoc });

    expect(run.ran).toContain('conformance');
    expect(run.ran).toContain('bookmark');
    // The unclosed bookmark in the fixture must actually be found, which only happens
    // if the namespace mapping reached the analyzer.
    expect(run.findings.some(f => f.code === 'bookmark/unmatched-start')).toBe(true);
  });

  it('skips the conformance analyzer for a Transitional package', () => {
    const run = analyzePackage({
      'word/document.xml': '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>'
    });

    expect(run.skipped).toContain('conformance');
  });
});

describe('replacement ordering', () => {
  it('applies the longest key first', () => {
    // Pinned as an invariant rather than through a scenario, because no pair in the
    // current table actually needs it: chart is a prefix of chartDrawing, but both
    // gain the same /2006 so either order gives the same string. This guards the next
    // entry, and testing the mechanism is the only honest way to guard something that
    // has no failing case yet.
    const keys = orderedStrictKeys();

    for (let i = 1; i < keys.length; i++) {
      expect(keys[i].length).toBeLessThanOrEqual(keys[i - 1].length);
    }
  });

  it('has at least one key that is a prefix of another, which is why order is considered', () => {
    const keys = orderedStrictKeys();
    const prefixPairs = keys.filter(a => keys.some(b => b !== a && b.startsWith(a)));

    expect(prefixPairs.length).toBeGreaterThan(0);
  });
});
