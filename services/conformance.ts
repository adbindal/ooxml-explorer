/**
 * ISO Strict packages, and how this engine reads them at all.
 *
 * ECMA-376 comes in two conformance classes. Transitional is what Office writes by
 * default and what every analyzer here was built against. **Strict is a different set of
 * namespace URIs for the same vocabulary**, and it does not merely swap the host — it
 * also drops the year segment, which sits in the middle of the URI:
 *
 *   Transitional  http://schemas.openxmlformats.org/spreadsheetml/2006/main
 *   Strict        http://purl.oclc.org/ooxml/spreadsheetml/main
 *
 * That difference already bit this codebase once: `oleObjects.ts` matched on a
 * `/spreadsheetml/2006/main` suffix and so read **nothing** out of Strict packages while
 * looking Strict-tolerant, and the test that should have caught it asserted a URI that
 * does not exist. Every other analyzer compares namespaces by exact equality against a
 * Transitional constant, so all of them had the same blind spot.
 *
 * WHY NORMALISE INSTEAD OF TEACHING SIXTY CALL SITES ABOUT BOTH SPELLINGS.
 *
 * There are about sixty namespace comparisons across fourteen modules, funnelling
 * through six constants. Rewriting them all would be a large change with sixty chances
 * to get one wrong, and every future analyzer would have to remember the rule. Mapping
 * the URIs once, before any analyzer sees the markup, is a single choke point: the
 * analyzers stay simple and exact, and Strict support is a property of the pipeline
 * rather than a discipline each author has to maintain.
 *
 * ⚠️ WHAT THIS DOES NOT DO. It is not a converter. Namespace mapping makes Strict markup
 * *readable* by Transitional-shaped code; it does not reconcile the genuine differences
 * between the two classes. Strict forbids VML, so an OLE object in a Strict package
 * carries a DrawingML preview rather than the `v:imagedata` one this engine looks for;
 * Strict also differs on some date and percentage encodings. Those remain unhandled, and
 * the `conformance` analyzer says so rather than letting a clean report imply otherwise.
 *
 * PROVENANCE. The Strict URIs are the set used by `pjfanning/ooxml-strict-converter`,
 * cross-checked against the pattern in ISO/IEC 29500. **Every Transitional target below
 * was verified independently against the Open XML SDK's published `namespaces.json`** —
 * which is how the bug in that project's own mapping file was caught: it maps
 * wordprocessingml to `.../wordprocessingml/main`, missing the `/2006`. Copying it
 * wholesale would have left every Strict Word document unreadable.
 */

import type { PackageParts } from './packageIntegrity';
import { finding, type Finding } from './findings';

/**
 * Strict URI → Transitional URI.
 *
 * Ordered longest-key-first at use time so a prefix never shadows a longer match.
 */
export const STRICT_TO_TRANSITIONAL: Readonly<Record<string, string>> = {
  'http://purl.oclc.org/ooxml/wordprocessingml/main': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  'http://purl.oclc.org/ooxml/spreadsheetml/main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'http://purl.oclc.org/ooxml/presentationml/main': 'http://schemas.openxmlformats.org/presentationml/2006/main',
  'http://purl.oclc.org/ooxml/drawingml/main': 'http://schemas.openxmlformats.org/drawingml/2006/main',
  'http://purl.oclc.org/ooxml/drawingml/chart': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  'http://purl.oclc.org/ooxml/drawingml/chartDrawing': 'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing',
  'http://purl.oclc.org/ooxml/drawingml/diagram': 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
  'http://purl.oclc.org/ooxml/drawingml/picture': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  'http://purl.oclc.org/ooxml/drawingml/lockedCanvas': 'http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas',
  'http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing':
    'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  'http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing':
    'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  'http://purl.oclc.org/ooxml/officeDocument/relationships':
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/math': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  'http://purl.oclc.org/ooxml/officeDocument/bibliography':
    'http://schemas.openxmlformats.org/officeDocument/2006/bibliography',
  'http://purl.oclc.org/ooxml/officeDocument/customXml':
    'http://schemas.openxmlformats.org/officeDocument/2006/customXml',
  'http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes':
    'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
  'http://purl.oclc.org/ooxml/officeDocument/extendedProperties':
    'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
  'http://purl.oclc.org/ooxml/officeDocument/customProperties':
    'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties',
  'http://purl.oclc.org/ooxml/officeDocument/characteristics':
    'http://schemas.openxmlformats.org/officeDocument/2006/characteristics',
  'http://purl.oclc.org/ooxml/schemaLibrary/main': 'http://schemas.openxmlformats.org/schemaLibrary/2006/main'
};

/** The marker that identifies Strict markup without parsing it. */
const STRICT_PREFIX = 'http://purl.oclc.org/ooxml/';

/**
 * Longest first, so a key that is a prefix of another cannot rewrite the longer one's
 * stem out from under it.
 *
 * ⚠️ No pair in the table above actually requires this today. `.../drawingml/chart` IS a
 * prefix of `.../drawingml/chartDrawing`, but both map to stems that differ only by the
 * same inserted `/2006`, so replacing the shorter first happens to produce the right
 * answer anyway. The ordering is insurance for the next entry, not a fix for a live bug
 * — said plainly so nobody later reads a passing test as proof it is load-bearing.
 */
export const orderedStrictKeys = (): string[] =>
  Object.keys(STRICT_TO_TRANSITIONAL).sort((a, b) => b.length - a.length);

const ORDERED_KEYS = orderedStrictKeys();

export type Conformance = 'strict' | 'transitional';

/** Does this markup use any Strict namespace? */
export const isStrictXml = (xml: string): boolean => xml.includes(STRICT_PREFIX);

/**
 * Rewrites Strict namespace URIs to their Transitional equivalents.
 *
 * Operates on the text rather than on a parsed tree deliberately: a DOM's
 * `namespaceURI` is read-only, so changing it means rebuilding every node and
 * reattaching every attribute — more code, and more ways to drop something. The URIs
 * appear in `xmlns` declarations and in relationship `Type` attributes, and rewriting
 * the base URI fixes both at once because a relationship type is that URI plus a
 * trailing segment.
 */
export const toTransitionalXml = (xml: string): string => {
  if (!isStrictXml(xml)) return xml;
  let out = xml;
  for (const strict of ORDERED_KEYS) {
    out = out.split(strict).join(STRICT_TO_TRANSITIONAL[strict]);
  }
  return out;
};

/** Which conformance class this package is written in. */
export const detectConformance = (parts: PackageParts): Conformance =>
  Object.values(parts).some(isStrictXml) ? 'strict' : 'transitional';

/**
 * Every part with Strict namespaces mapped to Transitional.
 *
 * Returns the input unchanged for a Transitional package, so the common case costs one
 * substring scan per part and allocates nothing.
 *
 * ⚠️ For ANALYSIS ONLY. The editor must keep showing the user their own bytes — nobody
 * wants to open a file and find the namespaces silently rewritten.
 */
export const normaliseParts = (parts: PackageParts): PackageParts => {
  if (detectConformance(parts) === 'transitional') return parts;
  const out: PackageParts = {};
  for (const [path, xml] of Object.entries(parts)) out[path] = toTransitionalXml(xml);
  return out;
};

/**
 * What the reader needs to know about analysing a Strict package.
 *
 * A `note`, not an error: Strict is a fully valid conformance class and choosing it is
 * usually deliberate. But a clean report on a Strict file would otherwise imply coverage
 * this engine does not have, so the limits are stated rather than left to be assumed.
 */
export const conformanceFindings = (parts: PackageParts): Finding[] => {
  if (detectConformance(parts) === 'transitional') return [];

  const strictParts = Object.entries(parts).filter(([, xml]) => isStrictXml(xml));
  return [
    finding(
      'conformance/iso-strict',
      strictParts[0]?.[0] ?? '',
      `This is an ISO Strict package (${strictParts.length} part(s) use purl.oclc.org namespaces). Its namespaces were mapped to their Transitional equivalents so the checks could read it, which is not a conversion: differences the two classes genuinely disagree on are not covered. Strict forbids VML, so an embedded object here carries a DrawingML preview rather than the VML one the OLE check looks for, and a missing preview may be reported that a Strict-aware reader would find.`,
      'Nothing to fix — this is valid. Treat OLE preview findings on this package with suspicion, and re-check anything surprising against a Strict-aware tool.',
      { severity: 'note', silent: true, subject: { conformance: 'strict' } }
    )
  ];
};
