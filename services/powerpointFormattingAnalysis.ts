/**
 * Composes PresentationML placeholder resolution over a whole package.
 *
 * The counterpart to `wordFormattingAnalysis` and `excelFormattingAnalysis`, with one
 * structural difference that shapes the whole module: **the inheritance chain is
 * carried by implicit relationships.** Nothing in `slide1.xml` points at its layout.
 * You open `ppt/slides/_rels/slide1.xml.rels` and look for the relationship whose
 * `Type` ends in `/slideLayout`, and likewise layout → master and master → theme.
 *
 * That is why a slide missing its layout relationship renders with default formatting
 * and no error: there is no dangling reference to detect, because there was never a
 * reference. Walking the chain here is what makes the failure visible.
 */

import {
  readPlaceholders,
  matchSlideToLayout,
  matchLayoutToMaster,
  readTransform,
  resolveColourMap,
  resolveSchemeColour,
  P_NAMESPACE,
  A_NAMESPACE,
  type Placeholder,
  type Transform,
  type ColourSlot
} from './powerpointResolver';
import { relsPathFor, type PackageParts } from './packageIntegrity';

const parseXml = (xml: string): Document | null => {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
};

/**
 * Follows one implicit relationship out of a part.
 *
 * Matches on the trailing segment of the `Type` URI rather than the whole thing,
 * because Transitional and Strict packages use different namespaces for the same
 * relationship (`schemas.openxmlformats.org/...` versus `purl.oclc.org/ooxml/...`) and
 * a full-URI comparison would report every Strict package as broken.
 */
const followRelationship = (
  parts: PackageParts,
  fromPart: string,
  typeSuffix: string
): { target: string | null; problem: string | null } => {
  const relsPath = relsPathFor(fromPart);
  const relsXml = parts[relsPath];
  if (relsXml === undefined) {
    return { target: null, problem: `${fromPart} has no relationship part, so its ${typeSuffix} cannot be found` };
  }
  const doc = parseXml(relsXml);
  if (!doc) return { target: null, problem: `${relsPath} is not well-formed XML` };

  const matches = Array.from(doc.getElementsByTagName('Relationship')).filter(rel => {
    const type = rel.getAttribute('Type') ?? '';
    return type.includes('/relationships/') && type.endsWith(`/${typeSuffix}`);
  });

  if (matches.length === 0) {
    return {
      target: null,
      problem:
        `${fromPart} has no ${typeSuffix} relationship. This is an implicit relationship with no reference in the XML, ` +
        `so nothing in the part looks wrong - the content simply falls back to defaults`
    };
  }
  if (matches.length > 1) {
    return { target: null, problem: `${fromPart} has ${matches.length} ${typeSuffix} relationships; exactly one is expected` };
  }

  // Targets are relative to the owning part's directory.
  const target = matches[0].getAttribute('Target') ?? '';
  const slash = fromPart.lastIndexOf('/');
  const base = slash === -1 ? [] : fromPart.slice(0, slash).split('/');
  const segments = [...base];
  for (const segment of target.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') segments.pop();
    else segments.push(segment);
  }
  return { target: segments.join('/'), problem: null };
};

export interface SlideChain {
  slidePath: string;
  slide: Document;
  layoutPath: string | null;
  layout: Document | null;
  masterPath: string | null;
  master: Document | null;
  themePath: string | null;
  theme: Document | null;
  problems: string[];
}

/**
 * Walks slide → layout → master → theme for one slide.
 *
 * Each hop is reported separately. A break anywhere below the slide silently degrades
 * rendering rather than failing, so naming which hop broke is most of the value.
 */
export const resolveSlideChain = (
  parts: PackageParts,
  slidePath: string,
  /**
   * An already-parsed slide, when the caller has one.
   *
   * Parsing twice produces two independent DOMs, so a node located in the first will
   * never be identity-equal to its counterpart in the second - and placeholder
   * matching, which compares shape identity, silently reports "not a placeholder".
   */
  parsedSlide?: Document
): SlideChain | null => {
  const slideXml = parts[slidePath];
  if (slideXml === undefined && !parsedSlide) return null;
  const slide = parsedSlide ?? (slideXml !== undefined ? parseXml(slideXml) : null);
  if (!slide) return null;

  const problems: string[] = [];
  const hop = (from: string | null, suffix: string): string | null => {
    if (!from) return null;
    const { target, problem } = followRelationship(parts, from, suffix);
    if (problem) problems.push(problem);
    return target;
  };

  const layoutPath = hop(slidePath, 'slideLayout');
  const layout = layoutPath ? parseXml(parts[layoutPath] ?? '') : null;
  if (layoutPath && !layout) problems.push(`${layoutPath} is referenced but missing or unparseable`);

  const masterPath = hop(layoutPath, 'slideMaster');
  const master = masterPath ? parseXml(parts[masterPath] ?? '') : null;
  if (masterPath && !master) problems.push(`${masterPath} is referenced but missing or unparseable`);

  const themePath = hop(masterPath, 'theme');
  const theme = themePath ? parseXml(parts[themePath] ?? '') : null;
  if (themePath && !theme) problems.push(`${themePath} is referenced but missing or unparseable`);

  return { slidePath, slide, layoutPath, layout, masterPath, master, themePath, theme, problems };
};

export interface PowerpointContext {
  /** Every slide part found, by path. */
  slides: Map<string, Document>;
  parts: PackageParts;
  unresolved: string[];
}

export const loadPowerpointContext = (parts: PackageParts): PowerpointContext => {
  const unresolved: string[] = [];
  const slides = new Map<string, Document>();

  for (const [path, content] of Object.entries(parts)) {
    if (!/^ppt\/slides\/[^/]+\.xml$/.test(path)) continue;
    const doc = parseXml(content);
    if (doc) slides.set(path, doc);
    else unresolved.push(`${path} is not well-formed XML`);
  }
  if (slides.size === 0) unresolved.push('no slide parts were found in the package');

  return { slides, parts, unresolved };
};

export interface ShapeAnalysis {
  slidePath: string;
  placeholder: Placeholder | null;
  transform: Transform;
  /** Where the shape's geometry actually comes from once inheritance is applied. */
  geometrySource: 'own' | 'layout' | 'master' | 'none';
  unresolved: string[];
  explanation: string[];
}

const shapeProperties = (shape: Element): Element | null =>
  Array.from(shape.children).find(
    el => el.namespaceURI === P_NAMESPACE && ['spPr', 'grpSpPr'].includes(el.localName)
  ) ?? null;

/**
 * Resolves one shape on a slide: its placeholder correspondence, where its geometry
 * comes from, and what its scheme colours resolve to.
 */
export const analyzeShape = (
  chain: SlideChain,
  shape: Element
): ShapeAnalysis => {
  const unresolved = [...chain.problems];
  const explanation: string[] = [];

  const slidePlaceholders = readPlaceholders(chain.slide);
  const placeholder = slidePlaceholders.find(p => p.shape === shape) ?? null;

  const transform = readTransform(shapeProperties(shape));
  let geometrySource: ShapeAnalysis['geometrySource'] = transform.inherits ? 'none' : 'own';

  if (placeholder) {
    explanation.push(`Placeholder: type="${placeholder.type}", idx=${placeholder.idx}.`);

    const layoutPlaceholders = chain.layout ? readPlaceholders(chain.layout) : [];
    const match = matchSlideToLayout(placeholder, layoutPlaceholders);
    explanation.push(...match.trace.map(t => `  ${t}`));
    unresolved.push(...match.problems);

    if (transform.inherits) {
      if (match.layoutPlaceholder) {
        const layoutTransform = readTransform(shapeProperties(match.layoutPlaceholder.shape));
        if (!layoutTransform.inherits) {
          geometrySource = 'layout';
          explanation.push('  geometry: inherited from the layout placeholder (the slide shape has no a:xfrm, which is correct)');
        } else if (chain.master) {
          const masterMatch = matchLayoutToMaster(match.layoutPlaceholder, readPlaceholders(chain.master));
          unresolved.push(...masterMatch.problems);
          // Finding a master counterpart is not enough - it must actually carry a
          // transform. A master placeholder that also inherits supplies nothing, and
          // reporting it as the source would name a shape that has no geometry.
          if (masterMatch.layoutPlaceholder) {
            const masterTransform = readTransform(shapeProperties(masterMatch.layoutPlaceholder.shape));
            if (!masterTransform.inherits) {
              geometrySource = 'master';
              explanation.push('  geometry: inherited from the master placeholder');
            }
          }
        }
      }
      if (geometrySource === 'none') {
        explanation.push('  geometry: not resolved - no ancestor placeholder supplied a transform');
      }
    } else {
      explanation.push('  geometry: set directly on this shape, overriding any inherited transform');
    }
  } else {
    explanation.push('Shape is not a placeholder, so it inherits no position or size.');
  }

  // Scheme colours referenced by this shape, resolved through the active map.
  const { map, source } = resolveColourMap(chain.master, chain.slide);
  const clrScheme = chain.theme
    ? chain.theme.getElementsByTagNameNS(A_NAMESPACE, 'clrScheme').item(0)
    : null;

  const schemeRefs = Array.from(shape.getElementsByTagNameNS(A_NAMESPACE, 'schemeClr'));
  if (schemeRefs.length > 0) {
    explanation.push(`Scheme colours (map from the ${source}):`);
    const seen = new Set<string>();
    for (const ref of schemeRefs) {
      const val = ref.getAttribute('val');
      if (!val || seen.has(val)) continue;
      seen.add(val);
      const resolved = resolveSchemeColour(val, map, clrScheme);
      const slot: ColourSlot | null = resolved.slot;
      explanation.push(`  ${val} → ${slot ?? '(unresolved)'}${resolved.value ? '' : ' (no value in the theme)'}`);
      unresolved.push(...resolved.problems);
    }
  }

  if (unresolved.length > 0) {
    explanation.push('Not established by this analysis (do not assert these):', ...unresolved.map(u => `  ${u}`));
  }

  return { slidePath: chain.slidePath, placeholder, transform, geometrySource, unresolved, explanation };
};

const normalizeMarkup = (xml: string): string =>
  xml
    .replace(/\s+xmlns(:[A-Za-z0-9_-]+)?="[^"]*"/g, '')
    .replace(/>\s+</g, '><')
    .replace(/\s+/g, ' ')
    // A serializer emits `<x/>` where a hand-written snippet may say `<x></x>`.
    // Without collapsing them, an empty spPr is enough to defeat the match.
    // The tag name must be followed by whitespace or the bracket, or the pattern
    // matches a different pair: `<p:spPr/></p:sp>` would collapse as `p:sp` + `Pr/`.
    .replace(/<([\w:.-]+)((?:\s[^>]*?)?)\s*><\/\1>/g, '<$1$2/>')
    .trim();

/**
 * Locates a shape from a snippet of markup across every slide.
 *
 * Refuses to guess, as the Word and Excel locators do. Decks are full of identical
 * shapes — placeholders on a template especially — and resolving the wrong one would be
 * a confidently wrong answer under a "Verified" badge.
 */
export const locateShapeByMarkup = (
  context: PowerpointContext,
  rawXml: string
): { slidePath: string; shape: Element } | null => {
  const needle = normalizeMarkup(rawXml);
  if (needle === '') return null;

  const matches: { slidePath: string; shape: Element }[] = [];
  for (const [slidePath, slide] of context.slides) {
    for (const localName of ['sp', 'pic', 'graphicFrame', 'grpSp']) {
      for (const shape of Array.from(slide.getElementsByTagNameNS(P_NAMESPACE, localName))) {
        if (normalizeMarkup(new XMLSerializer().serializeToString(shape)) === needle) {
          matches.push({ slidePath, shape });
        }
      }
    }
  }
  return matches.length === 1 ? matches[0] : null;
};

/** One-call entry point matching the Word and Excel sides. */
export const computePowerpointEvidenceForMarkup = (
  parts: PackageParts,
  rawXml: string
): { lines: string[]; unresolved: string[] } | null => {
  const context = loadPowerpointContext(parts);
  const located = locateShapeByMarkup(context, rawXml);
  if (!located) return null;

  // Reuse the document the shape was found in; see resolveSlideChain's note.
  const chain = resolveSlideChain(parts, located.slidePath, context.slides.get(located.slidePath));
  if (!chain) return null;

  const analysis = analyzeShape(chain, located.shape);
  return {
    lines: [`Slide: ${located.slidePath}`, ...analysis.explanation],
    unresolved: [...context.unresolved, ...analysis.unresolved]
  };
};
