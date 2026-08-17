/**
 * WordprocessingML style resolution.
 *
 * This computes the *effective* formatting of a run or paragraph — what Word will
 * actually render — from markup that never states it directly. It is the thing an
 * experienced OOXML engineer does in their head, and it is the reason
 * "is my markup right if I want it to look like X?" is a computation rather than a
 * lookup. No model is involved, so every answer here is verifiable.
 *
 * ECMA-376 §17.7.2 defines the cascade as six layers, applied in order:
 *
 *   1. docDefaults
 *   2. table style + conditional formatting (gated by w:tblLook)
 *   3. numbering properties
 *   4. paragraph style
 *   5. character style
 *   6. direct formatting
 *
 * All six layers are supported. Layers 2 and 3 are *supplied by the caller* rather
 * than resolved here, because both need context this module does not have: the cell's
 * position within its table (see ./wordTableStyles) and the numbering part (see
 * ./wordNumbering). A caller that omits layer 2 for a run that is in a table gets a
 * trace note saying so, rather than a silently incomplete answer.
 *
 * The second, less obvious axis: within layers 2–5, each style is itself resolved by
 * walking its `w:basedOn` chain to a root before it participates in the cascade above.
 * That roll-up uses four different merge rules depending on the element (see
 * MERGE_STRATEGY), and getting them wrong produces confident, plausible, wrong answers.
 */

import type { ResolvedNumbering } from './wordNumbering';
import type { ConditionalFormatType } from './wordTableStyles';

export const W_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/**
 * The twelve toggle properties (ECMA-376 §17.3.2).
 *
 * These behave differently in a style than as direct formatting: in a style they
 * *toggle* the inherited state, as direct formatting they set an absolute value.
 */
export const TOGGLE_PROPERTIES: ReadonlySet<string> = new Set([
  'b', 'bCs', 'caps', 'emboss', 'i', 'iCs',
  'imprint', 'outline', 'shadow', 'smallCaps', 'strike', 'vanish'
]);

/**
 * How a property element merges when rolling up a `w:basedOn` chain.
 *
 * - `attributes` — the derived element's attributes overlay the base's, one attribute
 *   at a time. Base `<w:spacing w:before="200" w:after="0"/>` plus derived
 *   `<w:spacing w:after="200"/>` yields `before="200" after="200"`.
 * - `replace` — the derived element replaces the base **entirely, dropping attributes
 *   it omits**. This is the surprising one: a derived `<w:top w:val="single" w:sz="18"/>`
 *   over a base that specified `w:color="FF0000"` loses the red. Anyone assuming
 *   attribute merging here will report a border colour that Word does not draw.
 * - `replace` is also the default for simple `w:val` properties, where the later
 *   element simply wins.
 */
type MergeStrategy = 'attributes' | 'replace' | 'children';

const MERGE_STRATEGY: ReadonlyMap<string, MergeStrategy> = new Map<string, MergeStrategy>([
  ['spacing', 'attributes'],
  ['ind', 'attributes'],
  ['rFonts', 'attributes'],
  // Border *containers* merge side by side, so a derived style setting only a top
  // border keeps the base's bottom border.
  ['pBdr', 'children'],
  ['tcBorders', 'children'],
  ['tblBorders', 'children'],
  // ...but each individual side replaces wholesale, dropping omitted attributes.
  ['top', 'replace'],
  ['bottom', 'replace'],
  ['left', 'replace'],
  ['right', 'replace'],
  ['start', 'replace'],
  ['end', 'replace'],
  ['between', 'replace'],
  ['bar', 'replace'],
  ['insideH', 'replace'],
  ['insideV', 'replace']
]);

export interface StyleDefinition {
  styleId: string;
  type: 'paragraph' | 'character' | 'table' | 'numbering' | 'unknown';
  name?: string;
  basedOn?: string;
  isDefault: boolean;
  pPr?: Element;
  rPr?: Element;
  /**
   * The `w:style` element itself.
   *
   * Needed because a table style's conditional blocks (`w:tblStylePr`) and band sizes
   * (`w:tblPr`) are siblings of `pPr`/`rPr`, not children, and a table style may have
   * neither — so reaching them via `pPr.parentElement` fails on exactly the styles
   * that need them most.
   */
  element: Element;
}

export interface StyleSheet {
  docDefaults: { pPr?: Element; rPr?: Element };
  styles: Map<string, StyleDefinition>;
}

/** A single property in the resolved result, with where it came from. */
export interface ResolvedProperty {
  /** Element local name, e.g. `b`, `sz`, `rFonts`. */
  name: string;
  /** The winning element. Null when a toggle resolved to "off". */
  element: Element | null;
  /** `w:val` of the winning element, when it has one. */
  value: string | null;
  /** Human-readable origin, e.g. `docDefaults`, `style:Heading1`, `direct`. */
  source: string;
  /**
   * Set when Word's behaviour and a strict reading of the standard disagree.
   *
   * Not an uncertainty — the value reported is what Word does, which [MS-OI29500]
   * states plainly. This records that a strictly conformant consumer would render the
   * same markup differently, which is exactly the kind of thing someone porting OOXML
   * to another renderer needs to know.
   */
  divergence?: string;
}

export interface ResolutionTrace {
  /** Layer label, e.g. `docDefaults`, `style:Heading1`. */
  layer: string;
  /** Property names this layer contributed or changed. */
  contributed: string[];
  /** Present when a layer was skipped because it is not implemented. */
  note?: string;
}

export interface ResolvedFormatting {
  properties: Map<string, ResolvedProperty>;
  trace: ResolutionTrace[];
  /** Convenience accessor: `w:val` of a property, or null. */
  get(name: string): string | null;
  /** Whether a toggle property is effectively on. */
  isOn(name: string): boolean;
}

const childElements = (parent: Element | undefined): Element[] =>
  parent ? Array.from(parent.children).filter(el => el.namespaceURI === W_NAMESPACE) : [];

const valOf = (element: Element): string | null => element.getAttributeNS(W_NAMESPACE, 'val');

/**
 * Interprets a toggle's `w:val`. Absent means on; `0`/`false`/`off` mean off.
 *
 * The Transitional schema accepts `on`/`off` alongside `0`/`1`/`true`/`false`, and
 * real files use all of them, so a naive `=== "0"` check misreads documents that Word
 * itself wrote.
 */
const toggleIsOn = (element: Element): boolean => {
  const raw = valOf(element);
  if (raw === null) return true;
  return !['0', 'false', 'off'].includes(raw.trim().toLowerCase());
};

/** Parses `styles.xml` into a style sheet keyed by styleId. */
export const parseStyles = (stylesXml: string): StyleSheet => {
  const doc = new DOMParser().parseFromString(stylesXml, 'application/xml');
  const styles = new Map<string, StyleDefinition>();

  const docDefaultsEl = doc.getElementsByTagNameNS(W_NAMESPACE, 'docDefaults').item(0);
  const nested = (parent: Element | null, outer: string, inner: string): Element | undefined => {
    const outerEl = parent?.getElementsByTagNameNS(W_NAMESPACE, outer).item(0);
    return outerEl?.getElementsByTagNameNS(W_NAMESPACE, inner).item(0) ?? undefined;
  };

  const docDefaults = {
    pPr: nested(docDefaultsEl, 'pPrDefault', 'pPr'),
    rPr: nested(docDefaultsEl, 'rPrDefault', 'rPr')
  };

  for (const el of Array.from(doc.getElementsByTagNameNS(W_NAMESPACE, 'style'))) {
    const styleId = el.getAttributeNS(W_NAMESPACE, 'styleId');
    if (!styleId) continue;
    const rawType = el.getAttributeNS(W_NAMESPACE, 'type') ?? '';
    const type = (['paragraph', 'character', 'table', 'numbering'] as const)
      .find(t => t === rawType) ?? 'unknown';

    styles.set(styleId, {
      styleId,
      element: el,
      type,
      name: el.getElementsByTagNameNS(W_NAMESPACE, 'name').item(0)?.getAttributeNS(W_NAMESPACE, 'val') ?? undefined,
      basedOn: el.getElementsByTagNameNS(W_NAMESPACE, 'basedOn').item(0)?.getAttributeNS(W_NAMESPACE, 'val') ?? undefined,
      isDefault: el.getAttributeNS(W_NAMESPACE, 'default') === '1',
      // Direct children only: a style's own pPr, not one nested inside tblStylePr.
      pPr: Array.from(el.children).find(c => c.namespaceURI === W_NAMESPACE && c.localName === 'pPr'),
      rPr: Array.from(el.children).find(c => c.namespaceURI === W_NAMESPACE && c.localName === 'rPr')
    });
  }

  return { docDefaults, styles };
};

/**
 * Walks a style's `basedOn` chain, returning it root-first.
 *
 * Root-first because that is application order — the base contributes, then each
 * derived style overrides. Cycles are tolerated rather than fatal: a malformed
 * styles.xml should degrade, not throw, and Word itself opens such files.
 */
export const styleChain = (sheet: StyleSheet, styleId: string): StyleDefinition[] => {
  const chain: StyleDefinition[] = [];
  const seen = new Set<string>();
  let current: string | undefined = styleId;

  while (current && !seen.has(current)) {
    seen.add(current);
    const style = sheet.styles.get(current);
    if (!style) break;
    chain.unshift(style);
    current = style.basedOn;
  }
  return chain;
};

/**
 * Merges one property element over an existing one, per that element's strategy.
 *
 * Recurses for container elements, which is what makes the border rules work: `w:pBdr`
 * merges side by side, while each side inside it replaces wholesale. Handling only the
 * direct children of `pPr`/`rPr` would apply the container's strategy to the sides and
 * silently lose borders the base style set.
 */
const mergeProperty = (existing: Element | null, incoming: Element): Element => {
  const hasElementChildren = incoming.children.length > 0;
  const strategy: MergeStrategy =
    MERGE_STRATEGY.get(incoming.localName) ?? (hasElementChildren ? 'children' : 'replace');

  if (!existing || strategy === 'replace') {
    return incoming;
  }

  const merged = existing.cloneNode(true) as Element;
  for (const attr of Array.from(incoming.attributes)) {
    merged.setAttributeNS(attr.namespaceURI, attr.name, attr.value);
  }
  if (strategy === 'attributes') {
    return merged;
  }

  // strategy === 'children': merge each incoming child against its counterpart.
  const owner = merged.ownerDocument;
  for (const child of Array.from(incoming.children)) {
    const counterpart = Array.from(merged.children).find(
      c => c.localName === child.localName && c.namespaceURI === child.namespaceURI
    ) ?? null;
    // importNode because the two elements can come from different parsed documents,
    // and appending a foreign node throws.
    const mergedChild = owner.importNode(mergeProperty(counterpart, child), true);
    if (counterpart) merged.replaceChild(mergedChild, counterpart);
    else merged.appendChild(mergedChild);
  }
  return merged;
};

interface LayerInput {
  label: string;
  container?: Element;
  /** Style layers toggle; direct formatting sets absolute values. */
  isStyleLayer: boolean;
}

const applyLayers = (layers: LayerInput[]): { properties: Map<string, ResolvedProperty>; trace: ResolutionTrace[] } => {
  const properties = new Map<string, ResolvedProperty>();
  const trace: ResolutionTrace[] = [];

  for (const layer of layers) {
    const contributed: string[] = [];

    for (const element of childElements(layer.container)) {
      const name = element.localName;
      const previous = properties.get(name) ?? null;

      if (TOGGLE_PROPERTIES.has(name)) {
        const incomingOn = toggleIsOn(element);
        const previouslyOn = previous ? previous.element !== null : false;

        // Word RESETS a toggle to the value the level specifies; it does not toggle.
        // [MS-OI29500] says so twice, naming the standard's rule and rejecting it:
        //
        //   §2.1.258 (Part 1 §17.7.8, Paragraph Styles) and §2.1.246 (§17.7.6, Table
        //   Styles), both: "The standard specifies that the resolved value of the
        //   toggle properties will toggle the previous level (True) or leave it
        //   unchanged (False) ... Word resets the value of the toggle property to the
        //   value specified by the [paragraph|table] style if a value is present."
        //
        // So a style layer behaves exactly like direct formatting here, which is why
        // this branch no longer distinguishes them.
        //
        // Caveat worth keeping: those two notes cover paragraph and table styles.
        // [MS-OI29500] has no equivalent note for character styles, so applying the
        // same rule there is inference from the pattern rather than a quoted
        // statement.
        const effectiveOn = incomingOn;

        // Strict ECMA-376 §17.7.3 instead combines levels with XOR:
        //   value_effective = val_table XOR val_paragraph XOR val_character
        // Where that disagrees with Word, report it as a *divergence* rather than as
        // uncertainty. Which value Word produces is documented and not in doubt; what
        // is worth telling a reader is that a strictly conformant consumer would
        // render this differently.
        const strictEcmaOn = layer.isStyleLayer
          ? (incomingOn ? !previouslyOn : previouslyOn)
          : incomingOn;

        properties.set(name, {
          name,
          element: effectiveOn ? element : null,
          value: effectiveOn ? valOf(element) : null,
          source: layer.label,
          ...(strictEcmaOn !== effectiveOn
            ? {
                divergence:
                  `Word applies ${effectiveOn ? 'on' : 'off'} here; a strictly conformant consumer ` +
                  `following ECMA-376 §17.7.3's XOR rule would apply ${strictEcmaOn ? 'on' : 'off'}`
              }
            : {})
        });
        contributed.push(name);
        continue;
      }

      const merged = mergeProperty(previous?.element ?? null, element);
      properties.set(merged.localName, {
        name: merged.localName,
        element: merged,
        value: valOf(merged),
        source: layer.label
      });
      contributed.push(name);
    }

    if (contributed.length > 0 || layer.container) {
      trace.push({ layer: layer.label, contributed: [...new Set(contributed)] });
    }
  }

  return { properties, trace };
};

const withAccessors = (
  properties: Map<string, ResolvedProperty>,
  trace: ResolutionTrace[]
): ResolvedFormatting => ({
  properties,
  trace,
  get: (name: string) => properties.get(name)?.value ?? null,
  isOn: (name: string) => (properties.get(name)?.element ?? null) !== null
});

/**
 * Cascade layers 2 and 3, supplied by the caller because both need context this module
 * does not have: the cell's position in its table, and the numbering part.
 *
 * Omitting them is legitimate - a run outside a table has no layer 2 - but a caller
 * that omits `tableStyle` for a run that IS in a table gets a trace note saying so,
 * rather than a silently incomplete answer.
 */
export interface CascadeContext {
  /** Conditional blocks in Word's application order; see wordTableStyles. */
  tableStyle?: { type: ConditionalFormatType; pPr?: Element; rPr?: Element }[];
  /** Resolved level for a numbered paragraph; see wordNumbering. */
  numbering?: ResolvedNumbering | null;
}

export interface RunResolutionInput extends CascadeContext {
  /** `w:pStyle` on the containing paragraph. */
  paragraphStyleId?: string;
  /** `w:rStyle` on the run. */
  characterStyleId?: string;
  /** The run's own `w:rPr`. */
  directRPr?: Element;
  /** Set when the run is inside a table, so the skipped layer can be reported. */
  insideTable?: boolean;
}

/**
 * Resolves the effective run properties for a run.
 *
 * Returns both the winning properties and a trace of which layer supplied each one —
 * the trace is the point. "Bold" is a fact; "bold because Heading1 sets it and nothing
 * below overrode it" is an explanation, and only the second is worth showing a user
 * trying to understand why their document looks the way it does.
 */
export const resolveRunProperties = (
  sheet: StyleSheet,
  input: RunResolutionInput
): ResolvedFormatting => {
  const layers: LayerInput[] = [
    { label: 'docDefaults', container: sheet.docDefaults.rPr, isStyleLayer: false }
  ];

  // Layer 2 - table style conditional formatting, in Word's application order.
  for (const block of input.tableStyle ?? []) {
    layers.push({ label: `tableStyle:${block.type}`, container: block.rPr, isStyleLayer: true });
  }
  // Layer 3 - numbering.
  if (input.numbering?.rPr) {
    layers.push({ label: `numbering:${input.numbering.numId}/${input.numbering.ilvl}`, container: input.numbering.rPr, isStyleLayer: false });
  }

  for (const style of input.paragraphStyleId ? styleChain(sheet, input.paragraphStyleId) : []) {
    layers.push({ label: `style:${style.styleId}`, container: style.rPr, isStyleLayer: true });
  }
  for (const style of input.characterStyleId ? styleChain(sheet, input.characterStyleId) : []) {
    layers.push({ label: `charStyle:${style.styleId}`, container: style.rPr, isStyleLayer: true });
  }

  layers.push({ label: 'direct', container: input.directRPr, isStyleLayer: false });

  const { properties, trace } = applyLayers(layers);

  if (input.insideTable && !input.tableStyle) {
    trace.unshift({
      layer: 'tableStyle',
      contributed: [],
      note: 'This run is in a table but no table style was supplied, so layer 2 of the cascade was not applied; a table style may override what is shown here.'
    });
  }
  return withAccessors(properties, trace);
};

export interface ParagraphResolutionInput extends CascadeContext {
  paragraphStyleId?: string;
  /** The paragraph's own `w:pPr`. */
  directPPr?: Element;
  insideTable?: boolean;
}

/**
 * Resolves effective paragraph properties.
 *
 * Simpler than runs: paragraph properties have no toggle semantics, so each layer
 * plainly overrides the one before, subject to the same `basedOn` merge rules.
 */
export const resolveParagraphProperties = (
  sheet: StyleSheet,
  input: ParagraphResolutionInput
): ResolvedFormatting => {
  const layers: LayerInput[] = [
    { label: 'docDefaults', container: sheet.docDefaults.pPr, isStyleLayer: false }
  ];

  // Layer 2 - table style conditional formatting, in Word's application order.
  for (const block of input.tableStyle ?? []) {
    layers.push({ label: `tableStyle:${block.type}`, container: block.pPr, isStyleLayer: false });
  }
  // Layer 3 - numbering. The level's own w:ind overrides the paragraph style's
  // indentation, which is why "I set an indent on my list style and nothing
  // happened" is such a common complaint.
  if (input.numbering?.pPr) {
    layers.push({ label: `numbering:${input.numbering.numId}/${input.numbering.ilvl}`, container: input.numbering.pPr, isStyleLayer: false });
  }

  for (const style of input.paragraphStyleId ? styleChain(sheet, input.paragraphStyleId) : []) {
    layers.push({ label: `style:${style.styleId}`, container: style.pPr, isStyleLayer: false });
  }

  layers.push({ label: 'direct', container: input.directPPr, isStyleLayer: false });

  const { properties, trace } = applyLayers(layers);

  if (input.insideTable && !input.tableStyle) {
    trace.unshift({
      layer: 'tableStyle',
      contributed: [],
      note: 'This paragraph is in a table but no table style was supplied, so layer 2 of the cascade was not applied; a table style may override what is shown here.'
    });
  }
  return withAccessors(properties, trace);
};

/**
 * Renders a resolution as an ordered explanation.
 *
 * Deliberately plain text rather than a component: this is intended to be handed to
 * the AI as pre-verified evidence, where every line is computed rather than recalled.
 */
export const explainResolution = (resolved: ResolvedFormatting): string[] => {
  const lines: string[] = [];
  for (const step of resolved.trace) {
    if (step.note) {
      lines.push(`${step.layer}: ${step.note}`);
    } else if (step.contributed.length > 0) {
      lines.push(`${step.layer} set: ${step.contributed.join(', ')}`);
    }
  }
  for (const property of resolved.properties.values()) {
    const state = property.element === null
      ? 'not applied'
      : property.value === null ? 'applied' : property.value;
    lines.push(
      `${property.name} = ${state} (from ${property.source})` +
      (property.divergence ? ` [${property.divergence}]` : '')
    );
  }
  return lines;
};
