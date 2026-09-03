/**
 * PresentationML placeholder inheritance and theme resolution.
 *
 * PowerPoint's cascade is a *shape correspondence* problem rather than a property
 * cascade: a placeholder on a slide inherits position, size and text style from the
 * matching placeholder on its layout, and that one from the master. Finding the match
 * is the hard part, and the matching rule is documented in exactly one place.
 *
 * Three facts drive almost every bug in this area:
 *
 *   1. **Slide → layout matches on `@idx`, never on `@type`.** [MS-OI29500] on `ph` is
 *      the only source that states the inheritance model at all; ECMA-376 describes
 *      `@idx` merely as something "used when applying templates". Notes slides use a
 *      *different* rule — they match the notes master on `@type` — so one matcher
 *      written for both is wrong for one of them.
 *
 *   2. **An absent `a:xfrm` means "inherit", not "zero".** `off` and `ext` are both
 *      optional, so a placeholder with no transform is the normal, correct encoding.
 *      A generator that helpfully writes `<a:off x="0" y="0"/>` has pinned the shape
 *      to the top-left corner and severed inheritance — reportedly the most common way
 *      a generated deck comes out visually wrong.
 *
 *   3. **`schemeClr` values are map keys, not theme slots.** `tx1` is whatever the
 *      colour map says it is, and the specification's own example maps it to `lt1`.
 *      The three vocabularies overlap on the accents, so the bug stays invisible until
 *      a dark master.
 *
 * Layout → master matching is **not documented anywhere** — see `matchLayoutToMaster`.
 */

export const P_NAMESPACE = 'http://schemas.openxmlformats.org/presentationml/2006/main';
export const A_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** `idx` sentinel meaning "this placeholder corresponds to nothing on the layout". */
export const NO_CORRESPONDENCE_IDX = 4294967295;

/** Theme slot names — the children of `a:clrScheme`. */
export const COLOUR_SLOTS = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'] as const;
/** Map keys — the *attribute names* on `p:clrMap`. Note this is a different list. */
export const COLOUR_MAP_KEYS = ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'] as const;
/** The four values that name a slot directly and therefore bypass the map. */
export const MAP_BYPASSING_VALUES = ['dk1', 'lt1', 'dk2', 'lt2'] as const;

export type ColourSlot = typeof COLOUR_SLOTS[number];

const named = (parent: Element | Document | null, ns: string, localName: string): Element | null =>
  parent ? parent.getElementsByTagNameNS(ns, localName).item(0) : null;

const allNamed = (parent: Element | Document | null, ns: string, localName: string): Element[] =>
  parent ? Array.from(parent.getElementsByTagNameNS(ns, localName)) : [];

export interface Placeholder {
  /** The owning shape (`p:sp`, `p:pic` or `p:graphicFrame`). */
  shape: Element;
  /** `@type`, defaulted to `obj` as the schema specifies. */
  type: string;
  /** `@idx`, defaulted to 0. */
  idx: number;
}

/**
 * Collects the placeholders in a slide, layout or master.
 *
 * `p:ph` is only meaningful on shapes, pictures and graphic frames — PowerPoint
 * rejects it on connectors and groups even though the schema permits it there.
 */
export const readPlaceholders = (part: Document): Placeholder[] =>
  allNamed(part, P_NAMESPACE, 'ph')
    .map(ph => {
      // Walk up to the owning shape rather than assuming a fixed nesting depth.
      let shape: Element | null = ph.parentElement;
      while (shape && !(shape.namespaceURI === P_NAMESPACE && ['sp', 'pic', 'graphicFrame'].includes(shape.localName))) {
        shape = shape.parentElement;
      }
      if (!shape) return null;
      const rawIdx = ph.getAttribute('idx');
      return {
        shape,
        type: ph.getAttribute('type') ?? 'obj',
        idx: rawIdx === null ? 0 : Number.parseInt(rawIdx, 10)
      };
    })
    .filter((p): p is Placeholder => p !== null);

export interface PlaceholderMatch {
  layoutPlaceholder: Placeholder | null;
  trace: string[];
  problems: string[];
}

/**
 * Matches a slide placeholder to its layout placeholder.
 *
 * On `@idx` only. Matching on `@type` instead is the single most common way
 * inheritance silently breaks: two placeholders can share a type, `@idx` defaults to 0
 * so an omitted attribute is still a real index, and a slide placeholder whose `@idx`
 * has no counterpart on the layout inherits nothing at all — with no error anywhere.
 *
 * `@idx = 0xFFFFFFFF` is a documented sentinel for "no correspondence". It is a legal
 * `unsignedInt`, so schema validation cannot catch its misuse.
 */
export const matchSlideToLayout = (
  slidePlaceholder: Placeholder,
  layoutPlaceholders: Placeholder[]
): PlaceholderMatch => {
  const trace: string[] = [];
  const problems: string[] = [];

  if (slidePlaceholder.idx === NO_CORRESPONDENCE_IDX) {
    trace.push('idx is the "no correspondence" sentinel; this placeholder inherits nothing by design');
    return { layoutPlaceholder: null, trace, problems };
  }

  const candidates = layoutPlaceholders.filter(p => p.idx === slidePlaceholder.idx);
  trace.push(`matching slide placeholder idx=${slidePlaceholder.idx} against the layout on idx (never on type)`);

  if (candidates.length === 0) {
    problems.push(
      `no layout placeholder has idx=${slidePlaceholder.idx}, so this shape inherits no position, size or text style. ` +
      `[MS-OI29500] requires a counterpart to exist; nothing reports its absence at runtime`
    );
    return { layoutPlaceholder: null, trace, problems };
  }
  if (candidates.length > 1) {
    problems.push(`the layout has ${candidates.length} placeholders with idx=${slidePlaceholder.idx}; idx must be unique on a layout`);
    return { layoutPlaceholder: null, trace, problems };
  }

  trace.push(`matched layout placeholder type="${candidates[0].type}"`);
  return { layoutPlaceholder: candidates[0], trace, problems };
};

/**
 * Matches a notes-slide placeholder to the notes master.
 *
 * Deliberately a separate function from `matchSlideToLayout`, because the rule is
 * genuinely different: notes slides match on `@type`, not `@idx`. Collapsing the two
 * into one matcher gives the wrong answer for one of them.
 */
export const matchNotesToMaster = (
  notesPlaceholder: Placeholder,
  masterPlaceholders: Placeholder[]
): PlaceholderMatch => {
  const trace = [`matching notes placeholder type="${notesPlaceholder.type}" against the notes master on type (not idx)`];
  const problems: string[] = [];
  const candidates = masterPlaceholders.filter(p => p.type === notesPlaceholder.type);

  if (candidates.length === 0) {
    problems.push(`the notes master has no placeholder of type "${notesPlaceholder.type}"`);
    return { layoutPlaceholder: null, trace, problems };
  }
  return { layoutPlaceholder: candidates[0], trace, problems };
};

/**
 * Matches a layout placeholder to the master.
 *
 * **This rule is not documented.** Neither ECMA-376 nor [MS-OI29500] states how a
 * layout placeholder corresponds to a master placeholder — the only documented rules
 * are slide→layout and notesSlide→notesMaster. What the ecosystem implements (matching
 * on `@type`, folding `ctrTitle` onto `title` and `subTitle` onto `body`) is observed
 * practice, so the result is flagged as such rather than presented as resolved.
 *
 * The constraint that makes this tractable: a master may only carry `title`, `body`,
 * `dt`, `ftr` and `sldNum`, so most layout placeholders have no master counterpart at
 * all and inherit only text styling via `p:txStyles`.
 */
export const MASTER_PLACEHOLDER_TYPES = ['title', 'body', 'dt', 'ftr', 'sldNum'] as const;

export const matchLayoutToMaster = (
  layoutPlaceholder: Placeholder,
  masterPlaceholders: Placeholder[]
): PlaceholderMatch => {
  const trace: string[] = [];
  const problems: string[] = [];

  const folded = layoutPlaceholder.type === 'ctrTitle' ? 'title'
    : layoutPlaceholder.type === 'subTitle' ? 'body'
    : layoutPlaceholder.type;

  if (!(MASTER_PLACEHOLDER_TYPES as readonly string[]).includes(folded)) {
    trace.push(
      `type "${layoutPlaceholder.type}" cannot exist on a master, so there is no shape-level ancestor; ` +
      `only text styling from p:txStyles applies`
    );
    return { layoutPlaceholder: null, trace, problems };
  }

  const candidates = masterPlaceholders.filter(p => p.type === folded);
  trace.push(
    `matched layout type "${layoutPlaceholder.type}" to master type "${folded}" by observed practice — ` +
    `no specification states the layout-to-master rule`
  );
  problems.push('layout-to-master placeholder matching is undocumented; this correspondence is observed practice, not specified behaviour');

  return { layoutPlaceholder: candidates[0] ?? null, trace, problems };
};

export interface Transform {
  offset: { x: number; y: number } | null;
  extent: { cx: number; cy: number } | null;
  /** Child bounding box, group shapes only. */
  childOffset: { x: number; y: number } | null;
  childExtent: { cx: number; cy: number } | null;
  rotation: number | null;
  flipH: boolean;
  flipV: boolean;
  /** True when no `a:xfrm` was present, meaning the shape inherits its geometry. */
  inherits: boolean;
}

const point = (el: Element | null, xName: string, yName: string) => {
  if (!el) return null;
  const x = Number.parseInt(el.getAttribute(xName) ?? '', 10);
  const y = Number.parseInt(el.getAttribute(yName) ?? '', 10);
  return Number.isNaN(x) || Number.isNaN(y) ? null : { x, y };
};

/**
 * Reads a shape's transform.
 *
 * `inherits` is the load-bearing field. An absent `a:xfrm` means the shape takes its
 * geometry from the placeholder it corresponds to; it does not mean the origin.
 *
 * Rotation is in 1/60000 of a degree. ECMA-376's prose for `xfrm` says 1/64000, but its
 * own `ST_Angle` type, the schema and Office all say 1/60000 — a 6.25% error that reads
 * as a rendering artefact rather than a bug.
 */
export const readTransform = (shapeProperties: Element | null): Transform => {
  const xfrm = shapeProperties
    ? Array.from(shapeProperties.children).find(
        el => el.namespaceURI === A_NAMESPACE && el.localName === 'xfrm'
      ) ?? null
    : null;

  if (!xfrm) {
    return {
      offset: null, extent: null, childOffset: null, childExtent: null,
      rotation: null, flipH: false, flipV: false, inherits: true
    };
  }

  const child = (localName: string) =>
    Array.from(xfrm.children).find(el => el.namespaceURI === A_NAMESPACE && el.localName === localName) ?? null;

  const rot = xfrm.getAttribute('rot');
  return {
    offset: point(child('off'), 'x', 'y'),
    extent: (() => {
      const ext = child('ext');
      const p = point(ext, 'cx', 'cy');
      return p ? { cx: p.x, cy: p.y } : null;
    })(),
    childOffset: point(child('chOff'), 'x', 'y'),
    childExtent: (() => {
      const p = point(child('chExt'), 'cx', 'cy');
      return p ? { cx: p.x, cy: p.y } : null;
    })(),
    rotation: rot === null ? null : Number.parseInt(rot, 10),
    flipH: ['1', 'true'].includes(xfrm.getAttribute('flipH') ?? ''),
    flipV: ['1', 'true'].includes(xfrm.getAttribute('flipV') ?? ''),
    inherits: false
  };
};

/** Rotation is stored in sixtieth-thousandths of a degree, not 1/64000. */
export const ROTATION_UNITS_PER_DEGREE = 60000;

export const rotationDegrees = (rotation: number | null): number | null =>
  rotation === null ? null : rotation / ROTATION_UNITS_PER_DEGREE;

/**
 * Maps a child's coordinates out of a group's child coordinate space into the parent's.
 *
 * Children of a group are expressed in the group's *child* space, so reading `a:off`
 * off a nested shape and plotting it on the slide is wrong by both the child offset and
 * the scale. Resizing a group in the UI changes `ext` but leaves `chExt` alone, which is
 * exactly how a group acquires a non-unit scale and why its children's stored sizes stop
 * matching what is drawn.
 *
 * A zero or absent child extent disables scaling on that axis rather than being an error.
 */
export const applyGroupTransform = (
  group: Transform,
  child: { offset: { x: number; y: number } | null; extent: { cx: number; cy: number } | null }
): { offset: { x: number; y: number } | null; extent: { cx: number; cy: number } | null; scale: { x: number; y: number } } => {
  const sx = group.extent && group.childExtent && group.childExtent.cx !== 0
    ? group.extent.cx / group.childExtent.cx
    : 1;
  const sy = group.extent && group.childExtent && group.childExtent.cy !== 0
    ? group.extent.cy / group.childExtent.cy
    : 1;

  const groupOffset = group.offset ?? { x: 0, y: 0 };
  const groupChildOffset = group.childOffset ?? { x: 0, y: 0 };

  return {
    offset: child.offset
      ? {
          x: Math.round(groupOffset.x + (child.offset.x - groupChildOffset.x) * sx),
          y: Math.round(groupOffset.y + (child.offset.y - groupChildOffset.y) * sy)
        }
      : null,
    extent: child.extent
      ? { cx: Math.round(child.extent.cx * sx), cy: Math.round(child.extent.cy * sy) }
      : null,
    scale: { x: sx, y: sy }
  };
};

export interface ColourResolution {
  /** The theme slot finally consulted. */
  slot: ColourSlot | null;
  /** The resolved value element from `a:clrScheme`, when found. */
  value: Element | null;
  trace: string[];
  problems: string[];
}

/**
 * Reads a `p:clrMap` (or `p:clrMapOvr/a:overrideClrMapping`) into key → slot.
 */
export const readColourMap = (mapElement: Element | null): Map<string, ColourSlot> => {
  const map = new Map<string, ColourSlot>();
  if (!mapElement) return map;
  for (const key of COLOUR_MAP_KEYS) {
    const slot = mapElement.getAttribute(key);
    if (slot && (COLOUR_SLOTS as readonly string[]).includes(slot)) {
      map.set(key, slot as ColourSlot);
    }
  }
  return map;
};

/**
 * Resolves `<a:schemeClr val="…"/>` to an entry in the theme's colour scheme.
 *
 * The trap this exists for: `val` is drawn from a 17-value vocabulary, the map's
 * *attribute names* from a 12-value one, and the theme's slots from a different
 * 12-value one. They overlap on the accents, which is why the bug stays invisible until
 * a dark master — where the specification's own example maps `tx1` onto `lt1`, the
 * *light* colour.
 *
 * `dk1`/`lt1`/`dk2`/`lt2` name slots directly and bypass the map entirely.
 */
export const resolveSchemeColour = (
  val: string,
  colourMap: Map<string, ColourSlot>,
  clrScheme: Element | null
): ColourResolution => {
  const trace: string[] = [];
  const problems: string[] = [];

  if (val === 'phClr') {
    trace.push('phClr takes the colour supplied by the style reference that pulled in this format scheme entry');
    return { slot: null, value: null, trace, problems };
  }

  let slot: ColourSlot | null = null;
  if ((MAP_BYPASSING_VALUES as readonly string[]).includes(val)) {
    slot = val as ColourSlot;
    trace.push(`"${val}" names a theme slot directly, bypassing the colour map`);
  } else {
    const mapped = colourMap.get(val);
    if (!mapped) {
      problems.push(`"${val}" is a colour map key with no entry in the active map, so it cannot be resolved`);
      return { slot: null, value: null, trace, problems };
    }
    slot = mapped;
    trace.push(`colour map sends "${val}" to theme slot "${mapped}"`);
  }

  const value = clrScheme
    ? Array.from(clrScheme.children).find(
        el => el.namespaceURI === A_NAMESPACE && el.localName === slot
      ) ?? null
    : null;

  if (!value) problems.push(`the theme colour scheme has no "${slot}" entry`);
  return { slot, value, trace, problems };
};

/**
 * Resolves the effective colour map for a slide or layout.
 *
 * A master carries a complete 12-attribute `p:clrMap`; slides and layouts carry an
 * optional override that is either "use the master's" or a **complete** replacement —
 * there is no partial override.
 */
export const resolveColourMap = (
  master: Document | null,
  child: Document | null
): { map: Map<string, ColourSlot>; source: 'override' | 'master' | 'none' } => {
  const override = child ? named(child, P_NAMESPACE, 'clrMapOvr') : null;
  if (override) {
    const replacement = named(override, A_NAMESPACE, 'overrideClrMapping');
    if (replacement) return { map: readColourMap(replacement), source: 'override' };
  }
  const masterMap = master ? named(master, P_NAMESPACE, 'clrMap') : null;
  if (masterMap) return { map: readColourMap(masterMap), source: 'master' };
  return { map: new Map(), source: 'none' };
};

/**
 * Resolves a `p:style` reference index against the theme's format scheme.
 *
 * The indexing convention is **not uniform**, which is the point of this function:
 *
 * - `fillRef`/`bgRef` address two lists through one integer. 0 and 1000 mean "none",
 *   1–999 index `fillStyleLst`, and **1001 and above index `bgFillStyleLst`**, 1-based,
 *   so 1001 is the *first* background fill.
 * - `lnRef`/`effectRef` address one list each and have **no such offset**. Applying the
 *   1001 rule to them is a bug.
 *
 * ECMA-376 also names the wrong list for `lnRef` — it says `fillStyleLst` where Office
 * reads `lnStyleLst` — so a literal implementation indexes the wrong collection.
 */
export const resolveStyleReference = (
  refKind: 'fillRef' | 'bgRef' | 'lnRef' | 'effectRef',
  idx: number
): { list: 'fillStyleLst' | 'bgFillStyleLst' | 'lnStyleLst' | 'effectStyleLst' | null; position: number | null; note: string } => {
  if (refKind === 'fillRef' || refKind === 'bgRef') {
    if (idx === 0 || idx === 1000) return { list: null, position: null, note: '0 and 1000 both mean no fill' };
    if (idx >= 1001) {
      return { list: 'bgFillStyleLst', position: idx - 1000, note: `${idx} indexes the background fill list, 1-based` };
    }
    return { list: 'fillStyleLst', position: idx, note: `${idx} indexes the fill style list, 1-based` };
  }
  if (refKind === 'lnRef') {
    return {
      list: 'lnStyleLst',
      position: idx,
      note: 'lnRef has no 1001 offset. ECMA-376 names fillStyleLst here, but Office reads lnStyleLst'
    };
  }
  return { list: 'effectStyleLst', position: idx, note: 'effectRef has no 1001 offset' };
};
