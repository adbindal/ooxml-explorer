/**
 * DrawingML chart extraction, shaped for **translating a chart into another format**.
 *
 * Charts are the largest deviation surface in OOXML — [MS-OI29500] logs 172 variations
 * for Part 1 §21.2, more than the whole of PresentationML — and they are also the thing
 * people most often need to convert rather than merely render. Those are different
 * jobs. A renderer can copy properties across; a converter has to know which properties
 * *mean* something and which are paint.
 *
 * So this module answers three questions a converter must answer and a renderer can
 * ignore:
 *
 *   1. **Where does the data actually live?** A series' values are either a literal
 *      list or a reference to a spreadsheet range, and a reference carries a *cache* of
 *      what those cells last held. The cache is not authoritative — it is what the
 *      producing application wrote at save time — but when the chart lives in a `.docx`
 *      or `.pptx` there is often no workbook to read, and the cache is the only data
 *      there is. Which case applies changes the entire porting strategy, so it is
 *      reported explicitly rather than silently flattened.
 *
 *   2. **What is structure and what is decoration?** Dropping `c:spPr` loses styling.
 *      Dropping `c:order` silently reorders the series. Both are "just properties" in
 *      the markup and the difference is invisible without knowing the format.
 *
 *   3. **What will not survive the trip?** Anything the target format has no concept of
 *      should be named up front, not discovered when the output looks wrong.
 */

export const C_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

/** Chart-type elements that may appear in a plot area, with their series-holder role. */
export const PLOT_TYPES = [
  'areaChart', 'area3DChart', 'lineChart', 'line3DChart', 'stockChart',
  'radarChart', 'scatterChart', 'pieChart', 'pie3DChart', 'doughnutChart',
  'barChart', 'bar3DChart', 'ofPieChart', 'surfaceChart', 'surface3DChart', 'bubbleChart'
] as const;

/**
 * Elements that carry appearance only.
 *
 * Named as a set because "can I drop this?" is the question a converter asks about
 * every element it does not recognise, and getting it wrong in either direction is
 * expensive: keeping paint bloats the target, dropping structure corrupts it.
 */
export const PRESENTATIONAL_ELEMENTS: ReadonlySet<string> = new Set([
  'spPr', 'txPr', 'dLbls', 'marker', 'smooth', 'gapWidth', 'overlap', 'varyColors',
  'roundedCorners', 'autoTitleDeleted', 'shape', 'bubble3D', 'explosion',
  'firstSliceAng', 'holeSize', 'legend', 'view3D', 'floor', 'sideWall', 'backWall',
  'majorGridlines', 'minorGridlines', 'tickLblPos', 'tickMarkSkip', 'lblAlgn', 'lblOffset'
]);

const named = (parent: Element | Document | null, localName: string): Element | null =>
  parent ? parent.getElementsByTagNameNS(C_NAMESPACE, localName).item(0) : null;

const childrenNamed = (parent: Element | null, localName: string): Element[] =>
  parent
    ? Array.from(parent.children).filter(
        el => el.namespaceURI === C_NAMESPACE && el.localName === localName
      )
    : [];

const directChild = (parent: Element | null, localName: string): Element | null =>
  childrenNamed(parent, localName)[0] ?? null;

const valOf = (el: Element | null): string | null => el?.getAttribute('val') ?? null;

export interface DataSource {
  /** How the data is expressed in the markup. */
  kind: 'numRef' | 'strRef' | 'multiLvlStrRef' | 'numLit' | 'strLit' | 'unknown';
  /** The spreadsheet range, when the data is a reference. */
  formula: string | null;
  /** Cached point values, index-ordered. Sparse points are filled with null. */
  cached: (string | null)[];
  /** `c:ptCount` — the number of slots the series declares. */
  declaredCount: number | null;
  /**
   * How many points actually carry a value.
   *
   * Fewer than `declaredCount` is normal, not a defect: a gap in a series omits its
   * point entirely rather than writing an empty one. A converter needs the difference
   * to distinguish "no data here" from "zero".
   */
  presentCount: number;
  /** Point indices that fall outside the declared count, which is a real defect. */
  outOfRangeIndices: number[];
  /** `c:formatCode` from the cache, when present. */
  formatCode: string | null;
  /**
   * Whether the values are *only* obtainable from the cache.
   *
   * True for a reference, because resolving the formula needs the workbook — which a
   * chart embedded in a document may not ship. A converter has to decide whether to
   * trust the cache or refuse; this is the flag that decision hangs on.
   */
  cacheIsOnlySource: boolean;
}

const readDataSource = (holder: Element | null): DataSource | null => {
  if (!holder) return null;

  const kinds = ['numRef', 'strRef', 'multiLvlStrRef', 'numLit', 'strLit'] as const;
  const found = kinds
    .map(kind => ({ kind, element: directChild(holder, kind) }))
    .find(entry => entry.element !== null);

  if (!found?.element) {
    return {
      kind: 'unknown', formula: null, cached: [], declaredCount: null,
      presentCount: 0, outOfRangeIndices: [], formatCode: null, cacheIsOnlySource: false
    };
  }

  const isReference = found.kind.endsWith('Ref');
  const cacheHolder = isReference
    ? directChild(found.element, 'numCache') ?? directChild(found.element, 'strCache')
      ?? directChild(found.element, 'multiLvlStrCache')
    : found.element;

  // Points carry an explicit index and may be sparse: a series with a gap simply omits
  // that point rather than writing an empty one. Reading them in document order would
  // shift every later value by one.
  const points = childrenNamed(cacheHolder, 'pt');
  const declaredCount = valOf(directChild(cacheHolder, 'ptCount'));
  const size = declaredCount !== null ? Number.parseInt(declaredCount, 10) : points.length;
  const cached: (string | null)[] = new Array(Number.isNaN(size) ? points.length : size).fill(null);

  const outOfRangeIndices: number[] = [];
  let presentCount = 0;
  for (const point of points) {
    const index = Number.parseInt(point.getAttribute('idx') ?? '', 10);
    const value = directChild(point, 'v')?.textContent ?? null;
    if (Number.isNaN(index)) {
      cached.push(value);
    } else if (index < cached.length) {
      cached[index] = value;
    } else {
      // The declaration and the data disagree. Silently dropping the point would lose
      // real data; silently growing the array would contradict ptCount. Report it.
      outOfRangeIndices.push(index);
      continue;
    }
    presentCount += 1;
  }

  return {
    kind: found.kind,
    formula: isReference ? directChild(found.element, 'f')?.textContent ?? null : null,
    cached,
    declaredCount: declaredCount === null ? null : Number.parseInt(declaredCount, 10),
    presentCount,
    outOfRangeIndices,
    formatCode: directChild(cacheHolder, 'formatCode')?.textContent ?? null,
    cacheIsOnlySource: isReference
  };
};

export interface ChartSeries {
  /**
   * `c:idx` — the series' stable identity, used by `c:dPt` and `c:dLbl` to attach to it.
   * Not a position.
   */
  index: number | null;
  /**
   * `c:order` — where the series appears in the legend and the stacking order.
   *
   * Distinct from `idx`, and conflating them silently reorders a chart. They coincide
   * often enough in simple charts to hide the bug.
   */
  order: number | null;
  name: { text: string | null; formula: string | null };
  categories: DataSource | null;
  values: DataSource | null;
  /** Bubble sizes, for bubble charts only. */
  bubbleSizes: DataSource | null;
}

export interface ChartAxis {
  id: string | null;
  kind: 'category' | 'value' | 'date' | 'series' | 'unknown';
  /** `c:crossAx` — the id of the axis this one crosses. Axes pair up by id. */
  crossAxisId: string | null;
  orientation: string | null;
  min: string | null;
  max: string | null;
  logBase: string | null;
  numberFormat: {
    code: string | null;
    /**
     * `sourceLinked="1"` means the format is inherited from the source cells rather
     * than stated here — so `code` alone is not the answer, and a converter that reads
     * it without checking this flag will format numbers wrongly.
     */
    sourceLinked: boolean;
  };
  /** `c:delete="1"` — the axis exists structurally but is not drawn. */
  deleted: boolean;
}

export interface ChartPlot {
  /** The chart-type element, e.g. `barChart`. */
  type: string;
  /** `c:grouping` — clustered, stacked, percentStacked, standard. */
  grouping: string | null;
  /** Axis ids this plot is bound to. */
  axisIds: string[];
  series: ChartSeries[];
}

export interface ChartModel {
  plots: ChartPlot[];
  axes: ChartAxis[];
  title: { text: string | null; formula: string | null } | null;
  /** Structural problems that would break a conversion. */
  problems: string[];
  /** Things a converter should decide about before starting. */
  translationNotes: string[];
}

const readSeries = (ser: Element): ChartSeries => {
  const parseIndex = (localName: string) => {
    const raw = valOf(directChild(ser, localName));
    if (raw === null) return null;
    const parsed = Number.parseInt(raw, 10);
    return Number.isNaN(parsed) ? null : parsed;
  };

  const tx = directChild(ser, 'tx');
  const txRef = tx ? directChild(tx, 'strRef') : null;

  return {
    index: parseIndex('idx'),
    order: parseIndex('order'),
    name: {
      text: txRef
        ? directChild(directChild(txRef, 'strCache'), 'pt')?.textContent ?? null
        : directChild(tx, 'v')?.textContent ?? null,
      formula: txRef ? directChild(txRef, 'f')?.textContent ?? null : null
    },
    categories: readDataSource(directChild(ser, 'cat') ?? directChild(ser, 'xVal')),
    values: readDataSource(directChild(ser, 'val') ?? directChild(ser, 'yVal')),
    bubbleSizes: readDataSource(directChild(ser, 'bubbleSize'))
  };
};

const AXIS_KINDS: Record<string, ChartAxis['kind']> = {
  catAx: 'category', valAx: 'value', dateAx: 'date', serAx: 'series'
};

const readAxis = (axis: Element): ChartAxis => {
  const scaling = directChild(axis, 'scaling');
  const numFmt = directChild(axis, 'numFmt');
  return {
    id: valOf(directChild(axis, 'axId')),
    kind: AXIS_KINDS[axis.localName] ?? 'unknown',
    crossAxisId: valOf(directChild(axis, 'crossAx')),
    orientation: valOf(directChild(scaling, 'orientation')),
    min: valOf(directChild(scaling, 'min')),
    max: valOf(directChild(scaling, 'max')),
    logBase: valOf(directChild(scaling, 'logBase')),
    numberFormat: {
      code: numFmt?.getAttribute('formatCode') ?? null,
      sourceLinked: ['1', 'true'].includes(numFmt?.getAttribute('sourceLinked') ?? '')
    },
    deleted: ['1', 'true'].includes(valOf(directChild(axis, 'delete')) ?? '')
  };
};

/**
 * Extracts a chart part into a shape suitable for translating to another format.
 */
export const readChart = (chartXml: string): ChartModel | null => {
  const doc = new DOMParser().parseFromString(chartXml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) return null;

  const plotArea = named(doc, 'plotArea');
  if (!plotArea) return null;

  const problems: string[] = [];
  const translationNotes: string[] = [];

  const plots: ChartPlot[] = [];
  for (const child of Array.from(plotArea.children)) {
    if (child.namespaceURI !== C_NAMESPACE) continue;
    if (!(PLOT_TYPES as readonly string[]).includes(child.localName)) continue;
    plots.push({
      type: child.localName,
      grouping: valOf(directChild(child, 'grouping')),
      axisIds: childrenNamed(child, 'axId').map(el => valOf(el) ?? '').filter(Boolean),
      series: childrenNamed(child, 'ser').map(readSeries)
    });
  }

  const axes: ChartAxis[] = Array.from(plotArea.children)
    .filter(el => el.namespaceURI === C_NAMESPACE && el.localName in AXIS_KINDS)
    .map(readAxis);

  // --- structural checks a converter needs to have passed before it starts ---

  if (plots.length === 0) problems.push('the plot area declares no chart type, so there is nothing to convert');
  if (plots.length > 1) {
    translationNotes.push(
      `this is a combination chart (${plots.map(p => p.type).join(' + ')}); the target format must support ` +
      `more than one plot type sharing an axis, or the conversion has to pick one and drop the rest`
    );
  }

  const axisIds = new Set(axes.map(a => a.id).filter(Boolean));
  for (const plot of plots) {
    for (const id of plot.axisIds) {
      if (!axisIds.has(id)) problems.push(`${plot.type} references axis ${id}, which is not defined in the plot area`);
    }
  }
  for (const axis of axes) {
    if (axis.crossAxisId && !axisIds.has(axis.crossAxisId)) {
      problems.push(`axis ${axis.id} crosses axis ${axis.crossAxisId}, which is not defined`);
    }
  }

  for (const plot of plots) {
    const orders = plot.series.map(s => s.order).filter(o => o !== null);
    if (new Set(orders).size !== orders.length) {
      problems.push(`${plot.type} has series sharing a c:order value, so their display order is ambiguous`);
    }
    for (const series of plot.series) {
      if (series.index !== null && series.order !== null && series.index !== series.order) {
        translationNotes.push(
          `a series has idx=${series.index} but order=${series.order}; these are different concepts ` +
          `(identity versus display position) and must not be conflated when porting`
        );
      }
      const values = series.values;
      if (values && values.outOfRangeIndices.length > 0) {
        problems.push(
          `a series declares ptCount=${values.declaredCount} but carries point indices ` +
          `${values.outOfRangeIndices.join(', ')} beyond it; those values were dropped`
        );
      }
      if (values && values.declaredCount !== null && values.presentCount < values.declaredCount) {
        translationNotes.push(
          `a series declares ${values.declaredCount} points but only ${values.presentCount} carry values; ` +
          `the gaps are genuine missing data and must not be read as zero`
        );
      }
      if (values?.cacheIsOnlySource) {
        translationNotes.push(
          `series values come from the reference "${values.formula ?? 'unknown range'}"; the cached numbers are ` +
          `what the producing application last wrote, not live data. Decide whether to trust the cache or resolve ` +
          `the workbook before converting`
        );
      }
    }
  }

  for (const axis of axes) {
    if (axis.numberFormat.sourceLinked) {
      translationNotes.push(
        `axis ${axis.id} has sourceLinked="1", so its number format comes from the source cells rather than ` +
        `from formatCode; reading formatCode alone would format the values wrongly`
      );
    }
    if (axis.logBase) {
      translationNotes.push(`axis ${axis.id} is logarithmic (base ${axis.logBase}); confirm the target supports log scaling`);
    }
    if (axis.orientation === 'maxMin') {
      translationNotes.push(`axis ${axis.id} is reversed (orientation="maxMin")`);
    }
    if (axis.deleted) {
      translationNotes.push(`axis ${axis.id} is present but not drawn (c:delete="1"); it still participates in scaling`);
    }
  }

  const titleEl = named(doc, 'title');
  const titleRef = titleEl ? named(titleEl, 'strRef') : null;

  return {
    plots,
    axes,
    title: titleEl
      ? {
          text: Array.from(titleEl.getElementsByTagNameNS(
            'http://schemas.openxmlformats.org/drawingml/2006/main', 't'
          )).map(t => t.textContent ?? '').join('') || null,
          formula: titleRef ? directChild(titleRef, 'f')?.textContent ?? null : null
        }
      : null,
    problems,
    translationNotes: [...new Set(translationNotes)]
  };
};

/**
 * Renders a chart model as an ordered account, for a person planning a conversion.
 */
export const explainChart = (model: ChartModel): string[] => {
  const lines: string[] = [];
  if (model.title?.text) lines.push(`Title: "${model.title.text}"`);

  for (const plot of model.plots) {
    lines.push(`${plot.type}${plot.grouping ? ` (${plot.grouping})` : ''}, ${plot.series.length} series:`);
    for (const series of plot.series) {
      const values = series.values;
      lines.push(
        `  · ${series.name.text ?? '(unnamed)'} — idx ${series.index}, order ${series.order}` +
        (values ? `, ${values.cached.length} points from ${values.kind}` : ', no values')
      );
      if (values?.formula) lines.push(`      range: ${values.formula}`);
    }
  }

  for (const axis of model.axes) {
    lines.push(
      `${axis.kind} axis ${axis.id}` +
      (axis.crossAxisId ? ` (crosses ${axis.crossAxisId})` : '') +
      (axis.min || axis.max ? `, range ${axis.min ?? 'auto'}–${axis.max ?? 'auto'}` : ', auto-scaled')
    );
  }

  if (model.problems.length > 0) {
    lines.push('Structural problems — fix before converting:', ...model.problems.map(p => `  ${p}`));
  }
  if (model.translationNotes.length > 0) {
    lines.push('Decide these before converting:', ...model.translationNotes.map(n => `  ${n}`));
  }
  return lines;
};

/**
 * Panel entry point, matching the other formats' `compute…ForMarkup` shape.
 *
 * A chart occupies a whole part, so unlike the Word, Excel and PowerPoint analyses
 * there is nothing to locate — if the open part is a chart, the chart *is* the subject.
 * `rawXml` is therefore unused, and saying so here is cheaper than a caller wondering
 * why selecting different elements gives the same answer.
 */
export const computeChartEvidenceForMarkup = (
  parts: Record<string, string>
): { lines: string[]; unresolved: string[] } | null => {
  const entry = Object.entries(parts).find(([path]) => /charts\/chart[^/]*\.xml$/.test(path));
  if (!entry) return null;

  const model = readChart(entry[1]);
  if (!model) return null;

  return {
    lines: [`Chart part: ${entry[0]}`, ...explainChart(model)],
    // Both lists cap the tier below Verified, and both should: a structural problem
    // means the model is unreliable, and a translation note means a decision has been
    // deferred rather than resolved.
    unresolved: [...model.problems, ...model.translationNotes]
  };
};
