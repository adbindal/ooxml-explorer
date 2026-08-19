import { describe, it, expect } from 'vitest';
import {
  readChart,
  explainChart,
  PRESENTATIONAL_ELEMENTS,
  C_NAMESPACE,
  chartFindings
} from '../services/chartSemantics';

/**
 * Written from the converter's point of view rather than the renderer's. What matters
 * is whether the module surfaces the things you cannot recover once porting has begun:
 * where the data really is, what is structure versus paint, and what the target format
 * will have to account for.
 */

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

const chart = (plotArea: string, extra = '') => `<?xml version="1.0"?>
<c:chartSpace xmlns:c="${C_NAMESPACE}" xmlns:a="${A}">
  <c:chart>${extra}<c:plotArea>${plotArea}</c:plotArea></c:chart>
</c:chartSpace>`;

const series = (idx: number, order: number, name: string) => `
  <c:ser>
    <c:idx val="${idx}"/><c:order val="${order}"/>
    <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx>
    <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/>
      <c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt><c:pt idx="2"><c:v>Mar</c:v></c:pt>
    </c:strCache></c:strRef></c:cat>
    <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:formatCode>0.00</c:formatCode><c:ptCount val="3"/>
      <c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt>
    </c:numCache></c:numRef></c:val>
  </c:ser>`;

const AXES = `
  <c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:crossAx val="222"/></c:catAx>
  <c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:crossAx val="111"/>
    <c:numFmt formatCode="General" sourceLinked="1"/></c:valAx>`;

const simpleBar = chart(`
  <c:barChart><c:grouping val="clustered"/>${series(0, 0, 'Revenue')}
    <c:axId val="111"/><c:axId val="222"/></c:barChart>
  ${AXES}`);

describe('the data model', () => {
  it('extracts plots, grouping and series', () => {
    const model = readChart(simpleBar)!;
    expect(model.plots).toHaveLength(1);
    expect(model.plots[0].type).toBe('barChart');
    expect(model.plots[0].grouping).toBe('clustered');
    expect(model.plots[0].series).toHaveLength(1);
  });

  it('reads the series name from its cache, not just its formula', () => {
    expect(readChart(simpleBar)!.plots[0].series[0].name).toEqual({
      text: 'Revenue', formula: 'Sheet1!$B$1'
    });
  });

  it('keeps categories and values separate', () => {
    const s = readChart(simpleBar)!.plots[0].series[0];
    expect(s.categories!.cached).toEqual(['Jan', 'Feb', 'Mar']);
    expect(s.values!.cached).toEqual(['1', '2', '3']);
  });

  it('carries the format code from the value cache', () => {
    expect(readChart(simpleBar)!.plots[0].series[0].values!.formatCode).toBe('0.00');
  });
});

describe('where the data actually lives - the converter\'s first question', () => {
  it('flags that referenced values are only available from the cache', () => {
    // A chart in a .docx or .pptx may ship no workbook at all, so the cache is
    // frequently the only data there is - and it is what the producing application
    // last wrote, not live data.
    const values = readChart(simpleBar)!.plots[0].series[0].values!;
    expect(values.kind).toBe('numRef');
    expect(values.cacheIsOnlySource).toBe(true);
    expect(values.formula).toBe('Sheet1!$B$2:$B$4');
  });

  it('names the range in the notes', () => {
    expect(readChart(simpleBar)!.translationNotes.join(' ')).toContain('Sheet1!$B$2:$B$4');
  });

  it('does not flag literal data, which needs no workbook', () => {
    const literal = chart(`
      <c:barChart>
        <c:ser><c:idx val="0"/><c:order val="0"/>
          <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>6</c:v></c:pt></c:numLit></c:val>
        </c:ser>
        <c:axId val="111"/><c:axId val="222"/>
      </c:barChart>${AXES}`);
    const values = readChart(literal)!.plots[0].series[0].values!;
    expect(values.kind).toBe('numLit');
    expect(values.cacheIsOnlySource).toBe(false);
    expect(values.cached).toEqual(['5', '6']);
  });

  it('places sparse points by index, not by document order', () => {
    // A gap omits the point entirely. Reading points in order would shift every later
    // value up by one - silently, and only for charts that have gaps.
    const sparse = chart(`
      <c:lineChart>
        <c:ser><c:idx val="0"/><c:order val="0"/>
          <c:val><c:numRef><c:f>S!$A$1:$A$4</c:f><c:numCache><c:ptCount val="4"/>
            <c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="3"><c:v>40</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="111"/><c:axId val="222"/>
      </c:lineChart>${AXES}`);
    expect(readChart(sparse)!.plots[0].series[0].values!.cached).toEqual(['10', null, null, '40']);
  });

  it('reports point indices that fall outside the declared count', () => {
    // The declaration and the data disagree. Dropping the point silently would lose
    // real data, so it is reported as a defect.
    const overflow = chart(`
      <c:barChart>
        <c:ser><c:idx val="0"/><c:order val="0"/>
          <c:val><c:numRef><c:f>S!$A$1:$A$2</c:f><c:numCache><c:ptCount val="2"/>
            <c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="7"><c:v>99</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="111"/><c:axId val="222"/>
      </c:barChart>${AXES}`);
    expect(readChart(overflow)!.problems.join(' ')).toContain('point indices 7');
  });

  it('treats a sparse series as data, not as a defect', () => {
    // A gap is legitimate - it means "no value here". Only the count differs, and a
    // converter needs to know so it does not read the gaps as zero.
    const sparse = chart(`
      <c:lineChart>
        <c:ser><c:idx val="0"/><c:order val="0"/>
          <c:val><c:numRef><c:f>S!$A$1:$A$4</c:f><c:numCache><c:ptCount val="4"/>
            <c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="3"><c:v>40</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="111"/><c:axId val="222"/>
      </c:lineChart>${AXES}`);
    const model = readChart(sparse)!;
    expect(model.problems).toEqual([]);
    expect(model.plots[0].series[0].values!.presentCount).toBe(2);
    expect(model.translationNotes.join(' ')).toContain('must not be read as zero');
  });
});

describe('idx and order are different things', () => {
  const swapped = chart(`
    <c:barChart>${series(0, 1, 'A')}${series(1, 0, 'B')}
      <c:axId val="111"/><c:axId val="222"/></c:barChart>${AXES}`);

  it('reads both without conflating them', () => {
    const model = readChart(swapped)!;
    expect(model.plots[0].series[0]).toMatchObject({ index: 0, order: 1 });
    expect(model.plots[0].series[1]).toMatchObject({ index: 1, order: 0 });
  });

  it('warns when they diverge, because conflating them reorders the chart', () => {
    expect(readChart(swapped)!.translationNotes.join(' ')).toContain('identity versus display position');
  });

  it('says nothing when they coincide, as they do in simple charts', () => {
    // Which is exactly why the bug hides.
    expect(readChart(simpleBar)!.translationNotes.join(' ')).not.toContain('identity versus display');
  });

  it('reports duplicate order values as ambiguous', () => {
    const dupes = chart(`
      <c:barChart>${series(0, 0, 'A')}${series(1, 0, 'B')}
        <c:axId val="111"/><c:axId val="222"/></c:barChart>${AXES}`);
    expect(readChart(dupes)!.problems.join(' ')).toContain('sharing a c:order');
  });
});

describe('axes', () => {
  it('reads kind, id and the cross reference', () => {
    const axes = readChart(simpleBar)!.axes;
    expect(axes.map(a => a.kind)).toEqual(['category', 'value']);
    expect(axes[0]).toMatchObject({ id: '111', crossAxisId: '222' });
  });

  it('flags sourceLinked, where formatCode alone would mislead', () => {
    // sourceLinked="1" means the format comes from the source cells. Reading
    // formatCode without checking it formats the values wrongly.
    expect(readChart(simpleBar)!.translationNotes.join(' ')).toContain('sourceLinked');
  });

  it('reports an axis reference that does not resolve', () => {
    const broken = chart(`
      <c:barChart>${series(0, 0, 'A')}<c:axId val="111"/><c:axId val="999"/></c:barChart>${AXES}`);
    expect(readChart(broken)!.problems.join(' ')).toContain('axis 999');
  });

  it('notes log scaling, reversed orientation and undrawn axes', () => {
    const exotic = chart(`
      <c:barChart>${series(0, 0, 'A')}<c:axId val="111"/><c:axId val="222"/></c:barChart>
      <c:catAx><c:axId val="111"/><c:scaling><c:orientation val="maxMin"/></c:scaling><c:crossAx val="222"/><c:delete val="1"/></c:catAx>
      <c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/><c:logBase val="10"/></c:scaling><c:crossAx val="111"/></c:valAx>`);
    const notes = readChart(exotic)!.translationNotes.join(' ');
    expect(notes).toContain('logarithmic');
    expect(notes).toContain('reversed');
    expect(notes).toContain('still participates in scaling');
  });
});

describe('combination charts', () => {
  const combo = chart(`
    <c:barChart>${series(0, 0, 'Bars')}<c:axId val="111"/><c:axId val="222"/></c:barChart>
    <c:lineChart>${series(1, 1, 'Line')}<c:axId val="111"/><c:axId val="222"/></c:lineChart>
    ${AXES}`);

  it('reads every plot type in the area', () => {
    expect(readChart(combo)!.plots.map(p => p.type)).toEqual(['barChart', 'lineChart']);
  });

  it('warns that the target must support a shared axis, or drop a plot', () => {
    expect(readChart(combo)!.translationNotes.join(' ')).toContain('combination chart');
  });
});

describe('structure versus paint', () => {
  it('classifies appearance-only elements so a converter can drop them safely', () => {
    for (const name of ['spPr', 'txPr', 'dLbls', 'marker', 'gapWidth', 'legend', 'majorGridlines']) {
      expect(PRESENTATIONAL_ELEMENTS.has(name)).toBe(true);
    }
  });

  it('does NOT classify structural elements as paint', () => {
    // Dropping spPr loses styling; dropping order silently reorders the chart. Both
    // are "just properties" in the markup.
    for (const name of ['ser', 'idx', 'order', 'cat', 'val', 'axId', 'crossAx', 'grouping', 'scaling']) {
      expect(PRESENTATIONAL_ELEMENTS.has(name)).toBe(false);
    }
  });
});

describe('robustness and output', () => {
  it('returns null for markup that is not a chart', () => {
    expect(readChart('<?xml version="1.0"?><root/>')).toBeNull();
  });

  it('returns null for malformed XML rather than throwing', () => {
    expect(readChart('<c:chartSpace><unclosed>')).toBeNull();
  });

  it('reports a plot area with no chart type', () => {
    expect(readChart(chart(AXES))!.problems.join(' ')).toContain('nothing to convert');
  });

  it('explains the chart in terms a converter can plan from', () => {
    const prose = explainChart(readChart(simpleBar)!).join('\n');
    expect(prose).toContain('barChart (clustered), 1 series');
    expect(prose).toContain('Revenue');
    expect(prose).toContain('range: Sheet1!$B$2:$B$4');
    expect(prose).toContain('Decide these before converting:');
  });
});

describe('chart findings — charts finally contribute to validation', () => {
  const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
  const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

  const rels = (body: string) =>
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;

  /** A minimal well-formed bar chart with one series referencing a range. */
  const chartXml = (extra = '', formula = 'Sheet1!$B$1:$B$2') => `<?xml version="1.0"?>
    <c:chartSpace ${C} ${R}><c:chart><c:plotArea>
      <c:barChart><c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:val><c:numRef><c:f>${formula}</c:f><c:numCache><c:ptCount val="2"/>
          <c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt>
        </c:numCache></c:numRef></c:val>
      </c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>
      <c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea></c:chart>${extra}</c:chartSpace>`;

  const path = 'word/charts/chart1.xml';

  it('reports a chart whose embedded workbook is gone', () => {
    // The OLE preview problem wearing a different hat: the chart draws perfectly from
    // its cache and can never be edited or refreshed again.
    const parts = {
      [path]: chartXml('<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>'),
      'word/charts/_rels/chart1.xml.rels': rels(
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/data.xlsx"/>`
      )
    };

    const problem = chartFindings(parts, path).find(p => p.code === 'chart/external-data-missing');
    expect(problem?.silent).toBe(true);
    expect(problem?.message).toContain('never be refreshed');
  });

  it('says nothing about external data when the workbook is present', () => {
    const parts = {
      [path]: chartXml('<c:externalData r:id="rId1"/>'),
      'word/charts/_rels/chart1.xml.rels': rels(
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/data.xlsx"/>`
      ),
      'word/embeddings/data.xlsx': 'BINARY'
    };

    expect(chartFindings(parts, path).map(p => p.code)).not.toContain('chart/external-data-missing');
  });

  it('reports a referencing chart that carries no workbook at all', () => {
    // The formulas name cells that exist nowhere in the package, so the cache is the
    // only data there is - correct on screen and impossible to verify.
    const problem = chartFindings({ [path]: chartXml() }, path).find(p => p.code === 'chart/cache-is-only-source');

    expect(problem?.message).toContain('only data there is');
  });

  it('does not claim cache-only for a chart with literal values', () => {
    const literal = `<?xml version="1.0"?><c:chartSpace ${C}><c:chart><c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/><c:order val="0"/>
        <c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>5</c:v></c:pt></c:numLit></c:val>
      </c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>
      <c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea></c:chart></c:chartSpace>`;

    expect(chartFindings({ [path]: literal }, path).map(p => p.code)).not.toContain('chart/cache-is-only-source');
  });

  it('turns a structural problem into an error finding', () => {
    const broken = `<?xml version="1.0"?><c:chartSpace ${C}><c:chart><c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/><c:order val="0"/></c:ser><c:axId val="99"/></c:barChart>
      <c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx>
    </c:plotArea></c:chart></c:chartSpace>`;

    const problem = chartFindings({ [path]: broken }, path).find(p => p.code === 'chart/structural-problem');
    expect(problem?.severity).toBe('error');
    // Structural faults are silent: the chart still draws from its cache.
    expect(problem?.silent).toBe(true);
  });

  it('reports translation risks as notes, not as defects', () => {
    const logAxis = `<?xml version="1.0"?><c:chartSpace ${C}><c:chart><c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/><c:order val="0"/>
        <c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>5</c:v></c:pt></c:numLit></c:val>
      </c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>
      <c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:crossAx val="1"/><c:scaling><c:logBase val="10"/></c:scaling></c:valAx>
    </c:plotArea></c:chart></c:chartSpace>`;

    const notes = chartFindings({ [path]: logAxis }, path).filter(p => p.code === 'chart/translation-risk');
    expect(notes.length).toBeGreaterThan(0);
    for (const n of notes) expect(n.severity).toBe('note');
  });

  it('returns nothing for a part that is not in the package', () => {
    expect(chartFindings({}, path)).toEqual([]);
  });

  it('returns nothing rather than throwing on malformed chart XML', () => {
    expect(chartFindings({ [path]: '<c:chartSpace><unclosed>' }, path)).toEqual([]);
  });
});
