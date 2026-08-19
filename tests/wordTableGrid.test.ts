import { describe, it, expect } from 'vitest';
import {
  readTableGrids,
  tableGridFindings,
  computeTableEvidenceForMarkup
} from '../services/wordTableGrid';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const doc = (body: string) =>
  new DOMParser().parseFromString(
    `<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`,
    'application/xml'
  );

/** A grid of `n` equal columns, 2880 twips each unless widths are given. */
const grid = (...widths: number[]) =>
  `<w:tblGrid>${widths.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
const gridOf = (n: number) => grid(...Array<number>(n).fill(2880));

/** A cell with optional tcPr contents. */
const tc = (pr = '') => `<w:tc>${pr ? `<w:tcPr>${pr}</w:tcPr>` : ''}<w:p/></w:tc>`;
const span = (n: number) => `<w:gridSpan w:val="${n}"/>`;
const tr = (...cells: string[]) => `<w:tr>${cells.join('')}</w:tr>`;
const tbl = (...parts: string[]) => `<w:tbl>${parts.join('')}</w:tbl>`;

const codes = (findings: { code: string }[]) => findings.map(f => f.code);

describe('readTableGrids — the row arithmetic', () => {
  it('accepts a table whose every row adds up to the grid', () => {
    const [table] = readTableGrids(doc(tbl(gridOf(3), tr(tc(), tc(), tc()), tr(tc(), tc(), tc()))));

    expect(table.columnCount).toBe(3);
    expect(table.rows).toHaveLength(2);
    expect(table.rows.map(r => r.columnsCovered)).toEqual([3, 3]);
    expect(table.problems).toEqual([]);
  });

  it('counts an absent w:gridSpan as one column, not zero', () => {
    // A default of 0 would make every ordinary table look under-filled; a default of
    // "skip the cell" would make every table look over-filled.
    const [table] = readTableGrids(doc(tbl(gridOf(2), tr(tc(), tc()))));

    expect(table.rows[0].cells.map(c => c.gridSpan)).toEqual([1, 1]);
    expect(table.problems).toEqual([]);
  });

  it('sums w:gridSpan across the row rather than counting cells', () => {
    const [table] = readTableGrids(
      doc(tbl(gridOf(4), tr(tc(span(2)), tc(span(2))), tr(tc(), tc(), tc(), tc())))
    );

    expect(table.rows[0].cells.map(c => c.gridSpan)).toEqual([2, 2]);
    expect(table.rows[0].columnsCovered).toBe(4);
    expect(table.problems).toEqual([]);
  });

  it('flags a row that covers fewer columns than the grid declares', () => {
    const [table] = readTableGrids(doc(tbl(gridOf(3), tr(tc(), tc()))), 'word/document.xml');

    const problem = table.problems.find(p => p.code === 'table/row-span-mismatch');
    expect(problem?.subject).toMatchObject({ row: '1', covered: '2', declared: '3' });
    expect(problem?.severity).toBe('error');
    expect(problem?.silent).toBe(true);
    expect(problem?.part).toBe('word/document.xml');
    expect(problem?.remediation).toContain('Add 1 more grid column');
  });

  it('flags a row that covers more columns than the grid declares', () => {
    const [table] = readTableGrids(doc(tbl(gridOf(2), tr(tc(), tc(), tc()))));

    const problem = table.problems.find(p => p.code === 'table/row-span-mismatch');
    expect(problem?.subject).toMatchObject({ covered: '3', declared: '2' });
    expect(problem?.remediation).toContain('Remove 1 grid column');
  });

  it('flags only the row that does not add up, not the whole table', () => {
    const [table] = readTableGrids(
      doc(tbl(gridOf(3), tr(tc(), tc(), tc()), tr(tc(), tc()), tr(tc(), tc(), tc())))
    );

    const mismatches = table.problems.filter(p => p.code === 'table/row-span-mismatch');
    expect(mismatches).toHaveLength(1);
    expect(mismatches[0].subject?.row).toBe('2');
  });

  it('counts w:gridBefore and w:gridAfter towards the row total', () => {
    // An indented row is legal and common. Ignoring the trPr edge counts would report
    // this correct table as broken — the false positive this check is likeliest to make.
    const [table] = readTableGrids(
      doc(
        tbl(
          gridOf(4),
          `<w:tr><w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/></w:trPr>${tc()}${tc()}</w:tr>`
        )
      )
    );

    expect(table.rows[0].gridBefore).toBe(1);
    expect(table.rows[0].gridAfter).toBe(1);
    expect(table.rows[0].columnsCovered).toBe(4);
    expect(table.problems).toEqual([]);
  });

  it('names the edge counts when an indented row still does not add up', () => {
    const [table] = readTableGrids(
      doc(tbl(gridOf(5), `<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>${tc()}${tc()}</w:tr>`))
    );

    const problem = table.problems.find(p => p.code === 'table/row-span-mismatch');
    expect(problem?.subject?.covered).toBe('3');
    expect(problem?.message).toContain('w:gridBefore 1');
  });

  it('offsets later cells by w:gridBefore when locating them on the grid', () => {
    const [table] = readTableGrids(
      doc(tbl(gridOf(3), `<w:tr><w:trPr><w:gridBefore w:val="2"/></w:trPr>${tc()}</w:tr>`))
    );

    expect(table.rows[0].cells[0].startColumn).toBe(2);
  });
});

describe('readTableGrids — vertical merging and its inverted default', () => {
  const vRestart = '<w:vMerge w:val="restart"/>';
  const vBare = '<w:vMerge/>';
  const vContinue = '<w:vMerge w:val="continue"/>';

  it('reads a bare <w:vMerge/> as continue, not as restart', () => {
    // The trap: every other reading of this element produces an orphan or a phantom
    // merge. A bare vMerge under a restart is a healthy two-row merge.
    const [table] = readTableGrids(
      doc(tbl(gridOf(2), tr(tc(vRestart), tc()), tr(tc(vBare), tc())))
    );

    expect(table.rows[0].cells[0].vMerge).toBe('restart');
    expect(table.rows[1].cells[0].vMerge).toBe('continue');
    expect(table.verticalMerges).toEqual([
      { column: 0, gridSpan: 1, startRow: 0, endRow: 1, rowSpan: 2 }
    ]);
    expect(table.problems).toEqual([]);
  });

  it('treats an explicit continue exactly as a bare one', () => {
    const [table] = readTableGrids(
      doc(tbl(gridOf(2), tr(tc(vRestart), tc()), tr(tc(vContinue), tc())))
    );

    expect(table.verticalMerges[0].rowSpan).toBe(2);
    expect(table.problems).toEqual([]);
  });

  it('flags a continuation in the first row, which has nothing above it', () => {
    const [table] = readTableGrids(doc(tbl(gridOf(2), tr(tc(vBare), tc()))));

    const problem = table.problems.find(p => p.code === 'table/vmerge-orphan');
    expect(problem?.subject).toMatchObject({ row: '1', column: '1' });
    expect(problem?.message).toContain('first row');
    // The remediation must name the inverted default, because writing <w:vMerge/> when
    // you meant restart is the way this gets produced.
    expect(problem?.message).toContain('omitted w:val means "continue"');
    expect(problem?.severity).toBe('error');
    expect(problem?.silent).toBe(true);
  });

  it('flags a continuation whose column has no restart above it', () => {
    const [table] = readTableGrids(
      doc(tbl(gridOf(2), tr(tc(vRestart), tc()), tr(tc(vBare), tc(vBare))))
    );

    const orphans = table.problems.filter(p => p.code === 'table/vmerge-orphan');
    expect(orphans).toHaveLength(1);
    // Column 1 continues a real merge; column 2 has no restart anywhere above it.
    expect(orphans[0].subject).toMatchObject({ row: '2', column: '2' });
  });

  it('does not let a merge survive a row that breaks it', () => {
    // Row 2 has a plain cell in column 1, which ends the merge. The continuation in row
    // 3 is therefore an orphan, even though a restart does appear somewhere above it.
    const [table] = readTableGrids(
      doc(tbl(gridOf(2), tr(tc(vRestart), tc()), tr(tc(), tc()), tr(tc(vBare), tc())))
    );

    const orphans = table.problems.filter(p => p.code === 'table/vmerge-orphan');
    expect(orphans).toHaveLength(1);
    expect(orphans[0].subject?.row).toBe('3');
    expect(table.verticalMerges[0].rowSpan).toBe(1);
  });

  it('lets a restart in a later row replace the merge open at that column', () => {
    const [table] = readTableGrids(
      doc(
        tbl(
          gridOf(1),
          tr(tc(vRestart)),
          tr(tc(vBare)),
          tr(tc(vRestart)),
          tr(tc(vBare))
        )
      )
    );

    expect(table.problems).toEqual([]);
    expect(table.verticalMerges).toEqual([
      { column: 0, gridSpan: 1, startRow: 0, endRow: 1, rowSpan: 2 },
      { column: 0, gridSpan: 1, startRow: 2, endRow: 3, rowSpan: 2 }
    ]);
  });

  it('detaches a continuation that a changed gridSpan shifted off its column', () => {
    // Row 1 merges at column 1. Row 2's first cell spans two columns, so its second
    // cell starts at column 3 — not column 2 — and the merge it meant to continue is
    // somewhere else. Word absorbs this; the table still looks merged.
    const [table] = readTableGrids(
      doc(tbl(gridOf(3), tr(tc(), tc(vRestart), tc()), tr(tc(span(2)), tc(vBare))))
    );

    const orphan = table.problems.find(p => p.code === 'table/vmerge-orphan');
    expect(orphan?.subject).toMatchObject({ row: '2', column: '3' });
    expect(orphan?.remediation).toContain('same grid column');
  });

  it('flags a continuation whose gridSpan disagrees with its restart', () => {
    const [table] = readTableGrids(
      doc(tbl(gridOf(2), tr(tc(`${span(2)}${vRestart}`)), tr(tc(vBare), tc())))
    );

    const problem = table.problems.find(p => p.code === 'table/vmerge-span-mismatch');
    expect(problem?.message).toContain('spans 1 column(s)');
    expect(problem?.message).toContain('spans 2');
    expect(problem?.severity).toBe('warning');
  });

  it('notes a restart nothing continues, without calling it an error', () => {
    const [table] = readTableGrids(doc(tbl(gridOf(2), tr(tc(vRestart), tc()), tr(tc(), tc()))));

    const problem = table.problems.find(p => p.code === 'table/vmerge-restart-alone');
    expect(problem?.severity).toBe('note');
    expect(problem?.subject).toMatchObject({ row: '1', column: '1' });
    expect(codes(table.problems)).not.toContain('table/vmerge-orphan');
  });

  it('tracks two independent merges in different columns at once', () => {
    const [table] = readTableGrids(
      doc(
        tbl(
          gridOf(2),
          tr(tc(vRestart), tc(vRestart)),
          tr(tc(vBare), tc(vBare)),
          tr(tc(vBare), tc(vBare))
        )
      )
    );

    expect(table.problems).toEqual([]);
    expect(table.verticalMerges.map(m => [m.column, m.rowSpan])).toEqual([[0, 3], [1, 3]]);
  });

  it('rejects a w:val outside the two-member ST_Merge enumeration, and reads it as continue', () => {
    const [table] = readTableGrids(
      doc(tbl(gridOf(1), tr(tc('<w:vMerge w:val="start"/>'))))
    );

    const problem = table.problems.find(p => p.code === 'table/merge-value-invalid');
    expect(problem?.subject?.value).toBe('start');
    expect(problem?.message).toContain('only "restart" and "continue"');
    // Treated as continue, so it also orphans — the safe direction to fail in.
    expect(codes(table.problems)).toContain('table/vmerge-orphan');
  });
});

describe('readTableGrids — nested tables', () => {
  const inner = tbl(gridOf(3), tr(tc(), tc(), tc()), tr(tc(), tc(), tc()));
  const outerCellWithTable = `<w:tc><w:p/>${inner}<w:p/></w:tc>`;

  it('does not measure a nested table\'s rows against the outer grid', () => {
    // The trap. A descendant walk for w:tr collects the inner table's 3-cell rows and
    // measures them against the outer table's 2-column grid, inventing two mismatches
    // in a document that is entirely correct.
    const tables = readTableGrids(
      doc(tbl(gridOf(2), tr(outerCellWithTable, tc()), tr(tc(), tc())))
    );

    const outer = tables.find(t => t.depth === 0)!;
    expect(outer.columnCount).toBe(2);
    expect(outer.rows).toHaveLength(2);
    expect(outer.rows.map(r => r.cells.length)).toEqual([2, 2]);
    expect(outer.problems).toEqual([]);
  });

  it('does not measure a nested table\'s cells against the outer grid either', () => {
    const tables = readTableGrids(
      doc(tbl(gridOf(2), tr(outerCellWithTable, tc())))
    );

    const outer = tables.find(t => t.depth === 0)!;
    expect(outer.rows[0].columnsCovered).toBe(2);
  });

  it('still reports a nested table that is broken, against its own grid', () => {
    // Skipping nested tables entirely would be the lazy fix for the trap above, and it
    // would silently stop checking a whole class of tables.
    const brokenInner = tbl(gridOf(3), tr(tc(), tc()));
    const tables = readTableGrids(
      doc(tbl(gridOf(2), tr(`<w:tc><w:p/>${brokenInner}</w:tc>`, tc()))),
      'word/document.xml'
    );

    expect(tables).toHaveLength(2);
    const nestedTable = tables.find(t => t.depth === 1)!;
    expect(nestedTable.columnCount).toBe(3);
    expect(nestedTable.problems.find(p => p.code === 'table/row-span-mismatch')?.message).toContain(
      'nested table'
    );
    expect(tables.find(t => t.depth === 0)!.problems).toEqual([]);
  });

  it('does not let a nested table\'s vMerge restart satisfy the outer table', () => {
    // Column tracking that pooled rows across nesting levels would see the inner
    // restart and treat the outer continuation as healthy.
    const innerWithMerge = tbl(gridOf(1), tr(tc('<w:vMerge w:val="restart"/>')), tr(tc('<w:vMerge/>')));
    const tables = readTableGrids(
      doc(
        tbl(
          gridOf(1),
          tr(`<w:tc><w:p/>${innerWithMerge}</w:tc>`),
          tr(tc('<w:vMerge/>'))
        )
      )
    );

    const outer = tables.find(t => t.depth === 0)!;
    expect(outer.problems.find(p => p.code === 'table/vmerge-orphan')?.subject?.row).toBe('2');
    expect(tables.find(t => t.depth === 1)!.problems).toEqual([]);
  });

  it('reports depth for tables nested two levels down', () => {
    const deepest = tbl(gridOf(1), tr(tc()));
    const middle = tbl(gridOf(1), tr(`<w:tc><w:p/>${deepest}</w:tc>`));
    const tables = readTableGrids(doc(tbl(gridOf(1), tr(`<w:tc><w:p/>${middle}</w:tc>`))));

    expect(tables.map(t => t.depth)).toEqual([0, 1, 2]);
  });
});

describe('readTableGrids — revision shadows', () => {
  it('ignores the pre-revision grid inside w:tblGridChange', () => {
    // CT_TblGridChange contains a second w:tblGrid (CT_TblGridBase) holding the old
    // gridCol list. Counting descendants doubles the column count of any track-changed
    // grid and reports every row as short.
    const gridWithHistory =
      `<w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/>` +
      `<w:tblGridChange w:id="1"><w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/>` +
      `<w:gridCol w:w="1440"/></w:tblGrid></w:tblGridChange></w:tblGrid>`;
    const [table] = readTableGrids(doc(tbl(gridWithHistory, tr(tc(), tc()))));

    expect(table.columnCount).toBe(2);
    expect(table.problems).toEqual([]);
  });

  it('ignores the pre-revision gridSpan inside w:tcPrChange', () => {
    // CT_TcPr contains w:tcPrChange, which contains a second w:tcPr (CT_TcPrInner) with
    // the old gridSpan. Reading it as current makes this correct row cover 4 columns.
    const cellWithHistory =
      `<w:tc><w:tcPr><w:gridSpan w:val="2"/>` +
      `<w:tcPrChange w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z">` +
      `<w:tcPr><w:gridSpan w:val="3"/></w:tcPr></w:tcPrChange></w:tcPr><w:p/></w:tc>`;
    const [table] = readTableGrids(doc(tbl(gridOf(3), `<w:tr>${cellWithHistory}${tc()}</w:tr>`)));

    expect(table.rows[0].cells.map(c => c.gridSpan)).toEqual([2, 1]);
    expect(table.problems).toEqual([]);
  });

  it('ignores a pre-revision w:vMerge inside w:tcPrChange', () => {
    const cellWithHistory =
      `<w:tc><w:tcPr>` +
      `<w:tcPrChange w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z">` +
      `<w:tcPr><w:vMerge/></w:tcPr></w:tcPrChange></w:tcPr><w:p/></w:tc>`;
    const [table] = readTableGrids(doc(tbl(gridOf(1), `<w:tr>${cellWithHistory}</w:tr>`)));

    expect(table.rows[0].cells[0].vMerge).toBeNull();
    expect(table.problems).toEqual([]);
  });
});

describe('readTableGrids — transparent wrappers', () => {
  it('sees rows wrapped in a structured document tag', () => {
    // CT_SdtRow is a legal child of w:tbl and carries real rows. Treating it as opaque
    // loses them; treating it as a row loses the table.
    const [table] = readTableGrids(
      doc(
        tbl(
          gridOf(2),
          tr(tc(), tc()),
          `<w:sdt><w:sdtContent>${tr(tc(), tc())}</w:sdtContent></w:sdt>`
        )
      )
    );

    expect(table.rows).toHaveLength(2);
    expect(table.problems).toEqual([]);
  });

  it('sees cells wrapped in a structured document tag', () => {
    const [table] = readTableGrids(
      doc(tbl(gridOf(2), `<w:tr>${tc()}<w:sdt><w:sdtContent>${tc()}</w:sdtContent></w:sdt></w:tr>`))
    );

    expect(table.rows[0].cells).toHaveLength(2);
    expect(table.problems).toEqual([]);
  });

  it('sees rows wrapped in w:customXml', () => {
    const [table] = readTableGrids(
      doc(tbl(gridOf(1), `<w:customXml w:element="e">${tr(tc())}</w:customXml>`))
    );

    expect(table.rows).toHaveLength(1);
  });
});

describe('readTableGrids — horizontal merging', () => {
  it('flags an hMerge continuation with no restart to its left', () => {
    const [table] = readTableGrids(
      doc(tbl(gridOf(2), tr(tc(), tc('<w:hMerge/>'))))
    );

    const problem = table.problems.find(p => p.code === 'table/hmerge-orphan');
    expect(problem?.subject).toMatchObject({ row: '1', cell: '2' });
    expect(problem?.severity).toBe('error');
  });

  it('accepts an hMerge run that starts with a restart', () => {
    const [table] = readTableGrids(
      doc(tbl(gridOf(2), tr(tc('<w:hMerge w:val="restart"/>'), tc('<w:hMerge/>'))))
    );

    expect(codes(table.problems)).not.toContain('table/hmerge-orphan');
  });

  it('notes that hMerge is superseded, once per table rather than once per cell', () => {
    const [table] = readTableGrids(
      doc(
        tbl(
          gridOf(2),
          tr(tc('<w:hMerge w:val="restart"/>'), tc('<w:hMerge/>')),
          tr(tc('<w:hMerge w:val="restart"/>'), tc('<w:hMerge/>'))
        )
      )
    );

    const legacy = table.problems.filter(p => p.code === 'table/hmerge-legacy');
    expect(legacy).toHaveLength(1);
    expect(legacy[0].severity).toBe('note');
    expect(legacy[0].message).toContain('w:gridSpan');
  });

  it('does not mention hMerge for a table that uses gridSpan', () => {
    const [table] = readTableGrids(doc(tbl(gridOf(2), tr(tc(span(2))))));

    expect(codes(table.problems)).not.toContain('table/hmerge-legacy');
  });

  it('does not carry an open hMerge across a plain cell', () => {
    const [table] = readTableGrids(
      doc(tbl(gridOf(3), tr(tc('<w:hMerge w:val="restart"/>'), tc(), tc('<w:hMerge/>'))))
    );

    expect(table.problems.find(p => p.code === 'table/hmerge-orphan')?.subject?.cell).toBe('3');
  });
});

describe('readTableGrids — widths', () => {
  it('reads gridCol widths and totals them', () => {
    const [table] = readTableGrids(doc(tbl(grid(1000, 2000, 3000), tr(tc(), tc(), tc()))));

    expect(table.columnWidths).toEqual([1000, 2000, 3000]);
    expect(table.gridWidthTwips).toBe(6000);
  });

  it('flags a w:tblW in twips that contradicts the grid', () => {
    const [table] = readTableGrids(
      doc(
        tbl(
          `<w:tblPr><w:tblW w:w="9999" w:type="dxa"/></w:tblPr>`,
          grid(1000, 2000),
          tr(tc(), tc())
        )
      )
    );

    const problem = table.problems.find(p => p.code === 'table/width-contradicts-grid');
    expect(problem?.subject).toMatchObject({ declared: '9999', grid: '3000' });
    expect(problem?.severity).toBe('warning');
  });

  it('does not compare a percentage or auto table width against twips', () => {
    // pct and auto are measured against the page or the containing cell, which this
    // module cannot see. Comparing them would be a guess dressed up as a finding.
    for (const type of ['pct', 'auto', 'nil']) {
      const [table] = readTableGrids(
        doc(
          tbl(
            `<w:tblPr><w:tblW w:w="5000" w:type="${type}"/></w:tblPr>`,
            grid(1000, 2000),
            tr(tc(), tc())
          )
        )
      );

      expect(table.declaredWidthTwips).toBeNull();
      expect(table.declaredWidthType).toBe(type);
      expect(codes(table.problems)).not.toContain('table/width-contradicts-grid');
    }
  });

  it('tolerates rounding drift rather than flagging every table', () => {
    const [table] = readTableGrids(
      doc(
        tbl(`<w:tblPr><w:tblW w:w="3010" w:type="dxa"/></w:tblPr>`, grid(1000, 2000), tr(tc(), tc()))
      )
    );

    expect(codes(table.problems)).not.toContain('table/width-contradicts-grid');
  });

  it('flags cell widths that contradict the grid they fill', () => {
    const cell = (w: number) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`;
    const [table] = readTableGrids(
      doc(tbl(grid(1000, 2000), `<w:tr>${cell(1000)}${cell(5000)}</w:tr>`))
    );

    const problem = table.problems.find(p => p.code === 'table/cell-width-mismatch');
    expect(problem?.subject).toMatchObject({ cells: '6000', grid: '3000' });
  });

  it('does not compare cell widths for a row that does not fill the grid', () => {
    const cell = (w: number) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`;
    const [table] = readTableGrids(
      doc(
        tbl(
          grid(1000, 2000, 3000),
          `<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>${cell(2000)}${cell(3000)}</w:tr>`
        )
      )
    );

    expect(codes(table.problems)).not.toContain('table/cell-width-mismatch');
  });

  it('reports a unit-suffixed gridCol width as unknown rather than as a number', () => {
    // The 2010 union admits "1.5in". Parsing that as an integer would silently produce
    // 1, and a width total that is confidently wrong.
    const [table] = readTableGrids(
      doc(tbl(`<w:tblGrid><w:gridCol w:w="1.5in"/><w:gridCol w:w="2880"/></w:tblGrid>`, tr(tc(), tc())))
    );

    expect(table.columnWidths).toEqual([null, 2880]);
    expect(table.gridWidthTwips).toBeNull();
    expect(codes(table.problems)).not.toContain('table/width-contradicts-grid');
  });
});

describe('readTableGrids — malformed and missing values', () => {
  it('flags a w:gridSpan with no w:val, which the schema requires', () => {
    const [table] = readTableGrids(doc(tbl(gridOf(2), tr(tc('<w:gridSpan/>'), tc()))));

    const problem = table.problems.find(p => p.code === 'table/span-value-invalid');
    expect(problem?.message).toContain('no w:val');
    expect(problem?.subject).toMatchObject({ row: '1', cell: '1' });
  });

  it('flags a non-numeric w:gridSpan and falls back to one column', () => {
    const [table] = readTableGrids(doc(tbl(gridOf(2), tr(tc('<w:gridSpan w:val="two"/>'), tc()))));

    expect(table.problems.find(p => p.code === 'table/span-value-invalid')?.message).toContain('"two"');
    expect(table.rows[0].columnsCovered).toBe(2);
  });

  it('flags a zero or negative w:gridSpan', () => {
    const [table] = readTableGrids(doc(tbl(gridOf(1), tr(tc('<w:gridSpan w:val="0"/>')))));

    expect(codes(table.problems)).toContain('table/span-value-invalid');
  });

  it('flags an unusable w:gridBefore against the row that carries it', () => {
    const [table] = readTableGrids(
      doc(tbl(gridOf(2), tr(tc(), tc()), `<w:tr><w:trPr><w:gridBefore/></w:trPr>${tc()}${tc()}</w:tr>`))
    );

    const problem = table.problems.find(p => p.code === 'table/span-value-invalid');
    expect(problem?.subject).toMatchObject({ row: '2', element: 'gridBefore' });
  });

  it('flags a table with rows but no grid at all', () => {
    const [table] = readTableGrids(doc(tbl(tr(tc(), tc()), tr(tc(), tc()))));

    expect(table.columnCount).toBe(0);
    const problem = table.problems.find(p => p.code === 'table/grid-missing');
    expect(problem?.severity).toBe('warning');
    expect(problem?.subject?.rows).toBe('2');
  });

  it('compares rows to one another when there is no grid to compare them to', () => {
    const [table] = readTableGrids(doc(tbl(tr(tc(), tc()), tr(tc(), tc()), tr(tc(), tc(), tc()))));

    const ragged = table.problems.filter(p => p.code === 'table/ragged-rows');
    expect(ragged).toHaveLength(1);
    expect(ragged[0].subject).toMatchObject({ row: '3', covered: '3', typical: '2' });
  });

  it('does not report ragged rows when a grid already settles the question', () => {
    // Reporting both would double-count: every ragged row is already a row-span
    // mismatch when a grid exists.
    const [table] = readTableGrids(doc(tbl(gridOf(2), tr(tc(), tc()), tr(tc(), tc(), tc()))));

    expect(codes(table.problems)).not.toContain('table/ragged-rows');
    expect(codes(table.problems)).toContain('table/row-span-mismatch');
  });

  it('returns nothing for a document with no tables', () => {
    expect(readTableGrids(doc('<w:p><w:r><w:t>plain</w:t></w:r></w:p>'))).toEqual([]);
  });

  it('does not throw on a table with a grid but no rows', () => {
    const [table] = readTableGrids(doc(tbl(gridOf(3))));

    expect(table.rows).toEqual([]);
    expect(table.problems).toEqual([]);
  });

  it('does not throw on a row with no cells', () => {
    const [table] = readTableGrids(doc(tbl(gridOf(2), '<w:tr/>')));

    expect(table.rows[0].cells).toEqual([]);
    expect(codes(table.problems)).toContain('table/row-span-mismatch');
  });

  it('ignores elements in another namespace that share these local names', () => {
    const alien =
      `<w:tbl><x:tblGrid xmlns:x="urn:not-word"><x:gridCol w:w="2880"/></x:tblGrid>` +
      `${gridOf(2)}${tr(tc(), tc())}</w:tbl>`;
    const [table] = readTableGrids(doc(alien));

    expect(table.columnCount).toBe(2);
    expect(table.problems).toEqual([]);
  });
});

describe('tableGridFindings', () => {
  it('flattens every table\'s problems into one list with the part attached', () => {
    const findings = tableGridFindings(
      doc(tbl(gridOf(2), tr(tc())) + tbl(gridOf(3), tr(tc()))),
      'word/header1.xml'
    );

    expect(findings).toHaveLength(2);
    expect(findings.every(f => f.part === 'word/header1.xml')).toBe(true);
    expect(findings[0].message).toContain('table 1');
    expect(findings[1].message).toContain('table 2');
  });

  it('returns an empty list for a healthy document', () => {
    expect(tableGridFindings(doc(tbl(gridOf(2), tr(tc(), tc()))), 'word/document.xml')).toEqual([]);
  });
});

describe('computeTableEvidenceForMarkup — panel wiring', () => {
  const part = (body: string) => ({
    'word/document.xml': `<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`
  });

  it('returns null when the part has no tables, so the panel degrades quietly', () => {
    expect(computeTableEvidenceForMarkup(part('<w:p/>'), '')).toBeNull();
  });

  it('returns null when no Word body part is present', () => {
    expect(computeTableEvidenceForMarkup({ 'word/styles.xml': '<w:styles/>' }, '')).toBeNull();
  });

  it('returns null rather than throwing on malformed XML', () => {
    expect(computeTableEvidenceForMarkup({ 'word/document.xml': '<w:document><oops>' }, '')).toBeNull();
  });

  it('counts nested tables separately and says why that matters', () => {
    const nested = tbl(gridOf(1), tr(tc()));
    const evidence = computeTableEvidenceForMarkup(
      part(tbl(gridOf(2), tr(`<w:tc><w:p/>${nested}</w:tc>`, tc()))),
      ''
    );

    expect(evidence!.lines[0]).toContain('2 table(s), 1 of them nested');
    expect(evidence!.lines[0]).toContain('do not count towards the grid');
  });

  it('describes the grid and row count of each table', () => {
    const evidence = computeTableEvidenceForMarkup(
      part(tbl(grid(1000, 2000), tr(tc(), tc()))),
      ''
    );

    expect(evidence!.lines.some(l => l.includes('declares 2 grid column(s) (3000 twips wide in total) and has 1 row(s)'))).toBe(true);
  });

  it('reports the span of a vertical merge in plain terms', () => {
    const evidence = computeTableEvidenceForMarkup(
      part(
        tbl(
          gridOf(2),
          tr(tc('<w:vMerge w:val="restart"/>'), tc()),
          tr(tc('<w:vMerge/>'), tc()),
          tr(tc('<w:vMerge/>'), tc())
        )
      ),
      ''
    );

    const line = evidence!.lines.find(l => l.includes('vertical merge'))!;
    expect(line).toContain('grid column 1 spans rows 1–3');
    expect(line).toContain('omitted w:val means continue');
  });

  it('surfaces a row that does not add up', () => {
    const evidence = computeTableEvidenceForMarkup(part(tbl(gridOf(3), tr(tc(), tc()))), '');

    expect(evidence!.lines.some(l => l.includes('covers 2 grid column(s)') && l.includes('declares 3'))).toBe(true);
  });

  it('narrows to the selected table when its shape identifies it uniquely', () => {
    const selected = tbl(gridOf(3), tr(tc(), tc()));
    const other = tbl(gridOf(2), tr(tc(), tc()));
    const evidence = computeTableEvidenceForMarkup(
      part(other + selected),
      `<w:tbl ${W}>${gridOf(3)}${tr(tc(), tc())}</w:tbl>`
    );

    // Table 2 is the broken one and the selected one; only it should be described.
    expect(evidence!.lines.some(l => l.startsWith('Table 2 declares 3 grid column(s)'))).toBe(true);
    expect(evidence!.lines.some(l => l.startsWith('Table 1 '))).toBe(false);
  });

  it('admits it cannot tell two identically-shaped tables apart', () => {
    const shape = tbl(gridOf(2), tr(tc(), tc()));
    const evidence = computeTableEvidenceForMarkup(
      part(shape + shape),
      `<w:tbl ${W}>${gridOf(2)}${tr(tc(), tc())}</w:tbl>`
    );

    expect(evidence!.unresolved.some(u => u.includes('2 tables in word/document.xml have the same shape'))).toBe(true);
    // Falls back to describing all of them rather than picking one.
    expect(evidence!.lines.filter(l => /^Table \d declares/.test(l))).toHaveLength(2);
  });

  it('describes every table when nothing is selected', () => {
    const evidence = computeTableEvidenceForMarkup(
      part(tbl(gridOf(1), tr(tc())) + tbl(gridOf(2), tr(tc(), tc()))),
      ''
    );

    expect(evidence!.lines.filter(l => /^Table \d declares/.test(l))).toHaveLength(2);
  });

  it('caps what it claims about widths it cannot compare', () => {
    const evidence = computeTableEvidenceForMarkup(
      part(tbl(`<w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>`, gridOf(2), tr(tc(), tc()))),
      ''
    );

    expect(evidence!.unresolved.some(u => u.includes('w:type="pct"'))).toBe(true);
  });

  it('always caps the claim about table styles, which it never reads', () => {
    const evidence = computeTableEvidenceForMarkup(part(tbl(gridOf(1), tr(tc()))), '');

    expect(evidence!.unresolved.some(u => u.includes('table style'))).toBe(true);
  });

  it('says so when a table has no grid at all', () => {
    const evidence = computeTableEvidenceForMarkup(part(tbl(tr(tc(), tc()))), '');

    expect(evidence!.lines.some(l => l.includes('no <w:tblGrid> at all'))).toBe(true);
  });
});
