import { describe, it, expect } from 'vitest';
import {
  readFormulas,
  readCalcSettings,
  formulaFindingsForSheet,
  formulaFindings,
  computeFormulaEvidenceForMarkup
} from '../services/excelFormulas';
import type { PackageParts } from '../services/packageIntegrity';

const S = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

const sheet = (cells: string) =>
  `<?xml version="1.0"?><worksheet ${S}><sheetData><row r="1">${cells}</row></sheetData></worksheet>`;

const workbook = (calcPr = '') =>
  `<?xml version="1.0"?><workbook ${S}><sheets><sheet name="S1" sheetId="1"/></sheets>${calcPr}</workbook>`;

const pkg = (cells: string, calcPr = ''): PackageParts => ({
  'xl/workbook.xml': workbook(calcPr),
  'xl/worksheets/sheet1.xml': sheet(cells)
});

describe('reading formulas', () => {
  it('reads the formula and its cached value as two separate things', () => {
    const [f] = readFormulas(sheet(`<c r="B2" t="n"><f>SUM(A1:A10)</f><v>55</v></c>`));

    expect(f.text).toBe('SUM(A1:A10)');
    expect(f.cachedValue).toBe('55');
    expect(f.kind).toBe('normal');
  });

  it('treats an absent @t as normal, which is the schema default', () => {
    expect(readFormulas(sheet(`<c r="A1"><f>1+1</f><v>2</v></c>`))[0].kind).toBe('normal');
  });

  it('recognises every formula kind the schema declares', () => {
    const cells =
      `<c r="A1"><f t="array" ref="A1:A2">X</f><v>1</v></c>` +
      `<c r="B1"><f t="shared" ref="B1:B2" si="0">Y</f><v>2</v></c>` +
      `<c r="C1"><f t="dataTable" ref="C1">Z</f><v>3</v></c>`;

    expect(readFormulas(sheet(cells)).map(f => f.kind)).toEqual(['array', 'shared', 'dataTable']);
  });

  it('ignores cells that hold a value but no formula', () => {
    expect(readFormulas(sheet(`<c r="A1" t="n"><v>42</v></c>`))).toEqual([]);
  });

  it('records a shared follower as having no formula text, which is correct not broken', () => {
    const [, follower] = readFormulas(
      sheet(`<c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c>`)
    );

    expect(follower.text).toBe('');
    expect(follower.range).toBeNull();
    expect(follower.sharedIndex).toBe('0');
  });

  it('returns nothing rather than throwing on malformed XML', () => {
    expect(readFormulas('<worksheet><unclosed>')).toEqual([]);
  });
});

describe('the shared formula whose master is gone', () => {
  it('reports a follower with no master, because the formula is genuinely lost', () => {
    // A follower stores no formula text - its formula exists only as an offset from
    // the master. Lose the master and the cell has a number and no way to recompute it.
    const problems = formulaFindingsForSheet(
      sheet(`<c r="B2"><f t="shared" si="0"/><v>4</v></c>`),
      'xl/worksheets/sheet1.xml',
      null
    );

    const problem = problems.find(p => p.code === 'formula/shared-master-missing');
    expect(problem?.silent).toBe(true);
    expect(problem?.subject?.cell).toBe('B2');
  });

  it('says nothing when the master is present', () => {
    const problems = formulaFindingsForSheet(
      sheet(`<c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c>`),
      'xl/worksheets/sheet1.xml',
      null
    );

    expect(problems.map(p => p.code)).not.toContain('formula/shared-master-missing');
  });

  it('matches master to follower on si, not on position', () => {
    // Two shared groups. The follower for si=1 has a master; the one for si=2 does not.
    // A check that just looked for "any master in the sheet" would report neither.
    const cells =
      `<c r="B1"><f t="shared" ref="B1:B2" si="1">A1*2</f><v>2</v></c>` +
      `<c r="B2"><f t="shared" si="1"/><v>4</v></c>` +
      `<c r="C2"><f t="shared" si="2"/><v>9</v></c>`;
    const problems = formulaFindingsForSheet(sheet(cells), 'xl/worksheets/sheet1.xml', null);

    const orphans = problems.filter(p => p.code === 'formula/shared-master-missing');
    expect(orphans).toHaveLength(1);
    expect(orphans[0].subject?.cell).toBe('C2');
  });

  it('does not mistake a master for a follower', () => {
    // A master has both ref and si. Testing only for si would flag every master.
    const problems = formulaFindingsForSheet(
      sheet(`<c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c>`),
      'xl/worksheets/sheet1.xml',
      null
    );

    expect(problems.map(p => p.code)).not.toContain('formula/shared-master-missing');
  });
});

describe('what the workbook says about its own numbers', () => {
  it('reads the calculation settings', () => {
    const calc = readCalcSettings({
      'xl/workbook.xml': workbook('<calcPr calcId="191029" calcMode="manual" fullCalcOnLoad="1" calcCompleted="0"/>')
    });

    expect(calc).toEqual({ fullCalcOnLoad: true, calcMode: 'manual', calcCompleted: false, calcId: '191029' });
  });

  it('reports fullCalcOnLoad as Excel disowning its own stored values', () => {
    const problems = formulaFindings(pkg(`<c r="A1"><f>1+1</f><v>2</v></c>`, '<calcPr fullCalcOnLoad="1"/>'));

    const problem = problems.find(p => p.code === 'formula/stale-by-declaration');
    expect(problem?.message).toContain('already disowned');
    expect(problem?.silent).toBe(true);
  });

  it('reports manual calculation, where values drift by design', () => {
    const problems = formulaFindings(pkg(`<c r="A1"><f>1+1</f><v>2</v></c>`, '<calcPr calcMode="manual"/>'));

    expect(problems.map(p => p.code)).toContain('formula/manual-calculation');
  });

  it('says nothing about calculation when the settings are ordinary', () => {
    const problems = formulaFindings(pkg(`<c r="A1"><f>1+1</f><v>2</v></c>`, '<calcPr calcId="191029"/>'));
    const codes = problems.map(p => p.code);

    expect(codes).not.toContain('formula/stale-by-declaration');
    expect(codes).not.toContain('formula/manual-calculation');
    expect(codes).not.toContain('formula/calculation-incomplete');
  });

  it('reports workbook settings once, not once per worksheet', () => {
    // Otherwise a ten-sheet workbook reports the same fact ten times and buries the
    // per-cell findings that are actually actionable.
    const parts: PackageParts = {
      'xl/workbook.xml': workbook('<calcPr fullCalcOnLoad="1"/>'),
      'xl/worksheets/sheet1.xml': sheet(`<c r="A1"><f>1+1</f><v>2</v></c>`),
      'xl/worksheets/sheet2.xml': sheet(`<c r="A1"><f>2+2</f><v>4</v></c>`),
      'xl/worksheets/sheet3.xml': sheet(`<c r="A1"><f>3+3</f><v>6</v></c>`)
    };

    expect(formulaFindings(parts).filter(p => p.code === 'formula/stale-by-declaration')).toHaveLength(1);
  });

  it('copes with a workbook that declares no calcPr at all', () => {
    expect(readCalcSettings({ 'xl/workbook.xml': workbook() })).toEqual({
      fullCalcOnLoad: false,
      calcMode: null,
      calcCompleted: null,
      calcId: null
    });
  });
});

describe('per-cell findings', () => {
  const check = (cells: string) => formulaFindingsForSheet(sheet(cells), 'xl/worksheets/sheet1.xml', null);

  it('reports a stored error value, and marks it visible unlike the rest', () => {
    const problem = check(`<c r="A1" t="e"><f>B1/0</f><v>#DIV/0!</v></c>`).find(
      p => p.code === 'formula/cached-error-value'
    );

    expect(problem?.silent).toBe(false);
    expect(problem?.subject?.error).toBe('#DIV/0!');
  });

  it('explains that a #REF! has lost the original reference', () => {
    const problem = check(`<c r="A1" t="e"><f>#REF!*2</f><v>#REF!</v></c>`).find(
      p => p.code === 'formula/cached-error-value'
    );

    expect(problem?.remediation).toContain('not recoverable');
  });

  it('does not treat an ordinary string value as an error', () => {
    // t="str" with text that merely looks unusual must not trip the error check.
    expect(check(`<c r="A1" t="str"><f>A2</f><v>#NotAnError</v></c>`).map(p => p.code)).not.toContain(
      'formula/cached-error-value'
    );
  });

  it('does not flag a STRING whose text happens to be an error literal', () => {
    // IFERROR(x,"#N/A") returns the characters #N/A as a string, and t="str" says so.
    // Matching on the value alone reports a working formula as broken - and this is the
    // case that makes the cell-type guard load-bearing rather than decorative.
    const codes = check(`<c r="A1" t="str"><f>IFERROR(B1,"#N/A")</f><v>#N/A</v></c>`).map(p => p.code);

    expect(codes).not.toContain('formula/cached-error-value');
  });

  it('reports a formula with no cached value, which renders as an empty cell', () => {
    expect(check(`<c r="A1"><f>SUM(B:B)</f></c>`).map(p => p.code)).toContain('formula/formula-without-value');
  });

  it('does not report a missing value for a shared follower', () => {
    // Followers are handled by the shared-master check; reporting them here as well
    // would double-count every filled-down column.
    const codes = check(`<c r="B2"><f t="shared" si="0"/></c>`).map(p => p.code);

    expect(codes).not.toContain('formula/formula-without-value');
  });

  it('flags a volatile function whose cached value is stale on principle', () => {
    const problem = check(`<c r="A1"><f>TODAY()</f><v>45000</v></c>`).find(p => p.code === 'formula/volatile-formula');

    expect(problem?.subject?.function).toBe('TODAY');
  });

  it('does not flag a function that merely contains a volatile name', () => {
    // NOWHERE() and RANDOMISE() are not NOW() and RAND(). Matching on the bare name
    // rather than name-plus-paren would report both.
    expect(check(`<c r="A1"><f>NOWHERE(A2)+RANDOMISE(B2)</f><v>1</v></c>`).map(p => p.code)).not.toContain(
      'formula/volatile-formula'
    );
  });

  it('flags a reference to another workbook as uncheckable', () => {
    const problem = check(`<c r="A1"><f>[1]Sheet1!A1*2</f><v>10</v></c>`).find(
      p => p.code === 'formula/external-reference'
    );

    expect(problem?.message).toContain('outside this package');
  });

  it('reports an array formula with no declared range', () => {
    expect(check(`<c r="A1"><f t="array">SUM(B:B)</f><v>1</v></c>`).map(p => p.code)).toContain(
      'formula/array-master-missing-ref'
    );
  });

  it('says nothing about a healthy sheet', () => {
    expect(check(`<c r="A1" t="n"><f>SUM(B1:B9)</f><v>45</v></c>`)).toEqual([]);
  });
});

describe('computeFormulaEvidenceForMarkup', () => {
  it('returns null for a sheet with no formulas', () => {
    expect(computeFormulaEvidenceForMarkup(pkg(`<c r="A1" t="n"><v>42</v></c>`))).toBeNull();
  });

  it('leads with the fact that the displayed number is stored, not computed', () => {
    const evidence = computeFormulaEvidenceForMarkup(pkg(`<c r="A1"><f>1+1</f><v>2</v></c>`));

    expect(evidence!.lines[0]).toContain('last recalculation');
  });

  it('refuses to claim a cached value is numerically wrong', () => {
    // The honest limit. Saying "this may be stale" is supportable; saying "this is 55
    // but should be 60" would need a calculation engine this does not have.
    const evidence = computeFormulaEvidenceForMarkup(pkg(`<c r="A1"><f>1+1</f><v>2</v></c>`));

    expect(evidence!.unresolved.some(u => u.includes('No formula was evaluated'))).toBe(true);
  });

  it('adds a second limit when the sheet reaches outside the package', () => {
    const evidence = computeFormulaEvidenceForMarkup(pkg(`<c r="A1"><f>[1]Sheet1!A1</f><v>7</v></c>`));

    expect(evidence!.unresolved.some(u => u.includes('external workbook'))).toBe(true);
  });

  it('reports the workbook calculation mode alongside the formulas', () => {
    const evidence = computeFormulaEvidenceForMarkup(pkg(`<c r="A1"><f>1+1</f><v>2</v></c>`, '<calcPr calcMode="manual"/>'));

    expect(evidence!.lines.some(l => l.includes('manual'))).toBe(true);
  });
});
