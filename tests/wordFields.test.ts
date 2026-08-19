import { describe, it, expect } from 'vitest';
import {
  readFields,
  parseInstruction,
  crossCheckFieldTargets,
  computeFieldEvidenceForMarkup
} from '../services/wordFields';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const doc = (body: string) =>
  new DOMParser().parseFromString(
    `<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`,
    'application/xml'
  );

const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const fldChar = (type: string, attrs = '') => `<w:r><w:fldChar w:fldCharType="${type}"${attrs}/></w:r>`;
const instr = (text: string) => `<w:r><w:instrText>${text}</w:instrText></w:r>`;

/** A complete complex field: begin, instruction, separate, result, end. */
const complex = (instruction: string, result: string, beginAttrs = '') =>
  fldChar('begin', beginAttrs) + instr(instruction) + fldChar('separate') + run(result) + fldChar('end');

const bookmark = (name: string, id = '1') =>
  `<w:bookmarkStart w:id="${id}" w:name="${name}"/>${run('anchor')}<w:bookmarkEnd w:id="${id}"/>`;

describe('parseInstruction', () => {
  it('separates the type, arguments and switches', () => {
    const parsed = parseInstruction(' REF Chapter1 \\h \\* MERGEFORMAT ');

    expect(parsed.type).toBe('REF');
    expect(parsed.arguments).toContain('Chapter1');
    expect(parsed.switches).toEqual(expect.arrayContaining(['\\h', '\\*']));
  });

  it('keeps a quoted argument together', () => {
    // TOC \o "1-3" - splitting on whitespace turns the range into two arguments and
    // loses it entirely.
    const parsed = parseInstruction(' TOC \\o "1-3" \\h ');

    expect(parsed.arguments).toContain('1-3');
  });

  it('uppercases the type, since field names are case-insensitive', () => {
    expect(parseInstruction(' ref Ch1 ').type).toBe('REF');
  });

  it('reports an empty instruction as having no type', () => {
    expect(parseInstruction('   ').type).toBeNull();
  });
});

describe('reading fields', () => {
  it('reads a simple field and treats its content as the cached result', () => {
    const index = readFields(doc(`<w:p><w:fldSimple w:instr=" REF Ch1 \\h ">${run('Chapter One')}</w:fldSimple></w:p>`));

    expect(index.fields).toHaveLength(1);
    expect(index.fields[0].kind).toBe('simple');
    expect(index.fields[0].cachedResult).toBe('Chapter One');
  });

  it('assembles a complex field from three unrelated runs', () => {
    const index = readFields(doc(`<w:p>${complex(' REF Ch1 \\h ', 'Chapter One')}</w:p>`));

    expect(index.fields).toHaveLength(1);
    expect(index.fields[0].instruction.trim()).toBe('REF Ch1 \\h');
    expect(index.fields[0].cachedResult).toBe('Chapter One');
  });

  it('distinguishes a never-calculated field from one showing nothing', () => {
    // No separate marker at all means the field has never run. That is different from
    // a field that ran and produced empty text, and collapsing them hides which.
    const never = readFields(doc(`<w:p>${fldChar('begin')}${instr(' PAGE ')}${fldChar('end')}</w:p>`));
    const empty = readFields(doc(`<w:p>${complex(' PAGE ', '')}</w:p>`));

    expect(never.fields[0].cachedResult).toBeNull();
    expect(empty.fields[0].cachedResult).toBe('');
  });

  it('does not treat a missing separate as a fault', () => {
    // A field that has never been calculated is legal and common. Reporting it as
    // broken would put noise on clean documents.
    const index = readFields(doc(`<w:p>${fldChar('begin')}${instr(' PAGE ')}${fldChar('end')}</w:p>`));

    expect(index.problems.map(p => p.code)).not.toContain('field/unbalanced-begin');
  });

  it('reads dirty and locked flags off the begin marker', () => {
    const index = readFields(doc(`<w:p>${complex(' REF Ch1 ', 'x', ' w:dirty="true" w:fldLock="true"')}</w:p>`));

    expect(index.fields[0].dirty).toBe(true);
    expect(index.fields[0].locked).toBe(true);
  });

  it('stops absorbing instruction text once the separate marker is passed', () => {
    // A stray instrText inside a field's RESULT must not be appended to that field's
    // instruction. Every well-formed fixture puts instrText before the separate, so
    // without this case the guard could be deleted and nothing would notice.
    const malformed =
      fldChar('begin') + instr(' REF Ch1 ') + fldChar('separate') + run('Chapter One') +
      instr(' PAGEREF Stray ') + fldChar('end');
    const index = readFields(doc(`<w:p>${malformed}</w:p>`));

    expect(index.fields[0].instruction.trim()).toBe('REF Ch1');
    expect(index.fields[0].parsed.arguments).not.toContain('Stray');
  });

  it('does not mistake instruction text for the field result', () => {
    const index = readFields(doc(`<w:p>${complex(' REF Ch1 \\h ', 'Chapter One')}</w:p>`));

    expect(index.fields[0].cachedResult).not.toContain('REF');
  });
});

describe('nesting', () => {
  it('pairs the right begin with the right end', () => {
    // A TOC carries a PAGEREF per entry inside its own result. Pairing the first begin
    // with the first end attributes the outer instruction to the inner field and
    // mis-reads everything after it.
    const inner = complex(' PAGEREF _Toc1 \\h ', '7');
    const outer =
      fldChar('begin') + instr(' TOC \\o "1-3" \\h ') + fldChar('separate') + run('Introduction\t') + inner + fldChar('end');
    const index = readFields(doc(`<w:p>${outer}</w:p>`));

    const types = index.fields.map(f => f.parsed.type);
    expect(types).toContain('TOC');
    expect(types).toContain('PAGEREF');
    expect(index.problems).toEqual([]);
  });

  it('records depth, so a nested field is distinguishable from a top-level one', () => {
    const inner = complex(' PAGEREF _Toc1 \\h ', '7');
    const outer = fldChar('begin') + instr(' TOC ') + fldChar('separate') + inner + fldChar('end');
    const index = readFields(doc(`<w:p>${outer}</w:p>`));

    expect(index.fields.find(f => f.parsed.type === 'PAGEREF')!.depth).toBe(1);
    expect(index.fields.find(f => f.parsed.type === 'TOC')!.depth).toBe(0);
  });

  it("includes a nested field's result in its parent's, which is what a reader sees", () => {
    const inner = complex(' PAGEREF _Toc1 \\h ', '7');
    const outer =
      fldChar('begin') + instr(' TOC ') + fldChar('separate') + run('Introduction ') + inner + fldChar('end');
    const index = readFields(doc(`<w:p>${outer}</w:p>`));

    expect(index.fields.find(f => f.parsed.type === 'TOC')!.cachedResult).toBe('Introduction 7');
  });
});

describe('broken field structure', () => {
  it('reports a field that opens and never closes', () => {
    const index = readFields(doc(`<w:p>${fldChar('begin')}${instr(' REF Ch1 ')}${run('text')}</w:p>`));

    const problem = index.problems.find(p => p.code === 'field/unbalanced-begin');
    expect(problem?.silent).toBe(true);
    expect(problem?.subject?.instruction).toContain('REF');
  });

  it('reports an end with no begin', () => {
    const index = readFields(doc(`<w:p>${run('text')}${fldChar('end')}</w:p>`));

    expect(index.problems.map(p => p.code)).toContain('field/unbalanced-end');
  });

  it('reports a separate outside any field', () => {
    const index = readFields(doc(`<w:p>${run('text')}${fldChar('separate')}</w:p>`));

    expect(index.problems.map(p => p.code)).toContain('field/orphan-separate');
  });

  it('flags an empty instruction as visible, unlike everything else here', () => {
    // Word prints an error placeholder for this one, so a reader does see it.
    const index = readFields(doc(`<w:p><w:fldSimple w:instr="  ">${run('x')}</w:fldSimple></w:p>`));

    expect(index.problems.find(p => p.code === 'field/empty-instruction')?.silent).toBe(false);
  });
});

describe('the stale cross-reference — the reason this module exists', () => {
  const withMissingTarget = (extra = '') =>
    doc(`<w:p>${bookmark('Chapter2')}</w:p><w:p><w:fldSimple w:instr=" REF Chapter1 \\h "${extra}>${run('The Beginning')}</w:fldSimple></w:p>`);

  it('catches a REF pointing at a bookmark that no longer exists', () => {
    const [problem] = crossCheckFieldTargets(withMissingTarget());

    expect(problem.code).toBe('field/stale-reference');
    expect(problem.subject?.target).toBe('Chapter1');
  });

  it('is silent: the document reads correctly and the reference is dead', () => {
    const [problem] = crossCheckFieldTargets(withMissingTarget());

    expect(problem.silent).toBe(true);
    expect(problem.message).toContain('The Beginning');
    expect(problem.message).toContain('stale');
  });

  it('says the cached text will NOT be recalculated when the field is not dirty', () => {
    const [problem] = crossCheckFieldTargets(withMissingTarget());

    expect(problem.message).toContain('not marked dirty');
  });

  it('escalates a locked field, because F9 will not fix it', () => {
    const [problem] = crossCheckFieldTargets(withMissingTarget(' w:fldLock="true"'));

    expect(problem.code).toBe('field/locked-stale-reference');
    expect(problem.message).toContain('will NOT fix it');
  });

  it('says nothing when the bookmark is there', () => {
    const healthy = doc(
      `<w:p>${bookmark('Chapter1')}</w:p><w:p><w:fldSimple w:instr=" REF Chapter1 \\h ">${run('The Beginning')}</w:fldSimple></w:p>`
    );

    expect(crossCheckFieldTargets(healthy)).toEqual([]);
  });

  it('checks PAGEREF and NOTEREF too', () => {
    const parts = doc(
      `<w:p><w:fldSimple w:instr=" PAGEREF Gone \\h ">${run('7')}</w:fldSimple></w:p>` +
        `<w:p><w:fldSimple w:instr=" NOTEREF Gone2 ">${run('3')}</w:fldSimple></w:p>`
    );

    expect(crossCheckFieldTargets(parts)).toHaveLength(2);
  });

  it('does not treat a HYPERLINK URL as a missing bookmark', () => {
    // Without \\l the argument is a URL. Reporting it as a dead bookmark would be
    // worse than saying nothing, because it is noise on every document with a link.
    const parts = doc(`<w:p><w:fldSimple w:instr=" HYPERLINK &quot;https://example.com&quot; ">${run('site')}</w:fldSimple></w:p>`);

    expect(crossCheckFieldTargets(parts)).toEqual([]);
  });

  it('does check a HYPERLINK that targets an anchor with \\l', () => {
    const parts = doc(`<w:p><w:fldSimple w:instr=" HYPERLINK \\l Gone ">${run('jump')}</w:fldSimple></w:p>`);

    expect(crossCheckFieldTargets(parts).map(p => p.subject?.target)).toEqual(['Gone']);
  });

  it('does not treat a STYLEREF argument as a bookmark', () => {
    // Its argument is a style name. This is the kind of false positive that makes a
    // report unreadable on a real document.
    const parts = doc(`<w:p><w:fldSimple w:instr=" STYLEREF &quot;Heading 1&quot; ">${run('Intro')}</w:fldSimple></w:p>`);

    expect(crossCheckFieldTargets(parts)).toEqual([]);
  });

  it('works for complex fields as well as simple ones', () => {
    const parts = doc(`<w:p>${complex(' REF Gone \\h ', 'old text')}</w:p>`);

    expect(crossCheckFieldTargets(parts).map(p => p.code)).toEqual(['field/stale-reference']);
  });
});

describe('computeFieldEvidenceForMarkup', () => {
  const part = (body: string) => ({
    'word/document.xml': `<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`
  });

  it('returns null when the part has no fields', () => {
    expect(computeFieldEvidenceForMarkup(part(`<w:p>${run('plain')}</w:p>`), '')).toBeNull();
  });

  it('leads with the fact that the displayed text is a cache', () => {
    const evidence = computeFieldEvidenceForMarkup(
      part(`<w:p><w:fldSimple w:instr=" REF Ch1 ">${run('One')}</w:fldSimple></w:p>`),
      ''
    );

    expect(evidence!.lines[0]).toContain('last time it was calculated');
  });

  it('surfaces the stale reference', () => {
    const evidence = computeFieldEvidenceForMarkup(
      part(`<w:p><w:fldSimple w:instr=" REF Gone ">${run('Old Title')}</w:fldSimple></w:p>`),
      ''
    );

    expect(evidence!.lines.some(l => l.includes('does not exist in this part'))).toBe(true);
  });

  it('caps the claim: a bookmark in another part cannot be seen from here', () => {
    const evidence = computeFieldEvidenceForMarkup(
      part(`<w:p><w:fldSimple w:instr=" REF Gone ">${run('x')}</w:fldSimple></w:p>`),
      ''
    );

    expect(evidence!.unresolved.some(u => u.includes('another part'))).toBe(true);
  });

  it('will not claim a page number is right', () => {
    const evidence = computeFieldEvidenceForMarkup(
      part(`<w:p><w:fldSimple w:instr=" PAGE ">${run('7')}</w:fldSimple></w:p>`),
      ''
    );

    expect(evidence!.unresolved.some(u => u.includes('laying the document out'))).toBe(true);
  });

  it('returns null rather than throwing on malformed XML', () => {
    expect(computeFieldEvidenceForMarkup({ 'word/document.xml': '<w:document><oops>' }, '')).toBeNull();
  });
});
