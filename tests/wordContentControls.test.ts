import { describe, it, expect } from 'vitest';
import {
  readContentControls,
  readDataStoreIds,
  normaliseItemId,
  contentControlFindings,
  computeContentControlEvidenceForMarkup,
  DS_NAMESPACE
} from '../services/wordContentControls';
import type { PackageParts } from '../services/packageIntegrity';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const doc = (body: string) =>
  new DOMParser().parseFromString(
    `<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`,
    'application/xml'
  );

const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;

/** A content control: props (given verbatim) wrapping some stored content. */
const sdt = (props: string, content = run('stored value')) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

const binding = (id: string, xpath = '/root/customer/name') =>
  `<w:dataBinding w:xpath="${xpath}" w:storeItemID="${id}"/>`;

const itemProps = (id: string) =>
  `<?xml version="1.0"?><ds:datastoreItem xmlns:ds="${DS_NAMESPACE}" ds:itemID="${id}"><ds:schemaRefs/></ds:datastoreItem>`;

const GUID = '{7DFD8E1E-1F9E-4A2B-9C3E-8F1A2B3C4D5E}';

describe('the data store', () => {
  it('collects item ids from the itemProps parts', () => {
    const ids = readDataStoreIds({ 'customXml/itemProps1.xml': itemProps(GUID) });

    expect(ids.has(normaliseItemId(GUID))).toBe(true);
  });

  it('ignores the data parts themselves, which carry no item id', () => {
    expect(readDataStoreIds({ 'customXml/item1.xml': '<root><customer/></root>' }).size).toBe(0);
  });

  it('reads ids only from itemProps parts, not from anything under customXml', () => {
    // Contrived on purpose: the element check alone would accept this, so only the
    // path filter rejects it. Item ids are declared in itemProps by definition, and
    // honouring one found elsewhere would let a data part vouch for a binding.
    const strayed = readDataStoreIds({ 'customXml/item1.xml': itemProps(GUID) });

    expect(strayed.size).toBe(0);
  });

  it('matches ids case-insensitively and without braces', () => {
    // Generators are inconsistent about both, and a binding differing only in case is
    // not broken - reporting it would be a false alarm on working documents.
    const ids = readDataStoreIds({ 'customXml/itemProps1.xml': itemProps(GUID.toLowerCase()) });

    expect(ids.has(normaliseItemId(GUID.toUpperCase()))).toBe(true);
    expect(normaliseItemId('{ABC}')).toBe(normaliseItemId('abc'));
  });

  it('survives an itemProps part that is not well-formed', () => {
    expect(() => readDataStoreIds({ 'customXml/itemProps1.xml': '<ds:x><unclosed>' })).not.toThrow();
  });
});

describe('reading controls', () => {
  it('reads alias, tag, id and stored content', () => {
    const { controls } = readContentControls(
      doc(`<w:p>${sdt('<w:alias w:val="Customer Name"/><w:tag w:val="cust_name"/><w:id w:val="101"/>')}</w:p>`)
    );

    expect(controls[0]).toMatchObject({ alias: 'Customer Name', tag: 'cust_name', id: '101', content: 'stored value' });
  });

  it('reads the binding', () => {
    const { controls } = readContentControls(doc(`<w:p>${sdt(binding(GUID))}</w:p>`));

    expect(controls[0].binding).toEqual({
      xpath: '/root/customer/name',
      storeItemID: GUID,
      prefixMappings: null
    });
  });

  it('finds controls wrapped at run level as well as block level', () => {
    // There are three w:sdt declarations (block, run, run-ruby) and they are matched by
    // element name, not by parent - so a control inside a paragraph must be found too.
    const { controls } = readContentControls(doc(`<w:p>${run('before')}${sdt('<w:tag w:val="inline"/>')}</w:p>`));

    expect(controls.map(c => c.tag)).toEqual(['inline']);
  });

  it('returns nothing for a document with no controls', () => {
    expect(readContentControls(doc(`<w:p>${run('plain')}</w:p>`)).controls).toEqual([]);
  });
});

describe('the broken binding — the reason this module exists', () => {
  const parts = (extra: Partial<PackageParts> = {}): PackageParts => ({
    'word/document.xml': `<?xml version="1.0"?><w:document ${W}><w:body><w:p>${sdt(
      `<w:alias w:val="Customer Name"/>${binding(GUID)}`,
      run('Acme Corp')
    )}</w:p></w:body></w:document>`,
    ...extra
  });

  it('reports a binding whose store item no part declares', () => {
    const problem = contentControlFindings(parts()).find(p => p.code === 'contentControl/binding-part-missing');

    expect(problem?.subject?.storeItemID).toBe(GUID);
    expect(problem?.silent).toBe(true);
  });

  it('says the control keeps displaying its stored text, which is the whole problem', () => {
    const problem = contentControlFindings(parts()).find(p => p.code === 'contentControl/binding-part-missing');

    expect(problem?.message).toContain('Acme Corp');
    expect(problem?.message).toContain('looks populated and is not connected');
  });

  it('says nothing when the store item is present', () => {
    const codes = contentControlFindings(parts({ 'customXml/itemProps1.xml': itemProps(GUID) })).map(p => p.code);

    expect(codes).not.toContain('contentControl/binding-part-missing');
  });

  it('matches the store item across brace and case differences', () => {
    // The binding uses braces and upper case; the itemProps uses neither. This is a
    // working document and must not be reported as broken.
    const codes = contentControlFindings(
      parts({ 'customXml/itemProps1.xml': itemProps(GUID.replace(/[{}]/g, '').toLowerCase()) })
    ).map(p => p.code);

    expect(codes).not.toContain('contentControl/binding-part-missing');
  });

  it('reports a binding missing one of its two required attributes', () => {
    const { problems } = readContentControls(
      doc(`<w:p>${sdt('<w:dataBinding w:xpath="/root/x"/>')}</w:p>`),
      'word/document.xml',
      new Set()
    );

    const problem = problems.find(p => p.code === 'contentControl/binding-incomplete');
    expect(problem?.message).toContain('w:storeItemID');
  });

  it('does not also report a part-missing problem for an incomplete binding', () => {
    // It has no storeItemID to look up, so reporting both would be two findings for
    // one fault and would send the reader looking for a part that was never named.
    const { problems } = readContentControls(
      doc(`<w:p>${sdt('<w:dataBinding w:xpath="/root/x"/>')}</w:p>`),
      'word/document.xml',
      new Set()
    );

    expect(problems.map(p => p.code)).not.toContain('contentControl/binding-part-missing');
  });

  it('does not check bindings at all when no store ids were supplied', () => {
    // Only one part available: the structural checks still run, but claiming a binding
    // is broken without having looked at the package would be a guess.
    const { problems } = readContentControls(doc(`<w:p>${sdt(binding(GUID))}</w:p>`));

    expect(problems.map(p => p.code)).not.toContain('contentControl/binding-part-missing');
  });
});

describe('placeholders and structure', () => {
  it('reports a control showing its placeholder as visible, unlike the rest', () => {
    const { problems } = readContentControls(
      doc(`<w:p>${sdt('<w:tag w:val="t"/><w:showingPlcHdr/>', run('Click here to enter text'))}</w:p>`)
    );

    const problem = problems.find(p => p.code === 'contentControl/showing-placeholder');
    expect(problem?.silent).toBe(false);
    expect(problem?.message).toContain('never filled');
  });

  it('does not report a placeholder on a control that has been filled', () => {
    const { problems } = readContentControls(doc(`<w:p>${sdt('<w:tag w:val="t"/>', run('Acme'))}</w:p>`));

    expect(problems.map(p => p.code)).not.toContain('contentControl/showing-placeholder');
  });

  it('reports two controls sharing an id, since anything driving by id reaches one', () => {
    const { problems } = readContentControls(
      doc(`<w:p>${sdt('<w:alias w:val="First"/><w:id w:val="5"/>')}${sdt('<w:alias w:val="Second"/><w:id w:val="5"/>')}</w:p>`)
    );

    const problem = problems.find(p => p.code === 'contentControl/duplicate-id');
    expect(problem?.message).toContain('First');
    expect(problem?.message).toContain('Second');
  });

  it('names the first control as the original for every later duplicate', () => {
    // Three controls sharing an id. Each duplicate must point back at the FIRST, not at
    // the previous duplicate - with only two controls the difference is invisible.
    const { problems } = readContentControls(
      doc(
        `<w:p>${sdt('<w:alias w:val="First"/><w:id w:val="5"/>')}` +
          `${sdt('<w:alias w:val="Second"/><w:id w:val="5"/>')}` +
          `${sdt('<w:alias w:val="Third"/><w:id w:val="5"/>')}</w:p>`
      )
    );

    const duplicates = problems.filter(p => p.code === 'contentControl/duplicate-id');
    expect(duplicates).toHaveLength(2);
    for (const d of duplicates) expect(d.message).toContain('First');
  });

  it('does not report distinct ids as duplicates', () => {
    const { problems } = readContentControls(
      doc(`<w:p>${sdt('<w:id w:val="5"/>')}${sdt('<w:id w:val="6"/>')}</w:p>`)
    );

    expect(problems.map(p => p.code)).not.toContain('contentControl/duplicate-id');
  });

  it('reports a control with no content element', () => {
    const { problems } = readContentControls(doc(`<w:p><w:sdt><w:sdtPr><w:tag w:val="t"/></w:sdtPr></w:sdt></w:p>`));

    expect(problems.map(p => p.code)).toContain('contentControl/no-content-element');
  });

  it('notes an unbound control without treating it as a fault', () => {
    const { problems } = readContentControls(doc(`<w:p>${sdt('<w:tag w:val="manual"/>')}</w:p>`));

    expect(problems.find(p => p.code === 'contentControl/unbound-control')?.severity).toBe('note');
  });
});

describe('computeContentControlEvidenceForMarkup', () => {
  const parts = (body: string, extra: Partial<PackageParts> = {}) => ({
    'word/document.xml': `<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`,
    ...extra
  });

  it('returns null when the part has no controls', () => {
    expect(computeContentControlEvidenceForMarkup(parts(`<w:p>${run('plain')}</w:p>`), '')).toBeNull();
  });

  it('leads with how many controls are bound versus filled by hand', () => {
    const evidence = computeContentControlEvidenceForMarkup(
      parts(`<w:p>${sdt(binding(GUID))}${sdt('<w:tag w:val="manual"/>')}</w:p>`),
      ''
    );

    expect(evidence!.lines[0]).toContain('1 bound to custom XML, 1 filled by hand');
  });

  it('describes the selected control', () => {
    const evidence = computeContentControlEvidenceForMarkup(
      parts(`<w:p>${sdt(`<w:tag w:val="cust"/>${binding(GUID)}`, run('Acme'))}</w:p>`),
      '<w:tag w:val="cust"/>'
    );

    expect(evidence!.lines.some(l => l.includes('"cust"') && l.includes('Acme'))).toBe(true);
  });

  it('admits the XPath was never evaluated', () => {
    // The limit this module deliberately does not cross: checking that an expression
    // selects a node needs a namespace-aware engine, and getting it subtly wrong would
    // produce confident false reports about working templates.
    const evidence = computeContentControlEvidenceForMarkup(parts(`<w:p>${sdt(binding(GUID))}</w:p>`), '');

    expect(evidence!.unresolved.some(u => u.includes('XPath expression'))).toBe(true);
  });

  it('does not add that caveat when nothing is bound', () => {
    const evidence = computeContentControlEvidenceForMarkup(parts(`<w:p>${sdt('<w:tag w:val="manual"/>')}</w:p>`), '');

    expect(evidence!.unresolved.some(u => u.includes('XPath expression'))).toBe(false);
  });
});
