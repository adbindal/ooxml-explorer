import { describe, it, expect } from 'vitest';
import { selectBestMatch } from '../services/storageService';
import type { ReferenceDoc } from '../services/staticKnowledgeBase';

/**
 * Regression tests for a shipped bug: the RAG lookup ignored the namespace prefix.
 *
 * `ragRouter` received `context.namespace` and used it only for display, so a tag
 * whose local name exists in several namespaces resolved by domain alone. 103 tags in
 * the corpus are in that position. Clicking a DrawingML `<a:bottom>` inside a .docx
 * returned the WordprocessingML `<w:bottom>` record - different attributes, different
 * parents - underneath a "Grounded" badge.
 */

const doc = (tag: string, namespace: string, domain: ReferenceDoc['domain']): ReferenceDoc =>
  ({ tag, namespace, domain, attributes: [], parents: [] });

// The real collision: three different elements, one local name.
const bottoms = [
  doc('bottom', 'w', 'docx'),
  doc('bottom', 'x', 'xlsx'),
  doc('bottom', 'a', 'shared')
];

describe('the bug this fixes', () => {
  it('does NOT return w:bottom for a DrawingML a:bottom inside a .docx', () => {
    const match = selectBestMatch(bottoms, 'docx', 'a');
    expect(match?.namespace).toBe('a');
    expect(match?.domain).toBe('shared');
  });

  it('still returns w:bottom for a genuine WordprocessingML w:bottom', () => {
    expect(selectBestMatch(bottoms, 'docx', 'w')?.namespace).toBe('w');
  });

  it('returns x:bottom in a spreadsheet, not w:bottom', () => {
    expect(selectBestMatch(bottoms, 'xlsx', 'x')?.namespace).toBe('x');
  });
});

describe('scope filtering', () => {
  it('excludes records from an unrelated domain', () => {
    // x:bottom is not reachable from a .docx at all.
    expect(selectBestMatch(bottoms, 'docx', 'x')).toBeNull();
  });

  it('always allows shared records into scope', () => {
    expect(selectBestMatch([doc('blip', 'a', 'shared')], 'pptx', 'a')?.namespace).toBe('a');
  });

  it('returns null when nothing is in scope', () => {
    expect(selectBestMatch([doc('sheetPr', 'x', 'xlsx')], 'docx', 'x')).toBeNull();
  });

  it('returns null for an empty candidate list', () => {
    expect(selectBestMatch([], 'docx', 'w')).toBeNull();
  });
});

describe('unmatched prefixes', () => {
  it('refuses to guess when the prefix is unknown and candidates are ambiguous', () => {
    // A document may bind an unconventional prefix. Returning a same-named element
    // from the wrong namespace under a "Grounded" badge is worse than admitting the
    // tag is not covered.
    expect(selectBestMatch(bottoms, 'docx', 'zz')).toBeNull();
  });

  it('accepts an unconventional prefix when there is nothing to choose between', () => {
    // <foo:blip> where foo is bound to the DrawingML namespace: only one candidate
    // is in scope, so there is no wrong answer available.
    const only = [doc('blip', 'a', 'shared')];
    expect(selectBestMatch(only, 'docx', 'foo')?.namespace).toBe('a');
  });
});

describe('domain preference', () => {
  it('prefers the document domain over shared when both match the prefix', () => {
    const both = [doc('ext', 'a', 'shared'), doc('ext', 'a', 'docx')];
    expect(selectBestMatch(both, 'docx', 'a')?.domain).toBe('docx');
  });

  it('falls back to domain preference when no namespace is supplied', () => {
    // Keyword-search hits have no prefix to work with.
    expect(selectBestMatch(bottoms, 'docx')?.domain).toBe('docx');
  });
});
