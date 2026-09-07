import { describe, it, expect } from 'vitest';
import {
  renderAttribute,
  renderAttributes,
  renderValues,
  attributeSearchText,
  MAX_LISTED_VALUES
} from '../services/schemaFacts';
import type { AttributeSpec } from '../services/staticKnowledgeBase';

/**
 * Rendering attribute specs for the prompt and the search index.
 *
 * Two things are load-bearing here and both are easy to get quietly wrong: the value
 * list must be capped, because the largest enumeration in the corpus has 197 values
 * against a local context window of 6k–9k shared with the output; and the prompt must
 * never claim more than the schema said, because whatever it prints appears under a
 * Grounded badge.
 */

const spec = (over: Partial<AttributeSpec> = {}): AttributeSpec =>
  ({ name: 'w:val', type: 'string', ...over });

describe('rendering one attribute', () => {
  it('leads with the name, then the label, then what it accepts', () => {
    const line = renderAttribute(spec({
      type: 'enum',
      values: ['left', 'center', 'right'],
      label: 'Alignment Type',
      required: true
    }));

    expect(line).toBe('w:val (Alignment Type) — one of: left, center, right; required');
  });

  it('does not repeat the name as its own label', () => {
    expect(renderAttribute(spec({ label: 'w:val' }))).toBe('w:val — string');
  });

  it('reports the type only when there is no value list', () => {
    // "one of: left, center, right" already says it is an enumeration; adding the word
    // enum spends context to repeat what the list shows.
    const enumerated = renderAttribute(spec({ type: 'enum', values: ['a', 'b'] }));
    expect(enumerated).not.toContain('enum');
    expect(renderAttribute(spec({ type: 'integer' }))).toContain('integer');
  });

  it('says nothing about a type it does not know', () => {
    // Curated OPC records carry `unknown`, meaning no schema was consulted. Printing it
    // would dress a gap up as a finding.
    expect(renderAttribute(spec({ type: 'unknown' }))).toBe('w:val');
  });

  it('falls back to the type when a value list is present but empty', () => {
    // The corpus never emits this — ragCorpusInvariants pins that — but the renderer
    // must not depend on the generator being careful. `if (spec.values)` is true for
    // an empty array and would print "one of: " with nothing after it, which reads as
    // an element that accepts no values at all.
    const line = renderAttribute(spec({ type: 'integer', values: [] }));

    expect(line).not.toContain('one of');
    expect(line).toBe('w:val — integer');
  });

  it('renders bounds, length and version gates when the schema gave them', () => {
    expect(renderAttribute(spec({ type: 'integer', min: 0, max: 100 }))).toContain('0 to 100');
    expect(renderAttribute(spec({ type: 'integer', min: 0 }))).toContain('0 or greater');
    expect(renderAttribute(spec({ type: 'integer', max: 9 }))).toContain('9 or less');
    expect(renderAttribute(spec({ maxLength: 255 }))).toContain('max length 255');
    expect(renderAttribute(spec({ version: 'Office2010' }))).toContain('Office2010 and later');
  });

  it('treats a zero bound as a bound, not as absent', () => {
    // `if (spec.min)` is false for 0, which would silently drop the most common bound
    // in the corpus — "0 or greater" appears on hundreds of measurement attributes.
    expect(renderAttribute(spec({ type: 'integer', min: 0 }))).toContain('0 or greater');
    expect(renderAttribute(spec({ type: 'integer', max: 0 }))).toContain('0 or less');
  });
});

describe('the value cap', () => {
  const many = Array.from({ length: 197 }, (_, i) => `v${i}`);

  it('prints every value when the list is short enough', () => {
    const values = ['left', 'center', 'right'];
    expect(renderValues(values)).toBe('left, center, right');
  });

  it('caps a long list and says how many it left out', () => {
    const rendered = renderValues(many);

    expect(rendered.startsWith('v0, v1, v2')).toBe(true);
    expect(rendered).toContain(`… and ${197 - MAX_LISTED_VALUES} more`);
    expect(rendered.split(',').length).toBeLessThan(many.length);
  });

  it('never trails off without saying so', () => {
    // A bare ellipsis leaves the reader unable to tell "these are the values" from
    // "these are some of the values", which is the difference between a fact and a
    // half-fact under a Grounded badge.
    expect(renderValues(many)).toMatch(/… and \d+ more$/);
  });

  it('keeps the worst real element inside a sane prompt budget', () => {
    // The 197-value enumeration would cost roughly 986 tokens rendered in full. This is
    // the check that the cap actually binds, rather than existing and never applying.
    const rendered = renderAttribute(spec({ type: 'enum', values: many }));
    expect(rendered.length).toBeLessThan(200);
  });

  it('does not cap at the boundary, but does one past it', () => {
    const exactly = Array.from({ length: MAX_LISTED_VALUES }, (_, i) => `v${i}`);
    expect(renderValues(exactly)).not.toContain('more');
    expect(renderValues([...exactly, 'extra'])).toContain('and 1 more');
  });
});

describe('rendering the whole attribute block', () => {
  it('says None rather than printing an empty block', () => {
    expect(renderAttributes([])).toBe('None');
  });

  it('puts one attribute per line, indented', () => {
    const block = renderAttributes([spec({ name: 'w:a' }), spec({ name: 'w:b' })]);
    expect(block.split('\n').filter(Boolean)).toEqual(['    w:a — string', '    w:b — string']);
  });
});

describe('search text', () => {
  it('indexes the label and every value, uncapped', () => {
    // This is never shown to anyone, so the cap would only cost recall: each omitted
    // value is a query that silently cannot match.
    const text = attributeSearchText(spec({
      label: 'Alignment Type',
      values: ['left', 'center', 'right', 'distribute']
    }));

    expect(text).toContain('Alignment Type');
    expect(text).toContain('distribute');
  });

  it('survives an attribute carrying nothing but a name', () => {
    expect(() => attributeSearchText({ name: 'w:val', type: 'string' })).not.toThrow();
    expect(attributeSearchText({ name: 'w:val', type: 'string' })).toContain('w:val');
  });
});
