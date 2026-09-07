/**
 * Turning an attribute spec into text — for the prompt, and for the search index.
 *
 * The corpus used to record that `w:jc` has an attribute called `w:val`. It now records
 * that `w:val` is an enumeration of twelve alignment keywords, that it is required, and
 * that it is labelled "Alignment Type". That is the difference between grounding a
 * reader in the element's shape and actually answering their question.
 *
 * Both consumers went through `attributes.join(', ')` while attributes were strings.
 * They are objects now, and `join` still compiles — it just yields `[object Object]`.
 * In the search index that silently destroys recall; in the prompt it puts that string
 * in front of the model **under a Grounded badge**, which is worse than saying nothing.
 * One rendering function, used by both, is what keeps that from happening twice.
 *
 * ⚠️ THE VALUE LIST IS CAPPED, AND THAT IS NOT COSMETIC.
 *
 * The largest enumeration in the corpus has 197 values, and the worst single element
 * would cost roughly 986 tokens if every value were printed — against a local model
 * window of 6k–9k that is shared with the output. Rendering the lot would reintroduce
 * the unbounded-prompt defect that `promptBudget.ts` exists to prevent, in a path that
 * budgeting does not cover. The cap states what it omitted rather than trailing off, so
 * a reader can tell "these are the values" from "these are some of the values".
 */

import type { AttributeSpec } from './staticKnowledgeBase';

/**
 * How many enumerated values to print before summarising the rest.
 *
 * Twelve covers the ninetieth percentile of enumerations outright (p90 is 17, median 5),
 * so the cap is invisible for almost every attribute anyone asks about, and only bites
 * on the long tail where a full list would not have been read anyway.
 */
export const MAX_LISTED_VALUES = 12;

/** `left, center, right … and 185 more` — never a bare ellipsis. */
export const renderValues = (values: readonly string[], limit = MAX_LISTED_VALUES): string => {
  if (values.length <= limit) return values.join(', ');
  const omitted = values.length - limit;
  return `${values.slice(0, limit).join(', ')} … and ${omitted} more`;
};

/**
 * One attribute as a line of prompt context.
 *
 * Ordered by what a reader needs first: the name, then what it means, then what it
 * accepts, then whether it may be left out. Only facts the schema actually carried are
 * printed — an attribute with no enumeration says nothing about values rather than
 * claiming it accepts anything.
 */
export const renderAttribute = (spec: AttributeSpec, limit = MAX_LISTED_VALUES): string => {
  const parts: string[] = [spec.name];
  if (spec.label && spec.label !== spec.name) parts.push(`(${spec.label})`);

  const facts: string[] = [];
  if (spec.values?.length) facts.push(`one of: ${renderValues(spec.values, limit)}`);
  else if (spec.type && spec.type !== 'unknown') facts.push(spec.type);

  if (spec.required) facts.push('required');
  if (spec.min !== undefined && spec.max !== undefined) facts.push(`${spec.min} to ${spec.max}`);
  else if (spec.min !== undefined) facts.push(`${spec.min} or greater`);
  else if (spec.max !== undefined) facts.push(`${spec.max} or less`);
  if (spec.maxLength !== undefined) facts.push(`max length ${spec.maxLength}`);
  if (spec.version) facts.push(`${spec.version} and later`);

  return facts.length > 0 ? `${parts.join(' ')} — ${facts.join('; ')}` : parts.join(' ');
};

/** Every attribute, one per line, indented for the prompt block. */
export const renderAttributes = (specs: readonly AttributeSpec[], limit = MAX_LISTED_VALUES): string =>
  specs.length === 0 ? 'None' : `\n${specs.map(s => `    ${renderAttribute(s, limit)}`).join('\n')}`;

/**
 * Search text for one attribute.
 *
 * Deliberately richer than the prompt rendering and uncapped: this is never shown to
 * anyone, it only feeds BM25, and every omitted value is a query that cannot match.
 * Indexing the label and the values is what lets "alignment" find `w:jc`, which the
 * previous index — attribute names only — could not do.
 */
export const attributeSearchText = (spec: AttributeSpec): string =>
  [spec.name, spec.label ?? '', ...(spec.values ?? [])].join(' ');
