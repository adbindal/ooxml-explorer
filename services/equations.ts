/**
 * Office Math (OMML) — the equation that renders and means something else.
 *
 * Every equation in every Word document is `m:oMath`, and until this the engine had
 * nothing to say about any of them. Equations are a good fit for what this engine is
 * for, because the interesting OMML faults are not malformed markup that Word offers to
 * repair — they are **well-formed, schema-valid equations that draw cleanly and state
 * the wrong mathematics**.
 *
 * The four below all satisfy that. Each was checked against the ECMA-376 schema for
 * `shared-math.xsd` rather than assumed, because "this element is optional" is the whole
 * basis of the argument and getting it backwards would invent faults.
 *
 * | Fault                          | What the reader sees                    | Severity |
 * |--------------------------------|-----------------------------------------|----------|
 * | n-ary operator not stated      | some operator — just not the author's   | warning  |
 * | delimiter with one bracket set | mismatched brackets, e.g. `[x)`         | warning  |
 * | equation with no content       | nothing at all, where an equation was   | warning  |
 * | math run with no text          | a gap inside an otherwise fine equation | note     |
 *
 * ⚠️ WHAT THIS DELIBERATELY DOES NOT SAY.
 *
 * When `m:chr` is absent the consumer substitutes its own default, and this code does
 * **not** name that character. ECMA-376 states a default and the Open XML SDK's schema
 * data does not carry it, so naming one here would be a citation this project cannot
 * support — and five citations have already been retracted on it. The finding says the
 * operator is unstated, which is verifiable, actionable, and enough.
 */

import type { Finding, Severity } from './findings';
import type { PackageParts } from './packageIntegrity';

/** Office Math lives in the main document and every other Word story. */
export const EQUATION_HOST_PART =
  /^word\/(?:document\d*|header[^/]*|footer[^/]*|footnotes\d*|endnotes\d*)\.xml$/;

const M_NAMESPACE = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

type EquationKind =
  | 'nary-operator-unstated'
  | 'delimiter-asymmetric'
  | 'empty-equation'
  | 'empty-math-run';

/**
 * Severity and silence per kind, decided once.
 *
 * All four are silent: every one of them renders without complaint. None is an error,
 * because none makes the document unopenable — they make it *wrong*, which is a
 * different and more dangerous thing, and calling it an error would put it beside
 * faults that stop the file loading.
 */
const EQUATION_RULES = {
  'nary-operator-unstated': { severity: 'warning', silent: true },
  'delimiter-asymmetric': { severity: 'warning', silent: true },
  'empty-equation': { severity: 'warning', silent: true },
  'empty-math-run': { severity: 'note', silent: true }
} as const satisfies Record<EquationKind, { severity: Severity; silent: boolean }>;

const finding = (
  kind: EquationKind,
  part: string,
  message: string,
  remediation: string,
  subject?: Record<string, string>
): Finding => ({
  code: `equation/${kind}`,
  ...EQUATION_RULES[kind],
  part,
  message,
  remediation,
  ...(subject ? { subject } : {})
});

const parseXml = (xml: string): Document | null => {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
};

/** Direct math children with the given local name. */
const mathChildren = (element: Element, localName: string): Element[] =>
  Array.from(element.children).filter(
    el => el.namespaceURI === M_NAMESPACE && el.localName === localName
  );

const mathChild = (element: Element | null, localName: string): Element | null =>
  element ? mathChildren(element, localName)[0] ?? null : null;

/** `m:val` on a character element, which the schema makes required and single-character. */
const charValue = (element: Element | null): string | null =>
  element?.getAttributeNS(M_NAMESPACE, 'val') ?? element?.getAttribute('m:val') ?? null;

/** A short, stable description of where in the document an equation sits. */
const positionOf = (doc: Document, element: Element): string => {
  const equations = Array.from(doc.getElementsByTagNameNS(M_NAMESPACE, 'oMath'));
  const containing = equations.findIndex(eq => eq === element || eq.contains(element));
  return containing >= 0 ? `equation ${containing + 1}` : 'an equation';
};

/**
 * n-ary operators — Σ, ∏, ∫, ⋃ — whose operator character is not written down.
 *
 * `m:naryPr/m:chr` is optional in the schema (verified: `minOccurs="0"` on `CT_NaryPr`).
 * An n-ary object without it is perfectly valid and draws perfectly, using whatever
 * character the consumer defaults to. If the author meant a summation and the attribute
 * was lost — in a conversion, a hand edit, a template — the equation still renders, and
 * it now says something else. Nothing about the output reveals it.
 */
const naryFindings = (doc: Document, part: string): Finding[] =>
  Array.from(doc.getElementsByTagNameNS(M_NAMESPACE, 'nary'))
    .filter(nary => charValue(mathChild(mathChild(nary, 'naryPr'), 'chr')) === null)
    .map(nary =>
      finding(
        'nary-operator-unstated', part,
        `An n-ary operator in ${positionOf(doc, nary)} does not state its operator character: there is no m:naryPr/m:chr. ` +
        `The equation renders using whatever character the consumer defaults to, so a summation that lost this attribute draws as a different operator and reads as correct.`,
        `Write the operator explicitly as <m:naryPr><m:chr m:val="∑"/></m:naryPr>, using the character the equation is meant to show.`,
        { position: positionOf(doc, nary) }
      )
    );

/**
 * Delimiters where one bracket is customised and the other is left to the default.
 *
 * `m:begChr`, `m:sepChr` and `m:endChr` are each optional (verified on `CT_DPr`). Setting
 * one alone is legal and produces a delimiter that opens one way and closes another —
 * `[x)` — which draws cleanly and looks like a typo nobody made.
 */
const delimiterFindings = (doc: Document, part: string): Finding[] => {
  const out: Finding[] = [];
  for (const delimiter of Array.from(doc.getElementsByTagNameNS(M_NAMESPACE, 'd'))) {
    // No m:dPr at all means both ends default, which is symmetric and correct. That
    // needs no guard: mathChild(null, …) is null, so both reads below come back null and
    // the equality check below skips it. A guard here would be unfalsifiable code.
    const properties = mathChild(delimiter, 'dPr');
    const begin = charValue(mathChild(properties, 'begChr'));
    const end = charValue(mathChild(properties, 'endChr'));
    if ((begin === null) === (end === null)) continue; // both set, or both defaulted

    const stated = begin === null ? `closing "${end}"` : `opening "${begin}"`;
    const defaulted = begin === null ? 'opening' : 'closing';
    out.push(finding(
      'delimiter-asymmetric', part,
      `A delimiter in ${positionOf(doc, delimiter)} states its ${stated} bracket and leaves the ${defaulted} one to the default, so the two ends need not match. ` +
      `The equation draws either way, and a bracket that opens square and closes round reads as a typo rather than a lost attribute.`,
      `State both ends explicitly with m:begChr and m:endChr, or neither, so the pair cannot drift apart.`,
      { position: positionOf(doc, delimiter), stated: begin ?? end ?? '' }
    ));
  }
  return out;
};

/**
 * Equations with nothing in them, and math runs with no text.
 *
 * `CT_OMath` wraps `EG_OMathElements` with `minOccurs="0"`, so an empty `m:oMath` is
 * **schema-valid** — it is not a malformed document, it is an equation that draws as
 * blank space. `m:t` is likewise optional on a math run, which leaves a hole inside an
 * equation that is otherwise correct.
 */
const emptinessFindings = (doc: Document, part: string): Finding[] => {
  const out: Finding[] = [];

  for (const equation of Array.from(doc.getElementsByTagNameNS(M_NAMESPACE, 'oMath'))) {
    // Properties-only content is still empty: m:oMathPara wraps equations, and an
    // equation holding nothing but control properties draws nothing.
    const hasContent = Array.from(equation.children).some(
      el => el.namespaceURI !== M_NAMESPACE || !['oMathParaPr', 'ctrlPr'].includes(el.localName)
    );
    if (hasContent) continue;

    out.push(finding(
      'empty-equation', part,
      `An equation in this part contains nothing. The schema permits it — m:oMath may hold no elements at all — so the file opens cleanly and the equation occupies no space, which looks identical to a paragraph that never had one.`,
      `Remove the empty m:oMath, or restore the content it was meant to hold.`
    ));
  }

  for (const run of Array.from(doc.getElementsByTagNameNS(M_NAMESPACE, 'r'))) {
    if (mathChildren(run, 't').length > 0) continue;
    out.push(finding(
      'empty-math-run', part,
      `A math run in ${positionOf(doc, run)} carries no m:t, so it contributes no characters. The surrounding equation renders normally with a gap where this run's text should be.`,
      `Give the run its m:t text, or delete the run.`,
      { position: positionOf(doc, run) }
    ));
  }
  return out;
};

/** Every equation finding for one part. */
export function equationFindings(parts: PackageParts, partPath: string): Finding[] {
  const xml = parts[partPath];
  if (xml === undefined) return [];
  const doc = parseXml(xml);
  if (!doc) return [];
  // Purely an optimisation for the overwhelming majority of documents, which have no
  // equations. Removing it changes no result — every walk below would find nothing — so
  // no test covers it, and nothing should be made to depend on it.
  if (doc.getElementsByTagNameNS(M_NAMESPACE, 'oMath').length === 0) return [];

  return [
    ...naryFindings(doc, partPath),
    ...delimiterFindings(doc, partPath),
    ...emptinessFindings(doc, partPath)
  ];
}

/** True when any Word story in the package carries an equation. */
export const hasEquations = (parts: PackageParts): boolean =>
  Object.keys(parts).some(
    path => EQUATION_HOST_PART.test(path) && (parts[path] ?? '').includes('oMath')
  );
