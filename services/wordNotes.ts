/**
 * Footnotes and endnotes — a reference mark with nothing behind it.
 *
 * A note is split across two parts. The body carries a mark:
 *
 *   <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
 *        <w:footnoteReference w:id="3"/></w:r>
 *
 * and `word/footnotes.xml` carries the text, matched **by `@w:id`**. Break the pairing
 * and neither half complains: the superscript **1** still renders in the body, the
 * reader still sees a footnote marker, and there is simply no note at the bottom of the
 * page. Text extraction returns the mark and drops the content. Nothing reports a fault.
 *
 * The reverse case is quieter still — a `w:footnote` that nothing references is never
 * displayed at all. Its text sits in the file, invisible, and survives every copy,
 * conversion and review of the document.
 *
 * ⚠️ THE TRAP THAT WOULD FIRE ON EVERY DOCUMENT EVER WRITTEN.
 *
 * `footnotes.xml` always contains notes that **nothing in the body references**, and
 * they are not orphans. They are the separator graphics:
 *
 *   <w:footnote w:type="separator" w:id="-1">…the horizontal rule…</w:footnote>
 *   <w:footnote w:type="continuationSeparator" w:id="0">…</w:footnote>
 *
 * These are referenced from `sectPr/footnotePr/footnote`, not from the text, and Word
 * writes them into every document that has ever had a footnote. A naive "every note must
 * be referenced" check reports two false positives on **100% of real files**, which is
 * the fastest way to make a report unreadable. Only `w:type="normal"` — or an absent
 * `@w:type`, which means normal — is a content note.
 *
 * Verified against the Open XML SDK schema: `w:footnoteReference`/`w:endnoteReference`
 * are `CT_FtnEdnRef` with **required** `@w:id` and optional `@w:customMarkFollows`;
 * `w:footnote`/`w:endnote` in the notes part are `CT_FtnEdn` with **required** `@w:id`
 * and `@w:type` from `Normal | Separator | ContinuationSeparator | ContinuationNotice`;
 * and the `w:footnote` inside section properties is a *different type* (`CT_FtnEdnSepRef`)
 * carrying only an id. Note `@w:id` is `IntegerValue` — **signed**, so the `-1` above is
 * ordinary and legal, unlike the bookmark id union which excludes it.
 */

import { W_NAMESPACE } from './wordStyleResolver';
import { finding, renderFindings, type Finding, type Severity } from './findings';
import type { PackageParts } from './packageIntegrity';

/**
 * Severity and silence per kind.
 *
 * All silent. A note failure never stops the document rendering — it removes content
 * while leaving the marker that promises it, or hides content that is still in the file.
 */
const NOTE_RULES = {
  'missing-note': { severity: 'error', silent: true },
  'orphan-note': { severity: 'warning', silent: true },
  'duplicate-note-id': { severity: 'error', silent: true },
  'missing-separator': { severity: 'note', silent: true },
  'notes-part-missing': { severity: 'error', silent: true }
} as const satisfies Record<string, { severity: Severity; silent: boolean }>;

export type NoteProblemKind = keyof typeof NOTE_RULES;

const noteFinding = (
  kind: NoteProblemKind,
  part: string,
  message: string,
  remediation: string,
  subject?: Record<string, string>
): Finding => finding(`note/${kind}`, part, message, remediation, { ...NOTE_RULES[kind], subject });

export type NoteFlavour = 'footnote' | 'endnote';

/** Where each flavour keeps its text, and what the body element is called. */
const FLAVOURS: Record<NoteFlavour, { part: string; reference: string; note: string }> = {
  footnote: { part: 'word/footnotes.xml', reference: 'footnoteReference', note: 'footnote' },
  endnote: { part: 'word/endnotes.xml', reference: 'endnoteReference', note: 'endnote' }
};

/**
 * Note types that exist to draw the separator line rather than to carry content.
 *
 * Nothing in the body references these, and every real document has them.
 */
const STRUCTURAL_TYPES = new Set(['separator', 'continuationSeparator', 'continuationNotice']);

const isW = (el: Element, local: string) => el.namespaceURI === W_NAMESPACE && el.localName === local;
const wAttr = (el: Element, name: string) => el.getAttributeNS(W_NAMESPACE, name);

const parseXml = (xml: string | undefined): Document | null => {
  if (xml === undefined) return null;
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
};

export interface Note {
  id: string;
  /** `normal` when `@w:type` is absent — the schema default, and the content case. */
  type: string;
  /** True for the separator graphics, which the body never references. */
  structural: boolean;
  text: string;
}

/** Reads one notes part. */
export function readNotes(doc: Document, flavour: NoteFlavour): Note[] {
  const root = doc.documentElement;
  if (!root) return [];
  const { note: noteName } = FLAVOURS[flavour];

  return Array.from(root.children)
    .filter(el => isW(el, noteName))
    .map(el => {
      const type = wAttr(el, 'type') ?? 'normal';
      return {
        id: wAttr(el, 'id') ?? '',
        type,
        structural: STRUCTURAL_TYPES.has(type),
        text: Array.from(el.getElementsByTagName('*'))
          .filter(c => isW(c, 't'))
          .map(c => c.textContent ?? '')
          .join('')
      };
    });
}

/** Every note id referenced from a body part. */
export function readNoteReferences(doc: Document, flavour: NoteFlavour): string[] {
  const root = doc.documentElement;
  if (!root) return [];
  const { reference } = FLAVOURS[flavour];

  return Array.from(root.querySelectorAll('*'))
    .filter(el => isW(el, reference))
    .map(el => wAttr(el, 'id'))
    .filter((id): id is string => id !== null);
}

/** Word body parts that can carry note references. */
const WORD_BODY = /^word\/(?:document\d*|header[^/]*|footer[^/]*)\.xml$/;

/** True when this package has notes worth checking. */
export const hasNotes = (parts: PackageParts): boolean =>
  parts['word/footnotes.xml'] !== undefined || parts['word/endnotes.xml'] !== undefined;

/** Checks one flavour across the package. */
function checkFlavour(parts: PackageParts, flavour: NoteFlavour): Finding[] {
  const { part, reference, note: noteName } = FLAVOURS[flavour];
  const problems: Finding[] = [];

  const referenced = new Set<string>();
  const bodyParts = Object.keys(parts).filter(p => WORD_BODY.test(p));
  for (const path of bodyParts) {
    const doc = parseXml(parts[path]);
    if (doc) for (const id of readNoteReferences(doc, flavour)) referenced.add(id);
  }

  const notesDoc = parseXml(parts[part]);
  if (!notesDoc) {
    if (referenced.size > 0) {
      problems.push(noteFinding(
        'notes-part-missing', part,
        `The document places ${referenced.size} ${flavour} reference mark(s) but ${part} is absent or unreadable. Every mark still renders as a superscript number, so the page looks right and there is no note text anywhere in the file.`,
        `Restore ${part}, or remove the <w:${reference}> elements so the marks stop promising notes that do not exist.`
      ));
    }
    return problems;
  }

  const notes = readNotes(notesDoc, flavour);
  const byId = new Map<string, Note>();
  for (const n of notes) {
    if (byId.has(n.id)) {
      problems.push(noteFinding(
        'duplicate-note-id', part,
        `Two ${flavour}s share w:id "${n.id}". References match by id, so one of the two texts is unreachable — and which one a consumer picks is not something the markup decides.`,
        `Renumber one of them to an id unused in ${part}.`,
        { id: n.id }
      ));
      continue;
    }
    byId.set(n.id, n);
  }

  for (const id of referenced) {
    if (byId.has(id)) continue;
    problems.push(noteFinding(
      'missing-note', part,
      `A ${flavour} reference points at id ${id}, which ${part} does not define. The superscript mark still renders in the body — the reader sees a ${flavour} marker — and there is no note at the bottom of the page. Text extraction returns the mark and drops the content.`,
      `Add a <w:${noteName} w:id="${id}"> to ${part}, or delete the reference from the body.`,
      { id, flavour }
    ));
  }

  for (const n of notes) {
    // The separator graphics are referenced from sectPr, never from the text. Reporting
    // them would put two false positives on essentially every real document.
    if (n.structural || referenced.has(n.id)) continue;
    problems.push(noteFinding(
      'orphan-note', part,
      `${part} defines ${flavour} ${n.id}${n.text ? ` ("${n.text.slice(0, 60)}${n.text.length > 60 ? '…' : ''}")` : ''}, which nothing in the document references. It is never displayed, and its text travels with the file through every copy, conversion and review without appearing anywhere.`,
      `Add a <w:${reference} w:id="${n.id}"/> where the note belongs, or delete the ${flavour} from ${part}.`,
      { id: n.id, flavour }
    ));
  }

  // A document that uses notes but ships no separator will get Word's default rule; a
  // converter that reproduces the file faithfully gets nothing, and the notes run into
  // the body text with no divider.
  if (referenced.size > 0 && !notes.some(n => n.type === 'separator')) {
    problems.push(noteFinding(
      'missing-separator', part,
      `${part} declares no separator note, so nothing defines the rule drawn between the body text and the ${flavour}s. Word supplies a default; a consumer reproducing this file exactly draws no divider at all.`,
      `Add <w:${noteName} w:type="separator" w:id="-1"> with the separator content, as Word writes by default.`
    ));
  }

  return problems;
}

/** Every footnote and endnote finding in the package. */
export const noteFindings = (parts: PackageParts): Finding[] => [
  ...checkFlavour(parts, 'footnote'),
  ...checkFlavour(parts, 'endnote')
];

/** Evidence lines for the AI panel. */
export function computeNoteEvidenceForMarkup(
  parts: Record<string, string>
): { lines: string[]; unresolved: string[] } | null {
  if (!hasNotes(parts)) return null;

  const lines: string[] = [];
  const unresolved: string[] = [];

  for (const flavour of ['footnote', 'endnote'] as const) {
    const doc = parseXml(parts[FLAVOURS[flavour].part]);
    if (!doc) continue;
    const notes = readNotes(doc, flavour);
    const content = notes.filter(n => !n.structural);
    lines.push(
      `${FLAVOURS[flavour].part} holds ${content.length} ${flavour}(s) plus ${notes.length - content.length} ` +
        'separator note(s), which draw the rule above the notes and are referenced from the section properties rather than from the text.'
    );
  }

  const problems = noteFindings(parts);
  if (problems.length === 0) {
    lines.push('Every note reference resolves, and every note is referenced.');
  } else {
    lines.push(...renderFindings(problems));
  }

  unresolved.push(
    'Notes were matched only against the body parts supplied with this request. A reference from a part not loaded here would make a note look orphaned when it is not.'
  );

  return { lines, unresolved };
}
