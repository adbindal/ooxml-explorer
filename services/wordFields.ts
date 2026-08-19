/**
 * WordprocessingML fields — the text on the page is a CACHE, and it can lie.
 *
 * A field is a small program embedded in the document: `REF`, `PAGEREF`, `TOC`,
 * `HYPERLINK`, `PAGE`. What a reader sees is not the program's output — it is the
 * output **from the last time Word ran it**, stored in the file alongside the
 * instruction. That single fact is the whole reason this analyzer exists:
 *
 *   DELETE THE BOOKMARK A CROSS-REFERENCE POINTS AT, AND THE CROSS-REFERENCE KEEPS
 *   DISPLAYING THE OLD TEXT — CORRECTLY FORMATTED, INDEFINITELY, UNTIL SOMEONE PRESSES
 *   F9. THE DOCUMENT LOOKS PERFECT AND IS WRONG.
 *
 * Text extraction reads the cache and reports the stale value. A converter copies it.
 * A reviewer proofreads it. Nothing anywhere says the reference is dead — which is why
 * "see section 4.2" survives in documents whose section 4.2 was deleted years earlier.
 *
 * `@w:dirty="true"` marks a field Word will recalculate when the document next opens,
 * so a stale result there is temporary. **Its absence is the dangerous case**: the
 * cached text is being presented as current. And `@w:fldLock="true"` is worse still —
 * the field will not update even on F9, so the stale value is permanent by instruction.
 *
 * TWO SPELLINGS OF THE SAME THING.
 *
 *   Simple    <w:fldSimple w:instr=" REF Ch1 \h ">…result runs…</w:fldSimple>
 *
 *   Complex   <w:r><w:fldChar w:fldCharType="begin"/></w:r>
 *             <w:r><w:instrText> REF Ch1 \h </w:instrText></w:r>
 *             <w:r><w:fldChar w:fldCharType="separate"/></w:r>
 *             …result runs…
 *             <w:r><w:fldChar w:fldCharType="end"/></w:r>
 *
 * The complex form is three unrelated runs that only become a field by being in the
 * right order, so anything that edits runs can sever one. **`separate` is optional** —
 * a field that has never been calculated goes straight from instruction to `end` — so
 * its absence is not a fault, and treating it as one produces noise on clean documents.
 *
 * ⚠️ **Fields nest.** A `TOC` contains a `PAGEREF` for every entry, inside its own
 * result. A scanner that pairs the first `begin` with the first `end` gets the inner
 * field and mis-reads everything after it, so the walk keeps a stack.
 *
 * Verified against the Open XML SDK schema: `w:fldChar/@w:fldCharType` is **required**
 * and admits exactly `begin`, `separate`, `end`; `w:fldSimple/@w:instr` is **required**;
 * both `w:fldChar` and `w:instrText` are children of `w:r`. Note there are two
 * `fldSimple` declarations — `CT_SimpleField` and `CT_SimpleFieldRuby`, the latter
 * inside `w:ruby` — which is why this matches on element name rather than on the parent.
 */

import { W_NAMESPACE } from './wordStyleResolver';
import { readBookmarks } from './wordBookmarks';
import { finding, renderFindings, type Finding, type Severity } from './findings';

/**
 * Severity and silence per kind.
 *
 * Nearly all of these are SILENT, and that is the point: a field's failure mode is to
 * keep showing its last good answer. `empty-instruction` is the exception — Word prints
 * a visible `Error!` placeholder for that one.
 */
const FIELD_RULES = {
  'stale-reference': { severity: 'error', silent: true },
  'locked-stale-reference': { severity: 'error', silent: true },
  'unbalanced-begin': { severity: 'error', silent: true },
  'unbalanced-end': { severity: 'error', silent: true },
  'orphan-separate': { severity: 'warning', silent: true },
  'empty-instruction': { severity: 'warning', silent: false },
  'no-cached-result': { severity: 'note', silent: false },
  'dirty-result': { severity: 'note', silent: true },
  'locked-field': { severity: 'note', silent: true }
} as const satisfies Record<string, { severity: Severity; silent: boolean }>;

export type FieldProblemKind = keyof typeof FIELD_RULES;

const fieldFinding = (
  kind: FieldProblemKind,
  part: string,
  message: string,
  remediation: string,
  subject?: Record<string, string>
): Finding => finding(`field/${kind}`, part, message, remediation, { ...FIELD_RULES[kind], subject });

/**
 * Field types whose first argument names a bookmark.
 *
 * `HYPERLINK` only does so with the `\l` switch — without it the argument is a URL, and
 * reporting a URL as a missing bookmark would be worse than saying nothing.
 * `STYLEREF` is deliberately absent: its argument is a style name, not a bookmark.
 */
const BOOKMARK_REFERENCING = new Set(['REF', 'PAGEREF', 'NOTEREF']);

export interface ParsedInstruction {
  /** `REF`, `TOC`, … uppercased. null when the instruction is empty. */
  type: string | null;
  /** Non-switch arguments in order. */
  arguments: string[];
  /** Switches including the backslash, e.g. `\h`. */
  switches: string[];
}

/**
 * Splits a field instruction, respecting quoted arguments.
 *
 * `TOC \o "1-3" \h` has a quoted argument containing a space; splitting on whitespace
 * alone turns it into two arguments and loses the range.
 */
export const parseInstruction = (instruction: string): ParsedInstruction => {
  const tokens: string[] = [];
  let current = '';
  let inQuotes = false;
  for (const char of instruction) {
    if (char === '"') {
      inQuotes = !inQuotes;
      continue;
    }
    if (!inQuotes && /\s/.test(char)) {
      if (current) tokens.push(current);
      current = '';
      continue;
    }
    current += char;
  }
  if (current) tokens.push(current);

  const args: string[] = [];
  const switches: string[] = [];
  let type: string | null = null;
  for (const token of tokens) {
    if (token.startsWith('\\')) switches.push(token);
    else if (type === null) type = token.toUpperCase();
    else args.push(token);
  }
  return { type, arguments: args, switches };
};

export interface Field {
  kind: 'simple' | 'complex';
  instruction: string;
  parsed: ParsedInstruction;
  /**
   * The text a reader currently sees. `null` for a complex field with no `separate`,
   * which means it has never been calculated — deliberately distinct from `''`.
   */
  cachedResult: string | null;
  /** Word will recalculate this on open, so a stale result here is temporary. */
  dirty: boolean;
  /** The field will NOT update even on F9. A stale result is permanent by instruction. */
  locked: boolean;
  /** 0 for a top-level field; 1+ for one inside another field's result. */
  depth: number;
}

const isW = (el: Element, local: string) => el.namespaceURI === W_NAMESPACE && el.localName === local;
const attr = (el: Element, name: string) => el.getAttributeNS(W_NAMESPACE, name);
const isOn = (value: string | null) => value === '1' || value === 'true' || value === 'on';

const rootOf = (node: Document | Element): ParentNode =>
  'documentElement' in node && node.documentElement ? node.documentElement : (node as Element);

/** Visible text of a run, excluding instructions and deleted text. */
const runText = (el: Element): string =>
  Array.from(el.getElementsByTagName('*'))
    .filter(child => isW(child, 't'))
    .map(child => child.textContent ?? '')
    .join('');

export interface FieldIndex {
  fields: Field[];
  problems: Finding[];
}

/**
 * Reads every field in a body part.
 *
 * The complex-field walk is a stack rather than a scan, because fields nest: a `TOC`
 * carries a `PAGEREF` per entry inside its own result. Pairing the first `begin` with
 * the first `end` would attribute the outer field's instruction to the inner field's
 * result and mis-read everything after it.
 */
export function readFields(doc: Document | Element, part = ''): FieldIndex {
  const root = rootOf(doc);
  const fields: Field[] = [];
  const problems: Finding[] = [];

  // --- simple fields ---------------------------------------------------------
  for (const el of Array.from(root.querySelectorAll('*'))) {
    if (!isW(el, 'fldSimple')) continue;
    const instruction = attr(el, 'instr') ?? '';
    fields.push({
      kind: 'simple',
      instruction,
      parsed: parseInstruction(instruction),
      // A simple field's content IS its result; there is no separate marker.
      cachedResult: runText(el),
      dirty: isOn(attr(el, 'dirty')),
      locked: isOn(attr(el, 'fldLock')),
      depth: 0
    });
  }

  // --- complex fields --------------------------------------------------------
  interface OpenField {
    instruction: string;
    result: string;
    seenSeparate: boolean;
    dirty: boolean;
    locked: boolean;
  }
  const stack: OpenField[] = [];

  for (const run of Array.from(root.querySelectorAll('*'))) {
    if (!isW(run, 'r')) continue;

    for (const child of Array.from(run.children)) {
      if (isW(child, 'fldChar')) {
        const type = attr(child, 'fldCharType');
        if (type === 'begin') {
          stack.push({
            instruction: '',
            result: '',
            seenSeparate: false,
            dirty: isOn(attr(child, 'dirty')),
            locked: isOn(attr(child, 'fldLock'))
          });
        } else if (type === 'separate') {
          if (stack.length === 0) {
            problems.push(fieldFinding(
              'orphan-separate', part,
              'A <w:fldChar w:fldCharType="separate"/> appears outside any field. It separates a field’s instruction from its result, so on its own it marks nothing and the runs around it are ordinary text that a reader may take for a field result.',
              'Delete the stray separate marker, or restore the begin marker that used to precede it.'
            ));
          } else {
            stack[stack.length - 1].seenSeparate = true;
          }
        } else if (type === 'end') {
          const open = stack.pop();
          if (!open) {
            problems.push(fieldFinding(
              'unbalanced-end', part,
              'A <w:fldChar w:fldCharType="end"/> closes a field that was never opened. Word tolerates this and renders the surrounding text normally, so nothing looks wrong — but the content a reader takes for a field result is plain text that will never update.',
              'Delete the stray end marker, or restore the begin marker that used to match it.'
            ));
          } else {
            fields.push({
              kind: 'complex',
              instruction: open.instruction,
              parsed: parseInstruction(open.instruction),
              cachedResult: open.seenSeparate ? open.result : null,
              dirty: open.dirty,
              locked: open.locked,
              depth: stack.length
            });
          }
        }
        continue;
      }

      // Instruction text belongs to the innermost open field, before its separate.
      if (isW(child, 'instrText') && stack.length > 0) {
        const open = stack[stack.length - 1];
        if (!open.seenSeparate) open.instruction += child.textContent ?? '';
        continue;
      }
    }

    // Result text accrues to every open field that has passed its separate, so a
    // nested field's result is part of its parent's result too - which is what a
    // reader actually sees.
    if (stack.length > 0) {
      const text = runText(run);
      if (text) {
        for (const open of stack) {
          if (open.seenSeparate) open.result += text;
        }
      }
    }
  }

  for (const unclosed of stack) {
    problems.push(fieldFinding(
      'unbalanced-begin', part,
      `A field opens with instruction "${unclosed.instruction.trim() || '(empty)'}" and never closes. Word treats everything after it to the end of the story as part of the field, so text a reader sees as ordinary content is silently inside a field that will not update.`,
      'Insert <w:fldChar w:fldCharType="end"/> where the field should finish.',
      { instruction: unclosed.instruction.trim() }
    ));
  }

  // --- per-field observations ------------------------------------------------
  for (const field of fields) {
    const label = field.parsed.type ?? '(empty)';
    if (field.parsed.type === null) {
      problems.push(fieldFinding(
        'empty-instruction', part,
        'A field carries no instruction, so there is nothing for Word to evaluate. This one is visible: Word prints an error placeholder in the text.',
        'Give the field an instruction, or delete it and keep its current text as ordinary content.'
      ));
      continue;
    }
    if (field.cachedResult === null) {
      problems.push(fieldFinding(
        'no-cached-result', part,
        `The ${label} field has never been calculated — it has no result marker, so it displays nothing until the document is opened and fields are updated.`,
        'Update fields in Word, or accept that this renders as empty in any consumer that does not evaluate fields.',
        { field: label }
      ));
    }
    if (field.locked) {
      problems.push(fieldFinding(
        'locked-field', part,
        `The ${label} field is locked (w:fldLock), so it will not update even when a reader presses F9. Whatever it currently displays is what it will always display.`,
        'Remove w:fldLock if the field is meant to stay current.',
        { field: label }
      ));
    }
    if (field.dirty) {
      problems.push(fieldFinding(
        'dirty-result', part,
        `The ${label} field is marked dirty, so Word will recalculate it when the document opens. The text stored in the file is NOT what a reader will see — anything reading this file without evaluating fields is reading a value Word has already decided is out of date.`,
        'No action needed in Word. A converter or text extractor should be aware the stored result is stale by declaration.',
        { field: label }
      ));
    }
  }

  return { fields, problems };
}

/**
 * Cross-checks every bookmark-referencing field against the bookmarks that exist.
 *
 * This is the finding the module is for, and it needs both halves: the field says which
 * bookmark it points at, and only the bookmark index knows whether that bookmark is
 * still there. Neither analyzer can reach it alone.
 */
export function crossCheckFieldTargets(doc: Document | Element, part = ''): Finding[] {
  const { fields } = readFields(doc, part);
  const bookmarks = readBookmarks(doc, part);
  const names = new Set(bookmarks.bookmarks.map(b => b.name));
  const problems: Finding[] = [];

  for (const field of fields) {
    const type = field.parsed.type;
    if (type === null) continue;

    // HYPERLINK points at a bookmark only with \l; otherwise its argument is a URL.
    const referencesBookmark =
      BOOKMARK_REFERENCING.has(type) || (type === 'HYPERLINK' && field.parsed.switches.includes('\\l'));
    if (!referencesBookmark) continue;

    const target = field.parsed.arguments[0];
    if (target === undefined || names.has(target)) continue;

    const shown = field.cachedResult?.trim();
    const kind: FieldProblemKind = field.locked ? 'locked-stale-reference' : 'stale-reference';

    problems.push(fieldFinding(
      kind, part,
      `The ${type} field points at bookmark "${target}", which does not exist in this part. ` +
        (shown
          ? `It still displays "${shown.slice(0, 80)}${shown.length > 80 ? '…' : ''}" — the answer from the last time the field was calculated. ` +
            'The document reads correctly and the reference is dead. Text extraction returns this stale value; a reviewer proofreads it; nothing anywhere flags it.'
          : 'It has no cached result, so it renders as empty.') +
        (field.locked
          ? ' The field is also locked, so pressing F9 will NOT fix it — the stale value is permanent until the lock is removed.'
          : field.dirty
            ? ' It is marked dirty, so Word will try to recalculate on open and the reader will see an error rather than the stale text.'
            : ' It is not marked dirty, so Word presents the cached text as current and will not recalculate on its own.'),
      field.locked
        ? `Remove w:fldLock and restore bookmark "${target}", or repoint the field at a bookmark that exists.`
        : `Restore bookmark "${target}", or repoint the field at a bookmark that exists.`,
      { field: type, target }
    ));
  }

  return problems;
}

/** Word body parts, where fields live. */
const WORD_BODY = /^word\/(?:document\d*|header[^/]*|footer[^/]*|footnotes\d*|endnotes\d*)\.xml$/;

/**
 * Evidence lines for the AI panel.
 *
 * Leads with the count of fields whose displayed text is a cache, because that framing
 * is what makes the rest make sense.
 */
export function computeFieldEvidenceForMarkup(
  parts: Record<string, string>,
  rawXml: string
): { lines: string[]; unresolved: string[] } | null {
  const entry = Object.entries(parts).find(([path]) => WORD_BODY.test(path));
  if (!entry) return null;

  const [path, xml] = entry;
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) return null;

  const { fields, problems } = readFields(doc, path);
  const crossChecks = crossCheckFieldTargets(doc, path);
  if (fields.length === 0 && problems.length === 0) return null;

  const lines: string[] = [];
  const unresolved: string[] = [];

  const types = [...new Set(fields.map(f => f.parsed.type).filter(Boolean))];
  lines.push(
    `${path} contains ${fields.length} field(s)${types.length > 0 ? ` (${types.join(', ')})` : ''}. ` +
      'What each one displays is the result Word stored the last time it was calculated, not a value computed on open.'
  );

  const nested = fields.filter(f => f.depth > 0).length;
  if (nested > 0) {
    lines.push(`${nested} of them are nested inside another field's result, which is normal for a TOC.`);
  }

  // Describe the selected field when the user has one open.
  const selectedInstr = /w:instr="([^"]*)"/.exec(rawXml)?.[1];
  const selected = selectedInstr
    ? fields.find(f => f.instruction === selectedInstr)
    : undefined;
  if (selected) {
    lines.push(
      `The selected field is ${selected.parsed.type ?? 'untyped'}${
        selected.parsed.arguments.length > 0 ? ` targeting "${selected.parsed.arguments[0]}"` : ''
      }, currently displaying ${selected.cachedResult ? `"${selected.cachedResult.slice(0, 120)}"` : 'nothing'}.`
    );
  }

  lines.push(...renderFindings([...problems, ...crossChecks]));

  unresolved.push(
    `Fields were checked against bookmarks in ${path} only. A reference to a bookmark in another part cannot be resolved from here, so a field reported as pointing at a missing bookmark may be valid across the package.`
  );
  if (fields.some(f => f.parsed.type === 'TOC' || f.parsed.type === 'PAGE' || f.parsed.type === 'NUMPAGES')) {
    unresolved.push(
      'Whether a page-dependent field (TOC, PAGE, NUMPAGES) shows the right number cannot be determined without laying the document out, which is the renderer’s job.'
    );
  }

  return { lines, unresolved };
}
