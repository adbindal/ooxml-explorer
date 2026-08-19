/**
 * Content controls and their data bindings — the document that looks filled in and is not.
 *
 * A structured document tag (`w:sdt`) is the mechanism behind every template-driven
 * document pipeline: a placeholder in the file, bound to a value in a custom XML part,
 * which Word substitutes when the document opens. It is how mail-merge-style generation
 * works in modern Word, and it fails in a way that is very hard to see:
 *
 *   THE CONTROL'S CONTENT IS STORED IN THE FILE. IF THE BINDING BREAKS, THE CONTENT
 *   STAYS. A document whose bindings all point at a custom XML part that is no longer
 *   there opens showing the values from the last time it worked — or the placeholder
 *   text — and looks completely normal.
 *
 * For anyone generating documents this is the difference between "the template
 * populated" and "the template printed last month's numbers again". Nothing in the
 * rendered page distinguishes them.
 *
 * THE BINDING CHAIN.
 *
 *   <w:sdt><w:sdtPr>
 *     <w:dataBinding w:xpath="/root/customer/name" w:storeItemID="{GUID}"/>
 *   </w:sdtPr><w:sdtContent>…the stored value…</w:sdtContent></w:sdt>
 *                                  │
 *   @w:storeItemID ────────────────┘  matches ds:datastoreItem/@ds:itemID in
 *                                     customXml/itemProps<N>.xml, whose sibling
 *                                     customXml/item<N>.xml holds the actual data
 *
 * Both `@w:xpath` and `@w:storeItemID` are **required** — verified against the SDK
 * schema — so a binding missing either is malformed rather than partial.
 *
 * `w:showingPlcHdr` is the other signal worth reading. It means the control is
 * currently displaying its *placeholder*, not data. In a template that is correct; in a
 * document that has supposedly been generated it means **this field was never filled**,
 * and the grey prompt text is what a reader will see.
 *
 * ⚠️ **The XPath is not evaluated here.** Checking that `/root/customer/name` actually
 * selects a node needs a namespace-aware XPath engine and the `@w:prefixMappings`
 * bindings; getting that subtly wrong would produce confident false reports about
 * working templates. So this module verifies that the *part* the binding names exists
 * and says plainly that whether the expression matches anything inside it is unchecked.
 * A missing part is the common failure by a wide margin.
 *
 * Verified against the Open XML SDK schema: `w:dataBinding` declares `@w:xpath` and
 * `@w:storeItemID` as required plus optional `@w:prefixMappings`; `w:sdtPr` may carry
 * `w:alias`, `w:tag`, `w:id`, `w:lock`, `w:placeholder`, `w:showingPlcHdr` and
 * `w:temporary`; and there are three `w:sdt` declarations — `CT_SdtBlock`, `CT_SdtRun`
 * and `CT_SdtRunRuby` — which is why this matches on element name rather than parent.
 */

import { W_NAMESPACE } from './wordStyleResolver';
import { finding, renderFindings, type Finding, type Severity } from './findings';
import type { PackageParts } from './packageIntegrity';

/** The custom XML datastore namespace, which carries the item id a binding resolves to. */
export const DS_NAMESPACE = 'http://schemas.openxmlformats.org/officeDocument/2006/customXml';

/**
 * Severity and silence per kind.
 *
 * A broken binding is silent by construction — the stored content renders. A control
 * showing its placeholder is the one visible case, and even then only to someone who
 * knows the grey prompt text is not data.
 */
const SDT_RULES = {
  'binding-part-missing': { severity: 'error', silent: true },
  'binding-incomplete': { severity: 'error', silent: true },
  'showing-placeholder': { severity: 'warning', silent: false },
  'duplicate-id': { severity: 'warning', silent: true },
  'no-content-element': { severity: 'error', silent: false },
  'unbound-control': { severity: 'note', silent: true }
} as const satisfies Record<string, { severity: Severity; silent: boolean }>;

export type ContentControlProblemKind = keyof typeof SDT_RULES;

const sdtFinding = (
  kind: ContentControlProblemKind,
  part: string,
  message: string,
  remediation: string,
  subject?: Record<string, string>
): Finding => finding(`contentControl/${kind}`, part, message, remediation, { ...SDT_RULES[kind], subject });

export interface DataBinding {
  xpath: string | null;
  storeItemID: string | null;
  prefixMappings: string | null;
}

export interface ContentControl {
  /** `w:alias` — the friendly name shown in Word's UI. */
  alias: string | null;
  /** `w:tag` — the machine-readable name a generator keys on. */
  tag: string | null;
  id: string | null;
  binding: DataBinding | null;
  /** True when the control is displaying its placeholder rather than data. */
  showingPlaceholder: boolean;
  /** The text currently stored in the control. */
  content: string;
  /** `w:lock`: `sdtLocked`, `contentLocked`, `sdtContentLocked`, `unlocked`. */
  lock: string | null;
}

const isW = (el: Element, local: string) => el.namespaceURI === W_NAMESPACE && el.localName === local;
const attr = (el: Element, name: string) => el.getAttributeNS(W_NAMESPACE, name);
const childOf = (parent: Element, local: string) => Array.from(parent.children).find(c => isW(c, local)) ?? null;

const rootOf = (node: Document | Element): ParentNode =>
  'documentElement' in node && node.documentElement ? node.documentElement : (node as Element);

const parseXml = (xml: string | undefined): Document | null => {
  if (xml === undefined) return null;
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
};

/** Visible text inside an element, ignoring instruction and deleted text. */
const textOf = (el: Element): string =>
  Array.from(el.getElementsByTagName('*'))
    .filter(c => isW(c, 't'))
    .map(c => c.textContent ?? '')
    .join('');

/**
 * Every `ds:itemID` the package's custom XML parts declare, normalised.
 *
 * GUIDs are compared case-insensitively and without braces, because generators are
 * inconsistent about both and a binding that differs only in case is not broken —
 * reporting it as broken would be a false alarm on working documents.
 */
export function readDataStoreIds(parts: PackageParts): Set<string> {
  const ids = new Set<string>();
  for (const [path, xml] of Object.entries(parts)) {
    if (!/^customXml\/itemProps\d*\.xml$/.test(path)) continue;
    const doc = parseXml(xml);
    const root = doc?.documentElement;
    if (!root) continue;
    for (const el of [root, ...Array.from(root.querySelectorAll('*'))]) {
      if (el.namespaceURI !== DS_NAMESPACE || el.localName !== 'datastoreItem') continue;
      const id = el.getAttributeNS(DS_NAMESPACE, 'itemID');
      if (id) ids.add(normaliseItemId(id));
    }
  }
  return ids;
}

/** Braces and case are not significant when matching a store item id. */
export const normaliseItemId = (id: string): string => id.replace(/[{}]/g, '').toLowerCase();

export interface ContentControlIndex {
  controls: ContentControl[];
  problems: Finding[];
}

/**
 * Reads every content control in a body part.
 *
 * `storeIds` is the set of item ids the package actually declares; pass it to have
 * bindings checked. Omit it and the structural checks still run, which is what happens
 * when only one part is available.
 */
export function readContentControls(
  doc: Document | Element,
  part = '',
  storeIds?: Set<string>
): ContentControlIndex {
  const root = rootOf(doc);
  const controls: ContentControl[] = [];
  const problems: Finding[] = [];
  const seenIds = new Map<string, string>();

  for (const sdt of Array.from(root.querySelectorAll('*'))) {
    if (!isW(sdt, 'sdt')) continue;

    const props = childOf(sdt, 'sdtPr');
    const contentEl = childOf(sdt, 'sdtContent');
    const bindingEl = props ? childOf(props, 'dataBinding') : null;

    const binding: DataBinding | null = bindingEl
      ? {
          xpath: attr(bindingEl, 'xpath'),
          storeItemID: attr(bindingEl, 'storeItemID'),
          prefixMappings: attr(bindingEl, 'prefixMappings')
        }
      : null;

    const aliasEl = props ? childOf(props, 'alias') : null;
    const tagEl = props ? childOf(props, 'tag') : null;
    const idEl = props ? childOf(props, 'id') : null;
    const lockEl = props ? childOf(props, 'lock') : null;

    const control: ContentControl = {
      alias: aliasEl ? attr(aliasEl, 'val') : null,
      tag: tagEl ? attr(tagEl, 'val') : null,
      id: idEl ? attr(idEl, 'val') : null,
      binding,
      showingPlaceholder: props !== null && childOf(props, 'showingPlcHdr') !== null,
      content: contentEl ? textOf(contentEl) : '',
      lock: lockEl ? attr(lockEl, 'val') : null
    };
    controls.push(control);

    const label = control.alias ?? control.tag ?? control.id ?? 'an unnamed control';

    if (!contentEl) {
      problems.push(sdtFinding(
        'no-content-element', part,
        `Content control "${label}" has no w:sdtContent, so it wraps nothing. Word shows an empty control where the value belongs.`,
        'Add a w:sdtContent element containing the control’s current text.',
        { control: label }
      ));
    }

    if (control.showingPlaceholder) {
      problems.push(sdtFinding(
        'showing-placeholder', part,
        `Content control "${label}" is marked as showing its placeholder, so what appears in the document is prompt text rather than data${control.content ? ` — currently "${control.content.slice(0, 60)}"` : ''}. In a blank template that is correct; in a document that has supposedly been generated it means this field was never filled.`,
        'Populate the control and remove w:showingPlcHdr, or confirm this file is a template rather than a generated document.',
        { control: label }
      ));
    }

    if (control.id !== null) {
      const previous = seenIds.get(control.id);
      if (previous !== undefined) {
        problems.push(sdtFinding(
          'duplicate-id', part,
          `Two content controls share w:id "${control.id}" ("${previous}" and "${label}"). Word uses the id to address a control, so anything driving this document by id reaches only one of them and the other is never updated.`,
          'Renumber one of the controls to an id unused elsewhere in the part.',
          { id: control.id, control: label }
        ));
      } else {
        seenIds.set(control.id, label);
      }
    }

    if (binding === null) {
      problems.push(sdtFinding(
        'unbound-control', part,
        `Content control "${label}" has no data binding, so nothing populates it automatically — its text is whatever was typed or written into the file.`,
        'No action needed if the control is meant to be filled by hand. Add a w:dataBinding if it should be driven from custom XML.',
        { control: label }
      ));
      continue;
    }

    if (binding.xpath === null || binding.storeItemID === null) {
      problems.push(sdtFinding(
        'binding-incomplete', part,
        `The data binding on "${label}" is missing ${binding.xpath === null ? 'w:xpath' : 'w:storeItemID'}. Both are required, so the binding cannot resolve and the control keeps whatever text it already holds.`,
        'Supply both w:xpath and w:storeItemID, or remove the binding and treat the control as manually filled.',
        { control: label }
      ));
      continue;
    }

    if (storeIds && !storeIds.has(normaliseItemId(binding.storeItemID))) {
      problems.push(sdtFinding(
        'binding-part-missing', part,
        `Content control "${label}" is bound to custom XML store item ${binding.storeItemID}, which no part in this package declares. The binding cannot resolve, so the control displays the text already stored in the file${control.content ? ` — currently "${control.content.slice(0, 60)}"` : ''} — and keeps displaying it. A generated document in this state looks populated and is not connected to anything.`,
        `Restore the customXml part declaring itemID ${binding.storeItemID}, or repoint the binding at a store item that exists.`,
        { control: label, storeItemID: binding.storeItemID }
      ));
    }
  }

  return { controls, problems };
}

/** Word body parts, where content controls live. */
export const SDT_HOST_PART = /^word\/(?:document\d*|header[^/]*|footer[^/]*|footnotes\d*|endnotes\d*)\.xml$/;

/** Every content-control finding across the package. */
export function contentControlFindings(parts: PackageParts): Finding[] {
  const storeIds = readDataStoreIds(parts);
  return Object.keys(parts)
    .filter(p => SDT_HOST_PART.test(p))
    .flatMap(path => {
      const doc = parseXml(parts[path]);
      return doc ? readContentControls(doc, path, storeIds).problems : [];
    });
}

/**
 * Evidence lines for the AI panel.
 *
 * Leads with how many controls are bound versus manual, because that single ratio is
 * what tells someone whether they are looking at a template or a finished document.
 */
export function computeContentControlEvidenceForMarkup(
  parts: Record<string, string>,
  rawXml: string
): { lines: string[]; unresolved: string[] } | null {
  const path = Object.keys(parts).find(p => SDT_HOST_PART.test(p));
  if (path === undefined) return null;
  const doc = parseXml(parts[path]);
  if (!doc) return null;

  const storeIds = readDataStoreIds(parts);
  const { controls, problems } = readContentControls(doc, path, storeIds);
  if (controls.length === 0) return null;

  const lines: string[] = [];
  const unresolved: string[] = [];

  const bound = controls.filter(c => c.binding !== null).length;
  lines.push(
    `${path} contains ${controls.length} content control(s): ${bound} bound to custom XML, ${controls.length - bound} filled by hand. ` +
      'A bound control stores its current text in the document, so it keeps displaying that text whether or not the binding still resolves.'
  );

  const placeholders = controls.filter(c => c.showingPlaceholder).length;
  if (placeholders > 0) {
    lines.push(`${placeholders} control(s) are showing placeholder text rather than data.`);
  }

  // Name the selected control when the user has one open.
  const selectedTag = /<w:tag[^>]*w:val="([^"]*)"/.exec(rawXml)?.[1];
  const selected = selectedTag ? controls.find(c => c.tag === selectedTag) : undefined;
  if (selected) {
    lines.push(
      `The selected control is tagged "${selected.tag}"${selected.alias ? ` (shown as "${selected.alias}")` : ''}, ` +
        `${selected.binding ? `bound to ${selected.binding.xpath}` : 'not bound to anything'}, ` +
        `currently containing ${selected.content ? `"${selected.content.slice(0, 80)}"` : 'no text'}.`
    );
  }

  lines.push(...renderFindings(problems));

  // The limit this module deliberately does not cross.
  if (bound > 0) {
    unresolved.push(
      'Whether each binding’s XPath expression selects anything inside the custom XML part was not evaluated — that needs a namespace-aware XPath engine and the binding’s prefixMappings. Only the existence of the part it names was checked.'
    );
  }

  return { lines, unresolved };
}
