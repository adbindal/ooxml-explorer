/**
 * OLE objects — the embedded thing, and the picture standing in front of it.
 *
 * An OLE object is a foreign document (a spreadsheet in a report, a CAD drawing in a
 * deck) carried inside the package as an opaque binary. Because no OOXML consumer can
 * render a foreign binary, every OLE object ships with a **preview image** — and that
 * is where the interesting failure lives:
 *
 *   THE PREVIEW RENDERS WHETHER OR NOT THE OBJECT DATA IS THERE.
 *
 * Drop the embedding, keep the preview, and the page looks *pixel-identical*. Word and
 * PowerPoint open the file without complaint. The document only breaks when someone
 * double-clicks to edit, which may be months later and is nobody's regression test.
 * "It renders correctly" is not evidence the object survived, so a renderer or
 * converter has to check the relationship target, not the appearance.
 *
 * THE SAME CONCEPT, EXPRESSED THREE INCOMPATIBLE WAYS.
 *
 * All three formats embed OLE, and no two agree on how to say "embedded or linked" —
 * which matters directly when translating one spec into another:
 *
 *   Word         w:object > o:OLEObject/@Type = "Embed" | "Link"     an ATTRIBUTE
 *                preview: the sibling VML v:shape > v:imagedata/@r:id
 *
 *   PowerPoint   p:oleObj > p:embed | p:link                         a CHILD ELEMENT
 *                preview: p:oleObj > p:pic (a real DrawingML picture)
 *
 *   Excel        x:oleObject/@link = a formula reference             ATTRIBUTE PRESENCE
 *                preview: a VML shape in the sheet's legacy drawing
 *
 * So a converter cannot ask one question of all three. In PowerPoint the binding is a
 * choice between two elements; in Excel it is inferred from whether an attribute
 * exists at all; only Word states it in words.
 *
 * ⚠️ `o:OLEObject` ATTRIBUTES ARE UNPREFIXED AND PascalCase — `Type`, `ProgID`,
 * `ShapeID`, `DrawAspect`, `ObjectID`, `UpdateMode` — against the lowerCamelCase
 * convention the rest of OOXML follows. `getAttributeNS(officeNs, 'type')` finds
 * nothing twice over: wrong case, and the attributes are in no namespace at all.
 * Verified against the Open XML SDK schema for `urn:schemas-microsoft-com:office:office`.
 *
 * `@progId` names the application that owns the binary (`Excel.Sheet.12`,
 * `Package` for an embedded OOXML document). It is the only clue to what the object
 * *is* — the binary itself is an OLE compound file this code does not open.
 */

import { relsPathFor, resolveTarget, type PackageParts } from './packageIntegrity';

/**
 * Namespace matching tolerates Strict as well as Transitional packages, which use
 * different URIs (`schemas.openxmlformats.org/...` vs `purl.oclc.org/ooxml/...`) for
 * the same vocabulary. Comparing whole URIs reports every Strict package as broken.
 */
const nsEndsWith = (uri: string | null, suffix: string) => uri !== null && uri.endsWith(suffix);

const isWordprocessing = (el: Element) => nsEndsWith(el.namespaceURI, '/wordprocessingml/2006/main');
const isPresentation = (el: Element) => nsEndsWith(el.namespaceURI, '/presentationml/2006/main');
const isSpreadsheet = (el: Element) => nsEndsWith(el.namespaceURI, '/spreadsheetml/2006/main');

const OFFICE_NS = 'urn:schemas-microsoft-com:office:office';
const VML_NS = 'urn:schemas-microsoft-com:vml';

/** Reads an `r:`-namespace attribute without pinning the exact relationships URI. */
const relAttr = (el: Element, local: string): string | null => {
  for (const a of Array.from(el.attributes)) {
    if (a.localName === local && a.namespaceURI?.includes('/relationships')) return a.value;
  }
  return null;
};

const descendants = (root: ParentNode, predicate: (el: Element) => boolean): Element[] =>
  Array.from(root.querySelectorAll('*')).filter(predicate);

export type OleFormat = 'word' | 'excel' | 'powerpoint';

/** How the object is bound to its data. `unknown` means the markup did not say. */
export type OleBinding = 'embedded' | 'linked' | 'unknown';

export type OleProblemKind =
  | 'no-data-reference'
  | 'relationship-missing'
  | 'data-part-missing'
  | 'binding-mismatch'
  | 'no-preview'
  | 'unknown-binding'
  | 'no-prog-id';

export interface OleProblem {
  kind: OleProblemKind;
  message: string;
  remediation: string;
  /** True when the page still renders correctly despite this problem. */
  silent: boolean;
}

export interface OlePreview {
  /** 'vml' in Word and Excel, 'drawingml' in PowerPoint. */
  kind: 'vml' | 'drawingml';
  relationshipId: string | null;
  target: string | null;
  partExists: boolean | null;
}

export interface OleObject {
  format: OleFormat;
  element: Element;
  binding: OleBinding;
  /** Where the binding was read from — differs per format, so state it. */
  bindingEvidence: string;
  progId: string | null;
  shapeId: string | null;
  /** The relationship naming the object data. */
  dataRelationshipId: string | null;
  /** Resolved part path, or the external URI for a linked object. */
  dataTarget: string | null;
  dataIsExternal: boolean;
  /** null when the target is external and so not checkable from the package. */
  dataPartExists: boolean | null;
  preview: OlePreview | null;
  problems: OleProblem[];
}

interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

const parseXml = (xml: string): Document | null => {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
};

const readRelationships = (parts: PackageParts, ownerPart: string): Map<string, Relationship> | null => {
  const relsXml = parts[relsPathFor(ownerPart)];
  if (relsXml === undefined) return null;
  const doc = parseXml(relsXml);
  if (!doc) return null;

  const map = new Map<string, Relationship>();
  for (const rel of Array.from(doc.getElementsByTagName('Relationship'))) {
    const id = rel.getAttribute('Id');
    if (id === null) continue;
    map.set(id, {
      id,
      type: rel.getAttribute('Type') ?? '',
      target: rel.getAttribute('Target') ?? '',
      external: rel.getAttribute('TargetMode') === 'External'
    });
  }
  return map;
};

/**
 * Resolves the relationship carrying an object's data and says, precisely, which link
 * in the chain broke. Each break leaves the page rendering exactly as before.
 */
const resolveData = (
  parts: PackageParts,
  ownerPart: string,
  rels: Map<string, Relationship> | null,
  relId: string | null,
  problems: OleProblem[]
): Pick<OleObject, 'dataTarget' | 'dataIsExternal' | 'dataPartExists'> => {
  if (relId === null) {
    problems.push({
      kind: 'no-data-reference',
      message: 'The object names no relationship, so it has no data at all — there is nothing for a double-click to open.',
      remediation: 'Add an r:id pointing at an oleObject relationship.',
      silent: true
    });
    return { dataTarget: null, dataIsExternal: false, dataPartExists: null };
  }

  if (rels === null) {
    problems.push({
      kind: 'relationship-missing',
      message: `${ownerPart} has no relationship part, so relationship "${relId}" cannot be resolved and the object data cannot be found.`,
      remediation: `Create ${relsPathFor(ownerPart)} with a relationship for "${relId}".`,
      silent: true
    });
    return { dataTarget: null, dataIsExternal: false, dataPartExists: null };
  }

  const rel = rels.get(relId);
  if (!rel) {
    problems.push({
      kind: 'relationship-missing',
      message: `Relationship "${relId}" is referenced by the object but not declared in ${relsPathFor(ownerPart)}.`,
      remediation: `Add a Relationship with Id="${relId}" targeting the embedded object part.`,
      silent: true
    });
    return { dataTarget: null, dataIsExternal: false, dataPartExists: null };
  }

  if (rel.external) {
    // An external target is a path or URL on the authoring machine. Nothing in the
    // package can confirm it resolves, and it usually will not on another machine.
    return { dataTarget: rel.target, dataIsExternal: true, dataPartExists: null };
  }

  const target = resolveTarget(ownerPart, rel.target);
  const exists = parts[target] !== undefined;
  if (!exists) {
    problems.push({
      kind: 'data-part-missing',
      message: `The object points at "${target}", which is not in the package. The preview image still renders, so the page looks unchanged and the file opens without warning — the object only fails when someone double-clicks it to edit.`,
      remediation: `Restore ${target}, or remove the object and its relationship together.`,
      silent: true
    });
  }
  return { dataTarget: target, dataIsExternal: false, dataPartExists: exists };
};

/** Word: `w:object` wrapping a VML shape and an `o:OLEObject`. */
const readWordObject = (
  wObject: Element,
  parts: PackageParts,
  ownerPart: string,
  rels: Map<string, Relationship> | null
): OleObject => {
  const problems: OleProblem[] = [];
  const ole = descendants(wObject, el => el.namespaceURI === OFFICE_NS && el.localName === 'OLEObject')[0] ?? null;

  // Unprefixed, PascalCase — see the file header.
  const type = ole?.getAttribute('Type') ?? null;
  const binding: OleBinding = type === 'Embed' ? 'embedded' : type === 'Link' ? 'linked' : 'unknown';
  if (binding === 'unknown') {
    problems.push({
      kind: 'unknown-binding',
      message: `o:OLEObject/@Type is ${type === null ? 'absent' : `"${type}"`}; it must be "Embed" or "Link", so whether this object carries its data or points elsewhere is undetermined.`,
      remediation: 'Set Type="Embed" for an embedded object or Type="Link" for a linked one.',
      silent: true
    });
  }

  const relId = ole ? relAttr(ole, 'id') : null;
  const data = resolveData(parts, ownerPart, rels, relId, problems);

  // The preview is the VML shape's image, a sibling of o:OLEObject inside w:object.
  const imagedata = descendants(wObject, el => el.namespaceURI === VML_NS && el.localName === 'imagedata')[0] ?? null;
  const previewRelId = imagedata ? relAttr(imagedata, 'id') : null;
  const preview = imagedata
    ? {
        kind: 'vml' as const,
        relationshipId: previewRelId,
        ...resolvePreviewTarget(parts, ownerPart, rels, previewRelId)
      }
    : null;

  return finish(
    {
      format: 'word',
      element: wObject,
      binding,
      bindingEvidence: 'o:OLEObject/@Type',
      progId: ole?.getAttribute('ProgID') ?? null,
      shapeId: ole?.getAttribute('ShapeID') ?? null,
      dataRelationshipId: relId,
      ...data,
      preview
    },
    problems
  );
};

/** PowerPoint: `p:oleObj`, where the binding is a child element, not an attribute. */
const readPresentationObject = (
  oleObj: Element,
  parts: PackageParts,
  ownerPart: string,
  rels: Map<string, Relationship> | null
): OleObject => {
  const problems: OleProblem[] = [];
  const children = Array.from(oleObj.children).filter(isPresentation);
  const hasEmbed = children.some(c => c.localName === 'embed');
  const hasLink = children.some(c => c.localName === 'link');

  const binding: OleBinding = hasEmbed && !hasLink ? 'embedded' : hasLink && !hasEmbed ? 'linked' : 'unknown';
  if (binding === 'unknown') {
    problems.push({
      kind: 'unknown-binding',
      message:
        hasEmbed && hasLink
          ? 'p:oleObj declares both p:embed and p:link. They are alternatives, not options, so the object is both embedded and linked and consumers need not agree on which wins.'
          : 'p:oleObj declares neither p:embed nor p:link. In PresentationML the binding is a choice of child element, so with neither present the object states no binding at all.',
      remediation: 'Keep exactly one of p:embed or p:link.',
      silent: true
    });
  }

  const relId = relAttr(oleObj, 'id');
  const data = resolveData(parts, ownerPart, rels, relId, problems);

  // p:pic is a full DrawingML picture, not a placeholder — which is exactly why a
  // slide with no object data behind it looks perfectly normal.
  const pic = children.find(c => c.localName === 'pic') ?? null;
  const blip = pic
    ? descendants(pic, el => nsEndsWith(el.namespaceURI, '/drawingml/2006/main') && el.localName === 'blip')[0] ?? null
    : null;
  const previewRelId = blip ? relAttr(blip, 'embed') : null;
  const preview = pic
    ? {
        kind: 'drawingml' as const,
        relationshipId: previewRelId,
        ...resolvePreviewTarget(parts, ownerPart, rels, previewRelId)
      }
    : null;

  return finish(
    {
      format: 'powerpoint',
      element: oleObj,
      binding,
      bindingEvidence: 'p:embed / p:link child element',
      progId: oleObj.getAttribute('progId'),
      shapeId: oleObj.getAttribute('spid'),
      dataRelationshipId: relId,
      ...data,
      preview
    },
    problems
  );
};

/** Excel: `x:oleObject`, where linkage is inferred from an attribute's presence. */
const readSpreadsheetObject = (
  oleObject: Element,
  parts: PackageParts,
  ownerPart: string,
  rels: Map<string, Relationship> | null
): OleObject => {
  const problems: OleProblem[] = [];
  // @link holds a formula referring to the source. Its presence *is* the statement
  // that the object is linked; there is no attribute that says "embedded".
  const binding: OleBinding = oleObject.getAttribute('link') === null ? 'embedded' : 'linked';

  const relId = relAttr(oleObject, 'id');
  const data = resolveData(parts, ownerPart, rels, relId, problems);

  return finish(
    {
      format: 'excel',
      element: oleObject,
      binding,
      bindingEvidence: '@link attribute presence',
      progId: oleObject.getAttribute('progId'),
      shapeId: oleObject.getAttribute('shapeId'),
      dataRelationshipId: relId,
      // Excel keeps the preview in the sheet's legacy VML drawing, a separate part
      // reached through the worksheet's legacyDrawing relationship rather than from
      // the oleObject element. Not resolved here; reported as unknown, not absent.
      preview: null,
      ...data
    },
    problems
  );
};

const resolvePreviewTarget = (
  parts: PackageParts,
  ownerPart: string,
  rels: Map<string, Relationship> | null,
  relId: string | null
): { target: string | null; partExists: boolean | null } => {
  if (relId === null || rels === null) return { target: null, partExists: null };
  const rel = rels.get(relId);
  if (!rel) return { target: null, partExists: false };
  if (rel.external) return { target: rel.target, partExists: null };
  const target = resolveTarget(ownerPart, rel.target);
  return { target, partExists: parts[target] !== undefined };
};

/** Shared closing checks, run after every format-specific read. */
const finish = (partial: Omit<OleObject, 'problems'>, problems: OleProblem[]): OleObject => {
  if (partial.progId === null) {
    problems.push({
      kind: 'no-prog-id',
      message:
        'No progId, so nothing identifies which application owns the embedded binary. A consumer cannot tell an embedded spreadsheet from an embedded drawing without opening the compound file.',
      remediation: 'Set progId to the owning application, e.g. "Excel.Sheet.12" or "Package" for an embedded OOXML document.',
      silent: true
    });
  }

  // Only meaningful for the formats whose preview is reachable from the element.
  if (partial.format !== 'excel' && partial.preview === null) {
    problems.push({
      kind: 'no-preview',
      message:
        'The object has no preview image. Unlike a missing embedding this one is visible: consumers that cannot execute OLE have nothing to draw, so the object renders as blank space or an error box.',
      remediation: 'Add the preview image the producing application would normally write alongside the object.',
      silent: false
    });
  }

  if (partial.binding === 'linked' && partial.dataTarget !== null && !partial.dataIsExternal) {
    problems.push({
      kind: 'binding-mismatch',
      message: `The object declares itself linked, but its relationship resolves inside the package ("${partial.dataTarget}") instead of carrying TargetMode="External". A linked object points at a file outside the document; a relationship without TargetMode does not.`,
      remediation: 'Either set TargetMode="External" on the relationship, or change the binding to embedded.',
      silent: true
    });
  }

  if (partial.binding === 'embedded' && partial.dataIsExternal) {
    problems.push({
      kind: 'binding-mismatch',
      message: `The object declares itself embedded, but its relationship is TargetMode="External" and points at "${partial.dataTarget}" — a path on the machine that produced the file. The data is not in the package, so it will not resolve anywhere else.`,
      remediation: 'Embed the object data as a package part, or change the binding to linked.',
      silent: true
    });
  }

  return { ...partial, problems };
};

/**
 * Finds every OLE object in one part, whichever format it belongs to.
 *
 * `partPath` must be the part the XML came from — relationships resolve relative to it,
 * and an OLE object's data reference is meaningless without it.
 */
export function readOleObjects(parts: PackageParts, partPath: string): OleObject[] {
  const xml = parts[partPath];
  if (xml === undefined) return [];
  const doc = parseXml(xml);
  if (!doc?.documentElement) return [];

  const rels = readRelationships(parts, partPath);
  const root = doc.documentElement;
  const found: OleObject[] = [];

  for (const el of descendants(root, e => isWordprocessing(e) && e.localName === 'object')) {
    found.push(readWordObject(el, parts, partPath, rels));
  }
  for (const el of descendants(root, e => isPresentation(e) && e.localName === 'oleObj')) {
    found.push(readPresentationObject(el, parts, partPath, rels));
  }
  for (const el of descendants(root, e => isSpreadsheet(e) && e.localName === 'oleObject')) {
    found.push(readSpreadsheetObject(el, parts, partPath, rels));
  }

  return found;
}

/**
 * The objects that render correctly and are broken anyway.
 *
 * This is the list worth showing a user comparing a "before" and "after" file: every
 * one of these produces an identical-looking page, so no visual check and no rendering
 * test will catch them.
 */
export function findSilentlyBrokenOleObjects(objects: OleObject[]): OleObject[] {
  return objects.filter(o => o.problems.some(p => p.silent && p.kind !== 'no-prog-id'));
}

/** Body parts that can carry an OLE object, in any of the three formats. */
const OLE_HOST_PART =
  /^(?:word\/(?:document\d*|header[^/]*|footer[^/]*|footnotes\d*|endnotes\d*)\.xml|xl\/worksheets\/[^/]+\.xml|ppt\/slides\/[^/]+\.xml)$/;

/**
 * Evidence lines for the AI panel, format-agnostic.
 *
 * The point of surfacing this is that nothing else will: an object with no data behind
 * it renders identically to one that is intact, so neither the user nor a screenshot
 * test can see the difference.
 */
export function computeOleEvidenceForMarkup(
  parts: Record<string, string>
): { lines: string[]; unresolved: string[] } | null {
  const hostPath = Object.keys(parts).find(path => OLE_HOST_PART.test(path));
  if (hostPath === undefined) return null;

  const objects = readOleObjects(parts, hostPath);
  if (objects.length === 0) return null;

  const lines: string[] = [];
  const unresolved: string[] = [];

  lines.push(`${hostPath} contains ${objects.length} OLE object(s).`);

  for (const object of objects) {
    const what = object.progId ?? 'an unidentified application';
    lines.push(
      `An OLE object owned by ${what}, declared ${object.binding} (read from ${object.bindingEvidence}), ` +
        `with its data at ${object.dataTarget ?? 'no resolvable target'}.`
    );

    const present = oleDataIsPresent(object);
    if (present === null) {
      lines.push(
        'Whether the data resolves cannot be determined from the package: the target is external, so it lives on the machine that produced the file.'
      );
      unresolved.push(`The external target "${object.dataTarget}" cannot be checked from inside the package.`);
    }

    for (const problem of object.problems) lines.push(`${problem.message} ${problem.remediation}`);
  }

  const silent = findSilentlyBrokenOleObjects(objects);
  if (silent.length > 0) {
    lines.push(
      `${silent.length} of these object(s) will render exactly as intended and are broken anyway — the preview image is intact while the embedded data is not, so no visual check will catch this.`
    );
  }

  // The relationship resolves and the part exists; whether the bytes are a valid
  // compound file for the declared progId is a different question and not one this
  // code can answer.
  if (objects.some(o => o.dataPartExists === true)) {
    unresolved.push(
      'The embedded binaries were confirmed present but not opened, so whether their contents match the declared progId is unverified.'
    );
  }

  return { lines, unresolved };
}

/**
 * Whether an object's data is actually present, stated plainly.
 *
 * `null` means unknowable from the package alone — a linked object's target lives
 * outside it. That is deliberately not `false`: "we cannot check" and "it is missing"
 * are different answers, and only one of them is a defect.
 */
export function oleDataIsPresent(object: OleObject): boolean | null {
  if (object.dataIsExternal) return null;
  return object.dataPartExists === true;
}
