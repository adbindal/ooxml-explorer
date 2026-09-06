/**
 * PowerPoint faults that render and are broken anyway.
 *
 * `powerpointFormattingAnalysis.ts` answers "what does this shape inherit" for the AI
 * panel. This module asks the other question — "what is wrong with this deck" — and it
 * exists because three resolver capabilities were built and tested and then never
 * reached from production code, which meant the engine could describe these faults on
 * request but never found one on its own.
 *
 * Each check below is here because it satisfies the invisible-failure criterion: the
 * slide draws, PowerPoint raises no repair prompt, and the deck is wrong anyway.
 *
 * | Fault                                | Renders as                          | Severity |
 * |--------------------------------------|-------------------------------------|----------|
 * | notes placeholder with no master     | notes lose inherited formatting     | warning  |
 * | style index past the theme's list    | shape falls back to no fill/line    | warning  |
 * | group with no child coordinate space | children drawn in the wrong place   | warning  |
 *
 * All three are warnings rather than errors. Every one of them is a deck that opens,
 * and calling that an error would train people to ignore the category.
 */

import type { Finding } from './findings';
import type { PackageParts } from './packageIntegrity';
import {
  readPlaceholders,
  matchNotesToMaster,
  readTransform,
  applyGroupTransform,
  resolveStyleReference,
  rotationDegrees,
  P_NAMESPACE,
  A_NAMESPACE
} from './powerpointResolver';

/** Parts this analyzer reads. A deck with no slides has nothing for it to do. */
export const PPT_SLIDE_PART = /^ppt\/slides\/slide[^/]*\.xml$/;
const NOTES_SLIDE_PART = /^ppt\/notesSlides\/notesSlide[^/]*\.xml$/;

const parseXml = (xml: string): Document | null => {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
};

const finding = (
  kind: string,
  part: string,
  message: string,
  remediation: string,
  subject?: Record<string, string>
): Finding => ({
  code: `pptInheritance/${kind}`,
  severity: 'warning',
  part,
  message,
  remediation,
  silent: true,
  ...(subject ? { subject } : {})
});

/**
 * Notes placeholders whose type has no counterpart on the notes master.
 *
 * Notes slides match the master on `@type`, never `@idx` — `matchNotesToMaster` is a
 * separate function from the slide matcher for exactly that reason. When the type is
 * absent from the master the placeholder inherits nothing: the speaker notes still
 * print, in the default font, and the slide-image placeholder shows nothing at all.
 */
const notesFindings = (parts: PackageParts): Finding[] => {
  const masterPath = Object.keys(parts).find(p => /^ppt\/notesMasters\/notesMaster[^/]*\.xml$/.test(p));
  // No notes master is not itself a fault: a deck with no speaker notes has no reason
  // to carry one. Without it there is simply nothing to match against.
  if (!masterPath) return [];

  const masterDoc = parseXml(parts[masterPath] ?? '');
  if (!masterDoc) return [];
  const masterPlaceholders = readPlaceholders(masterDoc);

  const out: Finding[] = [];
  for (const path of Object.keys(parts).filter(p => NOTES_SLIDE_PART.test(p)).sort()) {
    const doc = parseXml(parts[path] ?? '');
    if (!doc) continue;

    for (const placeholder of readPlaceholders(doc)) {
      const match = matchNotesToMaster(placeholder, masterPlaceholders);
      if (match.layoutPlaceholder) continue;

      out.push(finding(
        'notes-placeholder-unmatched', path,
        `The notes placeholder type="${placeholder.type}" has no counterpart on ${masterPath}, so it inherits no formatting. ` +
        `The notes still print — in the default font rather than the one the master defines — and a sldImg placeholder shows no slide thumbnail.`,
        `Add a placeholder of type "${placeholder.type}" to the notes master, or remove the placeholder from the notes slide if it is not wanted.`,
        { placeholderType: placeholder.type, notesMaster: masterPath }
      ));
    }
  }
  return out;
};

/** How many entries a theme format-scheme list actually holds. */
const styleListLength = (theme: Document | null, list: string): number | null => {
  if (!theme) return null;
  const fmtScheme = theme.getElementsByTagNameNS(A_NAMESPACE, 'fmtScheme').item(0);
  if (!fmtScheme) return null;
  const el = Array.from(fmtScheme.children).find(
    c => c.namespaceURI === A_NAMESPACE && c.localName === list
  );
  // An empty list and an absent one are different: absent means the theme does not
  // define that axis at all, and indexing into it is not the document's fault.
  return el ? el.children.length : null;
};

const REF_KINDS = ['fillRef', 'lnRef', 'effectRef', 'bgRef'] as const;

/**
 * Style references that index past the end of the theme list they address.
 *
 * The indexing is not uniform — `fillRef`/`bgRef` reach two lists through one integer
 * with a 1000 offset, `lnRef`/`effectRef` reach one list each with no offset — which is
 * why this defers to `resolveStyleReference` rather than comparing numbers here.
 *
 * Out of range is silent: the shape keeps its explicit `a:solidFill` if it has one and
 * otherwise draws unfilled, so the slide looks plausible and is off-theme.
 */
const styleRefFindings = (parts: PackageParts, themeFor: (slide: string) => Document | null): Finding[] => {
  const out: Finding[] = [];

  for (const path of Object.keys(parts).filter(p => PPT_SLIDE_PART.test(p)).sort()) {
    const doc = parseXml(parts[path] ?? '');
    if (!doc) continue;
    const theme = themeFor(path);

    for (const kind of REF_KINDS) {
      for (const ref of Array.from(doc.getElementsByTagNameNS(A_NAMESPACE, kind))) {
        const idx = Number.parseInt(ref.getAttribute('idx') ?? '', 10);
        if (Number.isNaN(idx)) continue;

        const resolved = resolveStyleReference(kind, idx);
        // A null list is "no fill" (idx 0 or 1000), which is a legitimate choice.
        if (!resolved.list || resolved.position === null) continue;

        const available = styleListLength(theme, resolved.list);
        if (available === null || resolved.position <= available) continue;

        out.push(finding(
          'style-ref-out-of-range', path,
          `A ${kind} asks for entry ${resolved.position} of the theme's ${resolved.list}, which holds ${available}. ` +
          `${resolved.note}. The shape draws with no ${kind === 'lnRef' ? 'outline' : 'fill'} from the theme rather than reporting an error, so the slide looks deliberate and is off-theme.`,
          `Point the ${kind} at an entry the theme defines (1 to ${available}), or add the missing entries to the theme's ${resolved.list}.`,
          { refKind: kind, requested: String(resolved.position), available: String(available), list: resolved.list }
        ));
      }
    }
  }
  return out;
};

/**
 * Groups that position themselves but define no child coordinate space.
 *
 * Children of a group are stored in the group's *child* space, and `a:chOff`/`a:chExt`
 * are what map that space onto the slide. Both are optional in the schema, so a group
 * missing them is valid XML — but the mapping then degenerates to the identity, and
 * every child is plotted at its raw coordinates instead of relative to the group.
 *
 * Office writes both whenever it writes a group transform, so in practice this means
 * the file was produced by something else. The displacement is reported by running the
 * transform the renderer would run, which is what makes the consequence concrete rather
 * than theoretical.
 */
const groupFindings = (parts: PackageParts): Finding[] => {
  const out: Finding[] = [];

  for (const path of Object.keys(parts).filter(p => PPT_SLIDE_PART.test(p)).sort()) {
    const doc = parseXml(parts[path] ?? '');
    if (!doc) continue;

    for (const grpSpPr of Array.from(doc.getElementsByTagNameNS(P_NAMESPACE, 'grpSpPr'))) {
      const group = readTransform(grpSpPr);
      // A group that inherits its geometry has no child space to be missing. That case
      // arrives here as a null offset — `readTransform` returns every field null when
      // there is no a:xfrm — so testing `inherits` as well would be a condition no input
      // can falsify on its own.
      if (!group.offset || !group.extent) continue;
      // Both are needed. One without the other still leaves an axis unmapped, and the
      // half that is present makes the markup look deliberate.
      if (group.childOffset && group.childExtent) continue;

      const groupShape = grpSpPr.parentElement;
      if (!groupShape) continue;

      // Report against a real child, so the message names a shape that actually moves.
      const child = Array.from(groupShape.children).find(
        el => el.namespaceURI === P_NAMESPACE && ['sp', 'pic', 'grpSp', 'graphicFrame'].includes(el.localName)
      );
      if (!child) continue;

      const childXfrm = Array.from(child.getElementsByTagNameNS(A_NAMESPACE, 'xfrm')).at(0) ?? null;
      const childTransform = readTransform(childXfrm?.parentElement ?? null);
      if (!childTransform.offset) continue;

      const mapped = applyGroupTransform(group, childTransform);
      const drift = mapped.offset
        ? Math.abs(mapped.offset.x - childTransform.offset.x) + Math.abs(mapped.offset.y - childTransform.offset.y)
        : 0;
      // No drift means the identity mapping was the right answer anyway.
      if (drift === 0) continue;

      const spin = rotationDegrees(group.rotation);
      out.push(finding(
        'group-child-space-missing', path,
        `A group shape sets a:off and a:ext but ${group.childOffset ? 'no a:chExt' : 'no a:chOff'}, so it declares no child coordinate space. ` +
        `Its children are plotted at their stored coordinates instead of relative to the group — the first child moves by ${drift} EMU once the group transform is applied` +
        `${spin ? `, and the group's ${spin}° rotation turns about the wrong centre` : ''}. ` +
        `Every shape still draws, which is why this survives a visual check.`,
        `Write a:chOff and a:chExt on the group's a:xfrm. Setting them to the group's own a:off and a:ext reproduces the identity mapping explicitly.`,
        { driftEmu: String(drift) }
      ));
    }
  }
  return out;
};

/**
 * Every inheritance finding for a package.
 *
 * The theme lookup is passed in rather than resolved here so that this module does not
 * duplicate `resolveSlideChain`'s relationship walking, which is the part most likely to
 * be wrong and is already tested where it lives.
 */
export function powerpointInheritanceFindings(
  parts: PackageParts,
  themeFor: (slidePath: string) => Document | null
): Finding[] {
  return [...notesFindings(parts), ...styleRefFindings(parts, themeFor), ...groupFindings(parts)];
}
