/**
 * Audio and video in PresentationML — the clip, and the still image standing in front of it.
 *
 * A media object on a slide is a `p:pic` like any other picture. Its *picture* content is
 * the poster frame — the still shown before playback — and the clip itself hangs off the
 * non-visual properties as a relationship reference. That split is the whole problem:
 *
 *   THE POSTER FRAME RENDERS WHETHER OR NOT THE MEDIA IS THERE.
 *
 * Drop the video, keep the poster, and the slide is *pixel-identical*. PowerPoint opens
 * the deck without complaint, the thumbnail is right, a screenshot diff is clean. It
 * fails when someone presses play, which is typically in front of an audience. This is
 * the same shape as the OLE preview problem (`oleObjects.ts`), and it is worth checking
 * for exactly the same reason: "it renders correctly" is not evidence the data survived.
 *
 * ⚠️ THE BRIEF'S "r:link vs r:embed" IS NOT THE REAL AXIS — VERIFIED AGAINST THE SCHEMA.
 *
 * `a:videoFile`, `a:audioFile` and `a:quickTimeFile` each carry **`r:link` and nothing
 * else**; `r:link` is REQUIRED and there is no `r:embed` on any of them. So "is this
 * media embedded or linked?" cannot be answered by looking at which attribute was used.
 * It is answered by the RELATIONSHIP: `TargetMode="External"` means the clip lives on the
 * machine that produced the deck, and a relationship without it resolves to a part
 * inside the package. PowerPoint 2010+ routinely writes `a:videoFile r:link="rId2"`
 * where rId2 is an ordinary internal relationship to `ppt/media/media1.mp4` — linked
 * attribute, embedded reality.
 *
 * The one element that really does offer the choice is the Office 2010 extension
 * `p14:media`, which has both `r:embed` and `r:link`. PowerPoint writes it *alongside*
 * `a:videoFile` for the same clip, inside `p:nvPr/p:extLst/p:ext`. Both were checked
 * against the Open XML SDK's published schema data (see PROVENANCE below).
 *
 * WHY AN EXTERNAL LINK IS REPORTED AS AN ERROR AND NOT A STYLE NOTE.
 *
 * Unlike a linked OLE object — where the author chose "link" in a dialog and the markup
 * says so in words — media linking is a *silent default*: PowerPoint links rather than
 * embeds above a size threshold, so the deck stops being self-contained without anyone
 * deciding that it should. The recipient gets a poster frame and a dead play button.
 *
 * PROVENANCE — what is verified and what is not.
 *
 * Verified against the Open XML SDK's machine-readable schema data
 * (`data/schemas/*.json`, `data/namespaces.json`):
 *   • `p:nvPr` children include `a:audioCd`, `a:wavAudioFile`, `a:audioFile`,
 *     `a:videoFile`, `a:quickTimeFile`.
 *   • `a:videoFile` / `a:audioFile` / `a:quickTimeFile` attributes: `r:link` (required).
 *   • `a:wavAudioFile` attribute: `r:embed` (required) — the one media element that is
 *     genuinely embed-only.
 *   • `p14:media` attributes: `r:embed` and `r:link`; `p14` is
 *     `http://schemas.microsoft.com/office/powerpoint/2010/main`.
 *   • `p:pic` children: `p:nvPicPr`, `p:blipFill`, `p:spPr`, `p:style`, `p:extLst`;
 *     `p:nvPicPr` children: `p:cNvPr`, `p:cNvPicPr`, `p:nvPr`.
 *   • `a:blip` attributes: `r:embed`, `r:link`, `cstate`.
 *   • `p:timing` → `p:tnLst` … `p:childTnLst`/`p:subTnLst` → `p:video` / `p:audio` →
 *     `p:cMediaNode` → `p:tgtEl` → `p:spTgt/@spid`.
 *   • `p:cNvPr/@id` is `UInt32`, required.
 *
 * ⚠️ `@contentType` on `a:videoFile`/`a:audioFile` is a DISAGREEMENT between sources.
 * The SDK's schema data lists **only** `r:link` for both — the string `contentType` does
 * not occur anywhere in its DrawingML data. The ECMA-376 5th edition Transitional
 * `dml-main.xsd` *does* declare it: `CT_AudioFile` and `CT_VideoFile` each carry
 * `<xsd:attribute name="contentType" type="xsd:string" use="optional"/>`, while
 * `CT_QuickTimeFile` does not. This code follows ECMA (reads the attribute when present)
 * and never requires it, so the SDK's omission cannot cause a false report either way.
 *
 * ⚠️ NOT VERIFIED: the relationship `Type` URIs for media (the `/video`, `/audio` and
 * Microsoft `/media` type strings). No machine-readable source for them was consulted,
 * so **this module never matches on relationship Type** — it resolves purely by `Id`,
 * exactly as `oleObjects.ts` does. Nothing here depends on a URI that was not checked.
 *
 * ⚠️ NOT VERIFIED: the `p:ext/@uri` GUID that identifies the `p14:media` extension. This
 * code finds `p14:media` by namespace and local name at any depth under `p:nvPr` rather
 * than trusting a GUID that could not be confirmed.
 *
 * NAMESPACES ARE COMPARED BY EXACT EQUALITY. `conformance.ts` rewrites ISO Strict URIs to
 * their Transitional equivalents before any analyzer runs, so Strict-tolerant matching
 * here would be dead code that looks load-bearing. The `p14` namespace is a Microsoft
 * extension and is spelled the same in both conformance classes, so it is exempt from
 * that mapping by nature rather than by omission.
 */

import { relsPathFor, resolveTarget, type PackageParts } from './packageIntegrity';
import { finding, renderFindings, type Finding, type Severity } from './findings';

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P14_NS = 'http://schemas.microsoft.com/office/powerpoint/2010/main';

/**
 * Severity and silence per kind, decided once here.
 *
 * The two VISIBLE kinds are the poster-frame ones: with no still to draw, the slide shows
 * an empty frame or a broken-image box, so a human or a screenshot diff will catch them
 * without help. Everything else leaves the slide looking exactly as intended — which is
 * the entire reason this analyzer exists — and so is marked silent.
 *
 * `external-link` is an `error` rather than a warning because the clip is *not in the
 * package*: for every recipient except the author it is as absent as a deleted part, and
 * unlike a linked OLE object nobody chose it (PowerPoint links by size threshold).
 * `external-link-shadowed` is the same markup with an embedded `p14:media` sibling that
 * does resolve — modern PowerPoint plays the embedded copy, so it is a `note`, not damage.
 * `content-type-mismatch` is a `warning`: consumers overwhelmingly sniff the bytes or
 * trust the file extension, so a disagreeing declaration is a real defect that usually
 * still plays.
 */
const MEDIA_RULES = {
  'no-media-reference':      { severity: 'error',   silent: true },
  'relationship-missing':    { severity: 'error',   silent: true },
  'media-part-missing':      { severity: 'error',   silent: true },
  'external-link':           { severity: 'error',   silent: true },
  'content-type-mismatch':   { severity: 'warning', silent: true },
  'dangling-trigger':        { severity: 'warning', silent: true },
  'poster-part-missing':     { severity: 'warning', silent: false },
  'no-poster-frame':         { severity: 'warning', silent: false },
  'external-link-shadowed':  { severity: 'note',    silent: true }
} as const satisfies Record<string, { severity: Severity; silent: boolean }>;

export type MediaProblemKind = keyof typeof MEDIA_RULES;

const mediaFinding = (
  kind: MediaProblemKind,
  part: string,
  message: string,
  remediation: string,
  subject?: Record<string, string>
): Finding => finding(`media/${kind}`, part, message, remediation, { ...MEDIA_RULES[kind], subject });

export type MediaKind = 'video' | 'audio';

/**
 * The media-bearing children of `p:nvPr`, keyed by local name.
 *
 * All four are DrawingML (`a:`) even though they only ever appear inside PresentationML.
 * `a:wavAudioFile` is the odd one out and the only member of the family whose schema
 * attribute is `r:embed` rather than `r:link`; `a:audioCd` is deliberately absent because
 * it references a physical disc track, not a relationship, so none of the checks here
 * apply to it.
 */
const MEDIA_ELEMENTS: Readonly<Record<string, { attribute: 'link' | 'embed'; label: string; kind: MediaKind }>> = {
  videoFile: { attribute: 'link', label: 'a:videoFile', kind: 'video' },
  audioFile: { attribute: 'link', label: 'a:audioFile', kind: 'audio' },
  quickTimeFile: { attribute: 'link', label: 'a:quickTimeFile', kind: 'video' },
  wavAudioFile: { attribute: 'embed', label: 'a:wavAudioFile', kind: 'audio' }
};

/** One relationship reference to the clip. A picture usually carries two for one clip. */
export interface MediaReference {
  /** Which element declared it, e.g. `a:videoFile` or `p14:media`. */
  element: string;
  /** Which relationship attribute carried it — `r:link` or `r:embed`. */
  attribute: 'r:link' | 'r:embed';
  relationshipId: string | null;
  /** Resolved part path, or the raw URI when the relationship is external. */
  target: string | null;
  isExternal: boolean;
  /** `null` when the target is external, so presence cannot be checked from the package. */
  partExists: boolean | null;
  /** `@contentType`, present on `a:videoFile` and `a:audioFile` only. */
  declaredContentType: string | null;
  /** What `[Content_Types].xml` says the resolved part is, if it says anything. */
  packageContentType: string | null;
}

export interface MediaPoster {
  relationshipId: string | null;
  target: string | null;
  isExternal: boolean;
  partExists: boolean | null;
}

/** A timing-tree node that starts a clip: `p:video` or `p:audio` under `p:timing`. */
export interface MediaTrigger {
  part: string;
  element: Element;
  /** `p:video` or `p:audio`. */
  label: string;
  /** `p:spTgt/@spid` verbatim. */
  targetShapeId: string | null;
  /**
   * Whether the target shape is in this part.
   *
   * `null` means unknowable — see `readMediaTriggers` for the Office 2007 `@spid` form
   * that cannot be matched against a shape id at all.
   */
  targetExists: boolean | null;
  problems: Finding[];
}

export interface SlideMedia {
  part: string;
  /** The owning `p:pic`. */
  element: Element;
  kind: MediaKind;
  /** `p:cNvPr/@id` — what a timing trigger's `@spid` has to match. */
  shapeId: string | null;
  shapeName: string | null;
  /** Every reference to the clip on this picture, in document order. */
  references: MediaReference[];
  /** The still shown before playback: `p:blipFill/a:blip`. */
  poster: MediaPoster | null;
  /** Timing nodes in this part that start this picture. */
  triggers: MediaTrigger[];
  problems: Finding[];
}

interface Relationship {
  id: string;
  target: string;
  external: boolean;
}

const parseXml = (xml: string): Document | null => {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
};

/** Reads an `r:`-namespace attribute without pinning the exact relationships URI. */
const relAttr = (el: Element, local: string): string | null => {
  for (const a of Array.from(el.attributes)) {
    if (a.localName === local && a.namespaceURI?.includes('/relationships')) return a.value;
  }
  return null;
};

const childrenNamed = (parent: Element, ns: string, local: string): Element[] =>
  Array.from(parent.children).filter(c => c.namespaceURI === ns && c.localName === local);

const childNamed = (parent: Element, ns: string, local: string): Element | null =>
  childrenNamed(parent, ns, local)[0] ?? null;

const descendantsNamed = (root: ParentNode, ns: string, local: string): Element[] =>
  Array.from(root.querySelectorAll('*')).filter(el => el.namespaceURI === ns && el.localName === local);

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
      target: rel.getAttribute('Target') ?? '',
      external: (rel.getAttribute('TargetMode') ?? '').toLowerCase() === 'external'
    });
  }
  return map;
};

/**
 * `[Content_Types].xml` lookup, Override before Default.
 *
 * `packageIntegrity.ts` has an equivalent that is module-private, and it answers a
 * different question there: whether a part is declared at all. Undeclared parts are
 * already its finding, so this module deliberately does NOT re-report them — it only
 * compares a declaration that exists against the one the markup claims.
 */
const packageContentTypeOf = (parts: PackageParts, path: string): string | null => {
  const xml = parts['[Content_Types].xml'];
  if (xml === undefined) return null;
  const doc = parseXml(xml);
  if (!doc) return null;

  for (const override of Array.from(doc.getElementsByTagName('Override'))) {
    if ((override.getAttribute('PartName') ?? '').replace(/^\/+/, '') === path) {
      return override.getAttribute('ContentType');
    }
  }
  const extension = path.slice(path.lastIndexOf('.') + 1).toLowerCase();
  for (const dflt of Array.from(doc.getElementsByTagName('Default'))) {
    if ((dflt.getAttribute('Extension') ?? '').toLowerCase() === extension) return dflt.getAttribute('ContentType');
  }
  return null;
};

/** `video/mp4; codecs=avc1` and `VIDEO/MP4` are the same declaration. */
const sameContentType = (a: string, b: string): boolean =>
  a.split(';')[0].trim().toLowerCase() === b.split(';')[0].trim().toLowerCase();

const resolveReference = (
  parts: PackageParts,
  ownerPart: string,
  rels: Map<string, Relationship> | null,
  relId: string | null
): Pick<MediaReference, 'target' | 'isExternal' | 'partExists' | 'packageContentType'> => {
  const missing = { target: null, isExternal: false, partExists: null, packageContentType: null };
  if (relId === null || rels === null) return missing;

  const rel = rels.get(relId);
  if (!rel) return missing;
  // An external target is a path or URL on the authoring machine. Nothing in the package
  // can confirm it resolves, and it will not resolve anywhere else.
  if (rel.external) return { target: rel.target, isExternal: true, partExists: null, packageContentType: null };

  const target = resolveTarget(ownerPart, rel.target);
  return {
    target,
    isExternal: false,
    partExists: parts[target] !== undefined,
    packageContentType: packageContentTypeOf(parts, target)
  };
};

/** Collects the clip references hanging off one picture's `p:nvPr`. */
const readReferences = (
  nvPr: Element,
  parts: PackageParts,
  ownerPart: string,
  rels: Map<string, Relationship> | null
): { references: MediaReference[]; kind: MediaKind | null } => {
  const references: MediaReference[] = [];
  let kind: MediaKind | null = null;

  for (const child of Array.from(nvPr.children)) {
    if (child.namespaceURI !== A_NS) continue;
    const spec = MEDIA_ELEMENTS[child.localName];
    if (spec === undefined) continue;
    kind ??= spec.kind;
    const relId = relAttr(child, spec.attribute);
    references.push({
      element: spec.label,
      attribute: spec.attribute === 'link' ? 'r:link' : 'r:embed',
      relationshipId: relId,
      // `contentType` is unprefixed and optional; absent on a:quickTimeFile by schema.
      declaredContentType: child.getAttribute('contentType'),
      ...resolveReference(parts, ownerPart, rels, relId)
    });
  }

  // The Office 2010 extension, the only reference that genuinely chooses embed or link.
  // Found by namespace and local name rather than by the p:ext/@uri GUID, which no
  // machine-readable source was available to confirm.
  for (const media of descendantsNamed(nvPr, P14_NS, 'media')) {
    const embed = relAttr(media, 'embed');
    const relId = embed ?? relAttr(media, 'link');
    references.push({
      element: 'p14:media',
      attribute: embed === null ? 'r:link' : 'r:embed',
      relationshipId: relId,
      declaredContentType: null,
      ...resolveReference(parts, ownerPart, rels, relId)
    });
  }

  return { references, kind };
};

const readPoster = (
  pic: Element,
  parts: PackageParts,
  ownerPart: string,
  rels: Map<string, Relationship> | null
): MediaPoster | null => {
  const blipFill = childNamed(pic, P_NS, 'blipFill');
  const blip = blipFill ? childNamed(blipFill, A_NS, 'blip') : null;
  if (!blip) return null;

  // a:blip carries r:embed for a packaged image and r:link for an external one.
  const embed = relAttr(blip, 'embed');
  const relId = embed ?? relAttr(blip, 'link');
  const resolved = resolveReference(parts, ownerPart, rels, relId);
  return {
    relationshipId: relId,
    target: resolved.target,
    isExternal: resolved.isExternal,
    // A poster whose reference resolves to nothing — an undeclared id, or no id at all —
    // is a MISSING poster, not an unknowable one. `null` is reserved for the external
    // case, where the image may well be there on some other machine.
    partExists: !resolved.isExternal && resolved.target === null ? false : resolved.partExists
  };
};

/**
 * Every timing node that starts a clip, and whether its target shape is still there.
 *
 * ⚠️ `p:spTgt/@spid` IS NOT ALWAYS A SHAPE ID. The SDK's schema data records two
 * validators for it: a `StringValidator` for Office2007 and a numeric
 * `ST_DrawingElementId` from Office2010 on. `p:cNvPr/@id` is `UInt32` in every version.
 * So a 2007-era deck can legally write `spid="_x0000_s1026"`, which is a VML shape name
 * and matches no `p:cNvPr/@id` by design. Comparing those as strings would report every
 * such deck as having dangling triggers — a confident wrong answer. A non-numeric `@spid`
 * therefore yields `targetExists: null`, and no finding.
 */
export function readMediaTriggers(parts: PackageParts, partPath: string): MediaTrigger[] {
  const xml = parts[partPath];
  if (xml === undefined) return [];
  const doc = parseXml(xml);
  if (!doc?.documentElement) return [];

  const timing = descendantsNamed(doc.documentElement, P_NS, 'timing');
  if (timing.length === 0) return [];

  const shapeIds = new Set(
    descendantsNamed(doc.documentElement, P_NS, 'cNvPr')
      .map(el => el.getAttribute('id'))
      .filter((id): id is string => id !== null)
  );

  const triggers: MediaTrigger[] = [];
  for (const root of timing) {
    for (const node of Array.from(root.querySelectorAll('*'))) {
      if (node.namespaceURI !== P_NS) continue;
      if (node.localName !== 'video' && node.localName !== 'audio') continue;

      const cMediaNode = childNamed(node, P_NS, 'cMediaNode');
      const tgtEl = cMediaNode ? childNamed(cMediaNode, P_NS, 'tgtEl') : null;
      const spTgt = tgtEl ? childNamed(tgtEl, P_NS, 'spTgt') : null;
      const spid = spTgt?.getAttribute('spid') ?? null;

      const numeric = spid !== null && /^\d+$/.test(spid);
      const targetExists = spid === null || !numeric ? null : shapeIds.has(spid);

      const problems: Finding[] = [];
      if (targetExists === false) {
        problems.push(mediaFinding(
          'dangling-trigger', partPath,
          `A p:${node.localName} timing node targets shape id ${spid}, and no shape in ${partPath} has that id. The animation can never fire, so this clip will not play however the slide is advanced — and nothing on the slide looks wrong.`,
          `Point p:spTgt/@spid at the p:cNvPr/@id of the media picture, or delete the timing node.`,
          { spid }
        ));
      }

      triggers.push({
        part: partPath,
        element: node,
        label: `p:${node.localName}`,
        targetShapeId: spid,
        targetExists,
        problems
      });
    }
  }
  return triggers;
}

/** Shared closing checks, run after the markup has been read. */
const finish = (partial: Omit<SlideMedia, 'problems'>, problems: Finding[]): SlideMedia => {
  const part = partial.part;
  const { references, poster } = partial;

  if (references.length === 0 || references.every(r => r.relationshipId === null)) {
    problems.push(mediaFinding(
      'no-media-reference', part,
      'The media picture names no relationship, so there is no clip behind it at all — the poster frame is the entire content of the shape.',
      'Add r:link on the a:videoFile or a:audioFile pointing at the media relationship.'
    ));
  }

  for (const ref of references) {
    if (ref.relationshipId !== null && ref.target === null) {
      problems.push(mediaFinding(
        'relationship-missing', part,
        `${ref.element} references relationship "${ref.relationshipId}", which ${relsPathFor(part)} does not declare. The clip cannot be located, and the poster frame renders regardless.`,
        `Add a Relationship with Id="${ref.relationshipId}" targeting the media part.`,
        { relationshipId: ref.relationshipId }
      ));
    }

    if (ref.partExists === false && ref.target !== null) {
      problems.push(mediaFinding(
        'media-part-missing', part,
        `${ref.element} points at "${ref.target}", which is not in the package. The poster frame still renders, so the slide looks unchanged and the deck opens without warning — it fails only when someone presses play.`,
        `Restore ${ref.target}, or remove the media picture and its relationship together.`,
        { target: ref.target }
      ));
    }

    if (
      ref.declaredContentType !== null &&
      ref.packageContentType !== null &&
      !sameContentType(ref.declaredContentType, ref.packageContentType)
    ) {
      problems.push(mediaFinding(
        'content-type-mismatch', part,
        `${ref.element} declares contentType "${ref.declaredContentType}", but [Content_Types].xml types "${ref.target}" as "${ref.packageContentType}". A consumer that dispatches on the package declaration and one that trusts the markup will pick different decoders for the same bytes.`,
        'Make the two agree — normally by correcting the markup to the part\'s real media type.',
        { declared: ref.declaredContentType, package: ref.packageContentType }
      ));
    }
  }

  // An external reference is only harmless when a sibling reference on the SAME picture
  // resolves to a part that is actually in the package — that is the p14:media embedded
  // copy modern PowerPoint plays in preference to the link.
  const embeddedFallback = references.some(r => r.partExists === true);
  for (const ref of references.filter(r => r.isExternal)) {
    problems.push(
      embeddedFallback
        ? mediaFinding(
            'external-link-shadowed', part,
            `${ref.element} links out to "${ref.target}", which is outside the package, but this picture also carries a reference that resolves inside it. PowerPoint 2010 and later play the packaged copy, so the deck is still self-contained; an older consumer that follows the link is not.`,
            'Nothing to fix for current PowerPoint. Drop the external link if the deck must open in a pre-2010 consumer.',
            { target: ref.target ?? '' }
          )
        : mediaFinding(
            'external-link', part,
            `${ref.element} links out to "${ref.target}" with TargetMode="External". The clip is not in the package: it is a path on the machine that produced the deck, so it resolves nowhere else. The poster frame renders on every machine, so the slide looks complete and playback dies in front of the audience.`,
            'Re-insert the media so PowerPoint embeds it, or ship the linked file alongside the deck at the same path.',
            { target: ref.target ?? '' }
          )
    );
  }

  if (poster === null) {
    problems.push(mediaFinding(
      'no-poster-frame', part,
      'The media picture has no poster image. Unlike a missing clip this one is visible: there is nothing to draw before playback, so the shape appears as empty space.',
      'Add the p:blipFill/a:blip poster frame PowerPoint would normally write for the clip.'
    ));
  } else if (poster.partExists === false) {
    problems.push(mediaFinding(
      'poster-part-missing', part,
      `The poster image at "${poster.target ?? poster.relationshipId ?? 'an unresolvable reference'}" is not in the package, so the frame in front of the clip cannot be drawn. This one shows: the slide renders a broken or empty picture.`,
      'Restore the poster image part, or re-insert the media so PowerPoint regenerates it.'
    ));
  }

  return { ...partial, problems };
};

/**
 * Finds every audio or video picture in one part.
 *
 * `partPath` must be the part the XML came from — relationships resolve relative to it,
 * and a media reference means nothing without it.
 *
 * ⚠️ SCOPE: only `p:pic` owners are read. The schema permits `p:nvPr`, and therefore
 * `a:videoFile`, on `p:nvSpPr`, `p:nvCxnSpPr`, `p:nvGrpSpPr` and `p:nvGraphicFramePr` as
 * well, but no producer was confirmed to write media there and the poster-frame story —
 * the reason this check exists — is specific to `p:pic`. Media on a non-picture shape is
 * therefore NOT reported rather than guessed at.
 */
export function readMedia(parts: PackageParts, partPath: string): SlideMedia[] {
  const xml = parts[partPath];
  if (xml === undefined) return [];
  const doc = parseXml(xml);
  if (!doc?.documentElement) return [];

  const rels = readRelationships(parts, partPath);
  const triggers = readMediaTriggers(parts, partPath);
  const found: SlideMedia[] = [];

  for (const pic of descendantsNamed(doc.documentElement, P_NS, 'pic')) {
    const nvPicPr = childNamed(pic, P_NS, 'nvPicPr');
    const nvPr = nvPicPr ? childNamed(nvPicPr, P_NS, 'nvPr') : null;
    if (!nvPr) continue;

    const { references, kind } = readReferences(nvPr, parts, partPath, rels);
    // A p:pic with no media child is an ordinary picture and none of this module's
    // business. p14:media alone still counts: it is a clip either way.
    if (references.length === 0) continue;

    const cNvPr = nvPicPr ? childNamed(nvPicPr, P_NS, 'cNvPr') : null;
    const shapeId = cNvPr?.getAttribute('id') ?? null;

    found.push(
      finish(
        {
          part: partPath,
          element: pic,
          kind: kind ?? 'video',
          shapeId,
          shapeName: cNvPr?.getAttribute('name') ?? null,
          references,
          poster: readPoster(pic, parts, partPath, rels),
          triggers: shapeId === null ? [] : triggers.filter(t => t.targetShapeId === shapeId)
        },
        []
      )
    );
  }

  return found;
}

/**
 * The clips that render correctly and are broken anyway.
 *
 * The list worth leading with when someone compares a before and after deck: every one of
 * these leaves the slide pixel-identical, so no visual check and no screenshot test will
 * catch them. `silent` carries the whole meaning, so a missing poster frame — visible,
 * and therefore already discoverable — drops out here without being named as an exception.
 */
export function findSilentlyBrokenMedia(media: readonly SlideMedia[]): SlideMedia[] {
  return media.filter(m => m.problems.some(p => p.silent));
}

/**
 * Whether the clip is actually present, stated plainly.
 *
 * `null` means unknowable from the package alone — every reference points outside it.
 * That is deliberately NOT `false`: "we cannot check" and "it is missing" are different
 * answers and only one of them is a defect. A picture with both an external link and a
 * packaged copy is `true`, because a consumer that follows the packaged one succeeds.
 */
export function mediaDataIsPresent(media: SlideMedia): boolean | null {
  if (media.references.some(r => r.partExists === true)) return true;
  if (media.references.some(r => r.partExists === false)) return false;
  return null;
}

/** Every finding this module produces for one part — media problems and timing problems. */
export function mediaFindings(parts: PackageParts, partPath: string): Finding[] {
  return [
    ...readMedia(parts, partPath).flatMap(m => m.problems),
    ...readMediaTriggers(parts, partPath).flatMap(t => t.problems)
  ];
}

/** Parts that can carry a media picture and its timing tree. */
export const MEDIA_HOST_PART = /^ppt\/(?:slides|slideLayouts|slideMasters|notesSlides)\/[^/]+\.xml$/;

/**
 * Evidence lines for the AI panel.
 *
 * Unlike the OLE equivalent this scans for the first host part that actually contains
 * media rather than the first one that merely could. Bundles routinely carry a layout and
 * a master alongside the slide, and key order is insertion order — picking the first match
 * blind would report "no media" for a deck that has some, depending on how the bundle was
 * assembled.
 */
export function computeMediaEvidenceForMarkup(
  parts: Record<string, string>
): { lines: string[]; unresolved: string[] } | null {
  const hosts = Object.keys(parts).filter(path => MEDIA_HOST_PART.test(path));

  let hostPath: string | null = null;
  let media: SlideMedia[] = [];
  let triggers: MediaTrigger[] = [];
  for (const path of hosts) {
    const found = readMedia(parts, path);
    if (found.length === 0) continue;
    hostPath = path;
    media = found;
    triggers = readMediaTriggers(parts, path);
    break;
  }
  if (hostPath === null) return null;

  const lines: string[] = [];
  const unresolved: string[] = [];

  lines.push(`${hostPath} contains ${media.length} audio or video picture(s).`);

  for (const clip of media) {
    const named = clip.shapeName === null ? 'an unnamed shape' : `"${clip.shapeName}"`;
    const how = clip.references.map(r => `${r.element}/${r.attribute}`).join(' and ');
    lines.push(
      `A ${clip.kind} clip on ${named} (shape id ${clip.shapeId ?? 'unstated'}), referenced by ${how || 'nothing'}, ` +
        `with its data at ${clip.references.find(r => r.target !== null)?.target ?? 'no resolvable target'}.`
    );

    lines.push(
      clip.poster === null
        ? 'It has no poster frame, so the absence of the clip would be visible.'
        : 'A poster frame stands in front of it, which is why a missing clip would not change how the slide looks.'
    );

    if (mediaDataIsPresent(clip) === null) {
      lines.push(
        'Whether the clip resolves cannot be determined from the package: every reference is external, so the file lives on the machine that produced the deck.'
      );
      for (const external of clip.references.filter(r => r.isExternal)) {
        unresolved.push(`The external target "${external.target}" cannot be checked from inside the package.`);
      }
    }

    lines.push(...renderFindings(clip.problems));
  }

  lines.push(...renderFindings(triggers.flatMap(t => t.problems)));

  const silent = findSilentlyBrokenMedia(media);
  if (silent.length > 0) {
    lines.push(
      `${silent.length} of these clip(s) will render exactly as intended and are broken anyway — the poster frame is intact while the media behind it is not, so no visual check will catch this before someone presses play.`
    );
  }

  // The relationship resolves and the part is there. Whether the bytes decode, and
  // whether a given consumer has the codec, are different questions this cannot answer.
  if (media.some(m => mediaDataIsPresent(m) === true)) {
    unresolved.push(
      'The media parts were confirmed present but never decoded, so whether they are playable — valid bytes, a codec the consumer has — is unverified.'
    );
  }

  if (triggers.some(t => t.targetExists === null && t.targetShapeId !== null)) {
    unresolved.push(
      'A timing trigger uses the Office 2007 string form of p:spTgt/@spid, which cannot be matched against p:cNvPr/@id, so whether it still points at a live shape is unverified.'
    );
  }

  return { lines, unresolved };
}
