/**
 * OPC package integrity checks.
 *
 * These are deterministic checks over a parsed package - no model, no retrieval, no
 * judgement. Every finding is computed, so an answer built on one can be presented as
 * verified rather than merely grounded.
 *
 * They target the failure class that produces Word's "found unreadable content" dialog
 * rather than a wrong-looking render. Schema validity does not catch any of it: a
 * document can be perfectly valid WordprocessingML and still fail to open because a
 * part was never declared in [Content_Types].xml or an r:id points at nothing.
 *
 * The packaging checks are format-agnostic. Packaging is the one layer Word, Excel and
 * PowerPoint share completely, so they run unchanged against .docx, .xlsx and .pptx.
 *
 * The implicit-relationship checks are the exception, and deliberately so: they catch
 * a *required relationship that is simply absent*, which no reference-following check
 * can reach, and knowing which relationships are required is per-format knowledge. It
 * lives in a content-type-keyed table (IMPLICIT_RELATIONSHIPS) so it stays data.
 */

import { finding, type Finding, type Severity } from './findings';

export type IntegritySeverity = 'error' | 'warning';

/**
 * Severity, silence and the fix for each rule.
 *
 * `remediation` lives here rather than at each call site because it is a property of
 * the rule, not of the occurrence — every dangling relationship id is fixed the same
 * way. This also closes a real gap: integrity findings previously carried no
 * remediation at all, which made them the one analyzer that could tell you something
 * was wrong without telling you what to do about it.
 *
 * Most of these are SILENT. A package with a missing implicit relationship or an
 * untyped part still opens and still renders, which is precisely why it needs a
 * checker. Only a missing [Content_Types].xml or malformed XML stops the file loading.
 */
const INTEGRITY_RULES = {
  'missing-content-types': {
    severity: 'error', silent: false,
    remediation: 'Add a [Content_Types].xml declaring a Default or Override for every part.'
  },
  'untyped-part': {
    severity: 'warning', silent: true,
    remediation: 'Add an Override for the part, or a Default for its extension, in [Content_Types].xml.'
  },
  'dangling-relationship-id': {
    severity: 'error', silent: true,
    remediation: 'Declare the relationship in the part’s .rels, or remove the reference to it.'
  },
  'missing-relationship-target': {
    severity: 'error', silent: true,
    remediation: 'Restore the target part, or delete the relationship pointing at it.'
  },
  'missing-implicit-relationship': {
    severity: 'error', silent: true,
    remediation: 'Add the missing relationship to the part’s .rels.'
  },
  'ambiguous-implicit-relationship': {
    severity: 'error', silent: true,
    remediation: 'Remove all but the one correct relationship.'
  },
  'orphaned-rels-part': {
    severity: 'warning', silent: true,
    remediation: 'Delete the relationship part, or restore the part it describes.'
  },
  'malformed-xml': {
    severity: 'error', silent: false,
    remediation: 'Repair the XML so that it parses.'
  }
} as const satisfies Record<string, { severity: Severity; silent: boolean; remediation: string }>;

export type IntegrityRule = keyof typeof INTEGRITY_RULES;

const integrityFinding = (rule: IntegrityRule, part: string, message: string): Finding =>
  finding(`package/${rule}`, part, message, INTEGRITY_RULES[rule].remediation, {
    severity: INTEGRITY_RULES[rule].severity,
    silent: INTEGRITY_RULES[rule].silent
  });

/** Part path (no leading slash) to its text content. Binary parts map to ''. */
export type PackageParts = Record<string, string>;

const CONTENT_TYPES = '[Content_Types].xml';
const RELS_SUFFIX = '.rels';

/** OPC stores paths without a leading slash; Override/@PartName carries one. */
const normalizePath = (path: string): string => path.replace(/^\/+/, '');

/** `word/document.xml` -> `word/_rels/document.xml.rels` */
export const relsPathFor = (partPath: string): string => {
  const slash = partPath.lastIndexOf('/');
  const dir = slash === -1 ? '' : partPath.slice(0, slash + 1);
  const file = slash === -1 ? partPath : partPath.slice(slash + 1);
  return `${dir}_rels/${file}${RELS_SUFFIX}`;
};

/** `word/_rels/document.xml.rels` -> `word/document.xml` */
const ownerOfRels = (relsPath: string): string | null => {
  const match = relsPath.match(/^(.*?)_rels\/(.+)\.rels$/);
  if (!match) return null;
  const [, dir, file] = match;
  // The package-level `_rels/.rels` describes the package itself, which has no part.
  return file === '' ? null : `${dir}${file}`;
};

/**
 * Resolves a relationship Target against the owning part's directory.
 *
 * Targets are routinely relative and routinely climb: an image referenced from
 * `word/document.xml` is typically `media/image1.png`, while one referenced from
 * `word/header1.xml` may be `../media/image1.png`.
 */
export const resolveTarget = (ownerPart: string, target: string): string => {
  if (target.startsWith('/')) return normalizePath(target);
  const slash = ownerPart.lastIndexOf('/');
  const baseSegments = slash === -1 ? [] : ownerPart.slice(0, slash).split('/');
  const segments = [...baseSegments];
  for (const segment of target.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') segments.pop();
    else segments.push(segment);
  }
  return segments.join('/');
};

const parseXml = (xml: string): Document | null => {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
};

interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

const readRelationships = (relsXml: string): Relationship[] | null => {
  const doc = parseXml(relsXml);
  if (!doc) return null;
  return Array.from(doc.getElementsByTagName('Relationship')).map(el => ({
    id: el.getAttribute('Id') ?? '',
    type: el.getAttribute('Type') ?? '',
    target: el.getAttribute('Target') ?? '',
    external: (el.getAttribute('TargetMode') ?? '').toLowerCase() === 'external'
  }));
};

/** Content type declarations, split the two ways [Content_Types].xml expresses them. */
interface ContentTypeDeclarations {
  /** Lowercased file extension -> content type. */
  defaults: Map<string, string>;
  /** Normalized part path -> content type. */
  overrides: Map<string, string>;
}

const readContentTypes = (doc: Document): ContentTypeDeclarations => ({
  defaults: new Map(
    Array.from(doc.getElementsByTagName('Default')).map(el => [
      (el.getAttribute('Extension') ?? '').toLowerCase(),
      el.getAttribute('ContentType') ?? ''
    ])
  ),
  overrides: new Map(
    Array.from(doc.getElementsByTagName('Override')).map(el => [
      normalizePath(el.getAttribute('PartName') ?? ''),
      el.getAttribute('ContentType') ?? ''
    ])
  )
});

/**
 * The content type of a part, by the package's own declaration.
 *
 * Parts are identified by content type rather than by filename because filenames are
 * only conventional - `ppt/slides/slide1.xml` is a habit, not a rule, and a generator
 * is free to name it anything. The content type is what a consumer actually dispatches
 * on, so a check keyed on it stays correct for packages that do not follow the habit.
 */
const contentTypeOf = (
  path: string,
  declarations: ContentTypeDeclarations
): string | undefined => {
  const override = declarations.overrides.get(path);
  if (override !== undefined) return override.toLowerCase();
  const extension = path.slice(path.lastIndexOf('.') + 1).toLowerCase();
  return declarations.defaults.get(extension)?.toLowerCase();
};

/**
 * Collects relationship ids referenced by a part.
 *
 * Rather than enumerating the attributes that can hold one (`r:id`, `r:embed`,
 * `r:link`, `r:href`, `r:dm`, `r:lo`, `r:qs`, `r:cs`, `r:pict`, and more across the
 * three formats), this takes any attribute in the relationships namespace whose value
 * looks like a relationship id. Enumerating would silently miss the format-specific
 * ones - `r:embed` on a Word image, `r:id` on an Excel hyperlink, `r:embed` on a chart
 * reference - and a missed reference is an integrity hole, not a cosmetic gap.
 */
export const collectRelationshipRefs = (doc: Document): Set<string> => {
  const ids = new Set<string>();
  const walk = (element: Element) => {
    for (const attr of Array.from(element.attributes)) {
      if (attr.name.startsWith('r:') && /^rId\d+$/i.test(attr.value)) {
        ids.add(attr.value);
      }
    }
    for (const child of Array.from(element.children)) walk(child);
  };
  if (doc.documentElement) walk(doc.documentElement);
  return ids;
};

const isXmlPart = (path: string): boolean =>
  path.endsWith('.xml') || path.endsWith(RELS_SUFFIX);

// --- Implicit relationships ------------------------------------------------

const PML = 'application/vnd.openxmlformats-officedocument.presentationml';

export interface ImplicitRelationship {
  /** `[Content_Types].xml` content type of the part that must carry the relationship. */
  partContentType: string;
  /** Trailing segment of the expected relationship `Type` URI, e.g. `slideLayout`. */
  relationshipType: string;
  /** `exactly-one` additionally makes a second relationship of the type an error. */
  cardinality: 'exactly-one' | 'at-least-one';
  /** Reader-facing name of the part, for messages. */
  partLabel: string;
  /** What silently breaks when the relationship is absent. */
  consequence: string;
}

/**
 * Relationships that ECMA-376 requires but that nothing in the XML references.
 *
 * PresentationML carries its whole inheritance chain this way (§13.3.9, §13.3.10): a
 * slide finds its layout, a layout its master, a master its theme, and a notes slide
 * its notes master by opening the part's `.rels` and looking for the relationship of
 * the right *type*. There is no `r:id` in `slide1.xml` naming the layout - the
 * relationship is the only link.
 *
 * That is why `dangling-relationship-id` and `missing-relationship-target` cannot
 * cover this: both start from a reference in the XML and follow it outward. When the
 * relationship is simply absent there is no reference to start from, nothing dangles,
 * and every existing check passes. PowerPoint then opens the file without a
 * complaint, silently drops the slide's inherited placeholders, positions, fonts and
 * colours, and renders defaults. A wrong-looking deck with a clean bill of health is
 * the exact failure this table exists to catch.
 *
 * Keyed by content type so DOCX/XLSX entries can be added here as data rather than as
 * new branches.
 */
export const IMPLICIT_RELATIONSHIPS: readonly ImplicitRelationship[] = [
  {
    partContentType: `${PML}.slide+xml`,
    relationshipType: 'slideLayout',
    cardinality: 'exactly-one',
    partLabel: 'slide',
    consequence:
      'the slide inherits no placeholders, geometry or theme and renders with built-in defaults'
  },
  {
    partContentType: `${PML}.slideLayout+xml`,
    relationshipType: 'slideMaster',
    cardinality: 'exactly-one',
    partLabel: 'slide layout',
    consequence: 'the layout inherits nothing, so every slide using it loses the master'
  },
  {
    partContentType: `${PML}.slideMaster+xml`,
    relationshipType: 'theme',
    cardinality: 'exactly-one',
    partLabel: 'slide master',
    consequence: 'theme colours, fonts and effect styles resolve to defaults everywhere below it'
  },
  {
    partContentType: `${PML}.notesSlide+xml`,
    relationshipType: 'notesMaster',
    cardinality: 'exactly-one',
    partLabel: 'notes slide',
    consequence: 'speaker notes lose their master formatting and placeholder layout'
  }
];

/**
 * Matches a relationship `Type` URI on its trailing segment rather than in full.
 *
 * The same relationship has two spellings: Transitional packages use
 * `http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout`
 * and ISO Strict packages use
 * `http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout`. An equality
 * test against either constant would report every package written in the other
 * flavour as broken, so the stable part - the final segment - is what we compare.
 * The `/relationships/` guard keeps an unrelated vendor URI that happens to end in
 * the same word from matching.
 */
const hasRelationshipType = (relationshipTypeUri: string, expected: string): boolean => {
  const uri = relationshipTypeUri.toLowerCase();
  if (!uri.includes('/relationships/')) return false;
  return uri.slice(uri.lastIndexOf('/') + 1) === expected.toLowerCase();
};

/**
 * Runs every package-level integrity check and returns the findings, errors first.
 *
 * An empty array means the package's declarations and references are internally
 * consistent. It does not mean the document renders as intended - that is a different
 * question answered by the formatting resolvers, not by this.
 */
export const checkPackageIntegrity = (parts: PackageParts): Finding[] => {
  const findings: Finding[] = [];
  const paths = Object.keys(parts).map(normalizePath);
  const present = new Set(paths);
  let declarations: ContentTypeDeclarations | null = null;

  // --- Content types -------------------------------------------------------
  const contentTypesXml = parts[CONTENT_TYPES];
  if (contentTypesXml === undefined) {
    findings.push(integrityFinding(
      'missing-content-types',
      CONTENT_TYPES,
      `The package has no ${CONTENT_TYPES}. Every OPC package requires one; without it the file cannot be opened.`
    ));
  } else {
    const doc = parseXml(contentTypesXml);
    if (!doc) {
      findings.push(integrityFinding(
        'malformed-xml',
        CONTENT_TYPES,
        `${CONTENT_TYPES} is not well-formed XML.`
      ));
    } else {
      declarations = readContentTypes(doc);
      const { defaults, overrides } = declarations;

      for (const path of paths) {
        if (path === CONTENT_TYPES) continue; // the stream itself is not a part
        const extension = path.slice(path.lastIndexOf('.') + 1).toLowerCase();
        if (overrides.has(path) || defaults.has(extension)) continue;
        findings.push(integrityFinding(
          'untyped-part',
          path,
          `Part is not declared in ${CONTENT_TYPES}. Add an Override for it, or a Default for the "${extension}" extension.`
        ));
      }
    }
  }

  // --- Relationships -------------------------------------------------------
  for (const path of paths) {
    if (!path.endsWith(RELS_SUFFIX)) continue;

    const owner = ownerOfRels(path);
    if (owner && !present.has(owner)) {
      findings.push(integrityFinding(
        'orphaned-rels-part',
        path,
        `Relationship part describes "${owner}", which is not in the package.`
      ));
      continue;
    }

    const relationships = readRelationships(parts[path]);
    if (!relationships) {
      findings.push(integrityFinding(
        'malformed-xml',
        path,
        'Relationship part is not well-formed XML.'
      ));
      continue;
    }

    // Relationship targets are resolved against the *owning part's* directory, not
    // the _rels directory the file physically lives in. The package-level
    // `_rels/.rels` has no owning part, so its targets resolve from the package root.
    const resolveBase = owner ?? '';
    for (const rel of relationships) {
      if (rel.external || rel.target === '') continue;
      const resolved = resolveTarget(resolveBase, rel.target);
      if (!present.has(resolved)) {
        findings.push(integrityFinding(
          'missing-relationship-target',
          owner ?? path,
          `Relationship ${rel.id} points at "${rel.target}" (${resolved}), which is not in the package.`
        ));
      }
    }
  }

  // --- Dangling relationship references ------------------------------------
  for (const path of paths) {
    if (!isXmlPart(path) || path.endsWith(RELS_SUFFIX) || path === CONTENT_TYPES) continue;

    const doc = parseXml(parts[path]);
    if (!doc) {
      findings.push(integrityFinding(
        'malformed-xml',
        path,
        'Part is not well-formed XML.'
      ));
      continue;
    }

    const referenced = collectRelationshipRefs(doc);
    if (referenced.size === 0) continue;

    // Each part carries its own relationships. An image used by header1.xml must be
    // declared in word/_rels/header1.xml.rels - a declaration in document.xml.rels
    // does not satisfy it, and assuming otherwise is a classic source of packages
    // that open on some readers and fail on Word.
    const relsPath = relsPathFor(path);
    const declared = new Set(
      (present.has(relsPath) ? readRelationships(parts[relsPath]) ?? [] : []).map(r => r.id)
    );

    for (const id of referenced) {
      if (!declared.has(id)) {
        findings.push(integrityFinding(
          'dangling-relationship-id',
          path,
          `References relationship ${id}, which is not declared in ${relsPath}.`
        ));
      }
    }
  }

  // --- Implicit relationships ----------------------------------------------
  // Needs the content type declarations; if [Content_Types].xml is missing or
  // unparseable that is already reported above, and every part's identity is unknown,
  // so there is nothing meaningful to say here.
  if (declarations) {
    for (const path of paths) {
      if (path === CONTENT_TYPES || path.endsWith(RELS_SUFFIX)) continue;

      const contentType = contentTypeOf(path, declarations);
      if (contentType === undefined) continue;

      const expectations = IMPLICIT_RELATIONSHIPS.filter(
        e => e.partContentType.toLowerCase() === contentType
      );
      if (expectations.length === 0) continue;

      const relsPath = relsPathFor(path);
      const relationships = present.has(relsPath) ? readRelationships(parts[relsPath]) : [];
      // A malformed .rels is reported as malformed-xml above. Reading "no slideLayout
      // relationship" out of XML we could not parse would be a second, misleading
      // finding for the same underlying problem.
      if (relationships === null) continue;

      for (const expectation of expectations) {
        const matches = relationships.filter(rel =>
          hasRelationshipType(rel.type, expectation.relationshipType)
        );

        if (matches.length === 0) {
          findings.push(integrityFinding(
            'missing-implicit-relationship',
            path,
            `This ${expectation.partLabel} has no "${expectation.relationshipType}" relationship in ${relsPath}. The link is implicit - nothing in the part's XML references it - so no other check can see it missing, and the file still opens: ${expectation.consequence}.`
          ));
          continue;
        }

        if (expectation.cardinality === 'exactly-one' && matches.length > 1) {
          const ids = matches.map(rel => rel.id).join(', ');
          findings.push(integrityFinding(
            'ambiguous-implicit-relationship',
            path,
            `This ${expectation.partLabel} declares ${matches.length} "${expectation.relationshipType}" relationships (${ids}) in ${relsPath}, but may have exactly one. Because the link is implicit, the part's XML says nothing about which is intended and consumers are free to disagree; remove all but the correct one.`
          ));
        }
      }
    }
  }

  return findings.sort((a, b) =>
    a.severity === b.severity ? a.part.localeCompare(b.part) : a.severity === 'error' ? -1 : 1
  );
};
