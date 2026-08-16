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
 * The checks are format-agnostic. Packaging is the one layer Word, Excel and
 * PowerPoint share completely, so this runs unchanged against .docx, .xlsx and .pptx.
 */

export type IntegritySeverity = 'error' | 'warning';

export interface IntegrityFinding {
  severity: IntegritySeverity;
  /** Stable machine-readable identifier, suitable for grouping or suppression. */
  rule:
    | 'missing-content-types'
    | 'untyped-part'
    | 'dangling-relationship-id'
    | 'missing-relationship-target'
    | 'orphaned-rels-part'
    | 'malformed-xml';
  /** The part the finding is about. */
  part: string;
  message: string;
}

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
  target: string;
  external: boolean;
}

const readRelationships = (relsXml: string): Relationship[] | null => {
  const doc = parseXml(relsXml);
  if (!doc) return null;
  return Array.from(doc.getElementsByTagName('Relationship')).map(el => ({
    id: el.getAttribute('Id') ?? '',
    target: el.getAttribute('Target') ?? '',
    external: (el.getAttribute('TargetMode') ?? '').toLowerCase() === 'external'
  }));
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

/**
 * Runs every package-level integrity check and returns the findings, errors first.
 *
 * An empty array means the package's declarations and references are internally
 * consistent. It does not mean the document renders as intended - that is a different
 * question answered by the formatting resolvers, not by this.
 */
export const checkPackageIntegrity = (parts: PackageParts): IntegrityFinding[] => {
  const findings: IntegrityFinding[] = [];
  const paths = Object.keys(parts).map(normalizePath);
  const present = new Set(paths);

  // --- Content types -------------------------------------------------------
  const contentTypesXml = parts[CONTENT_TYPES];
  if (contentTypesXml === undefined) {
    findings.push({
      severity: 'error',
      rule: 'missing-content-types',
      part: CONTENT_TYPES,
      message: `The package has no ${CONTENT_TYPES}. Every OPC package requires one; without it the file cannot be opened.`
    });
  } else {
    const doc = parseXml(contentTypesXml);
    if (!doc) {
      findings.push({
        severity: 'error',
        rule: 'malformed-xml',
        part: CONTENT_TYPES,
        message: `${CONTENT_TYPES} is not well-formed XML.`
      });
    } else {
      const defaults = new Set(
        Array.from(doc.getElementsByTagName('Default'))
          .map(el => (el.getAttribute('Extension') ?? '').toLowerCase())
      );
      const overrides = new Set(
        Array.from(doc.getElementsByTagName('Override'))
          .map(el => normalizePath(el.getAttribute('PartName') ?? ''))
      );

      for (const path of paths) {
        if (path === CONTENT_TYPES) continue; // the stream itself is not a part
        const extension = path.slice(path.lastIndexOf('.') + 1).toLowerCase();
        if (overrides.has(path) || defaults.has(extension)) continue;
        findings.push({
          severity: 'error',
          rule: 'untyped-part',
          part: path,
          message: `Part is not declared in ${CONTENT_TYPES}. Add an Override for it, or a Default for the "${extension}" extension.`
        });
      }
    }
  }

  // --- Relationships -------------------------------------------------------
  for (const path of paths) {
    if (!path.endsWith(RELS_SUFFIX)) continue;

    const owner = ownerOfRels(path);
    if (owner && !present.has(owner)) {
      findings.push({
        severity: 'warning',
        rule: 'orphaned-rels-part',
        part: path,
        message: `Relationship part describes "${owner}", which is not in the package.`
      });
      continue;
    }

    const relationships = readRelationships(parts[path]);
    if (!relationships) {
      findings.push({
        severity: 'error',
        rule: 'malformed-xml',
        part: path,
        message: 'Relationship part is not well-formed XML.'
      });
      continue;
    }

    // Relationship targets are resolved against the *owning part's* directory, not
    // the _rels directory the file physically lives in.
    const resolveBase = owner ?? path.replace(/_rels\/\.rels$/, 'x');
    for (const rel of relationships) {
      if (rel.external || rel.target === '') continue;
      const resolved = resolveTarget(resolveBase, rel.target);
      if (!present.has(resolved)) {
        findings.push({
          severity: 'error',
          rule: 'missing-relationship-target',
          part: owner ?? path,
          message: `Relationship ${rel.id} points at "${rel.target}" (${resolved}), which is not in the package.`
        });
      }
    }
  }

  // --- Dangling relationship references ------------------------------------
  for (const path of paths) {
    if (!isXmlPart(path) || path.endsWith(RELS_SUFFIX) || path === CONTENT_TYPES) continue;

    const doc = parseXml(parts[path]);
    if (!doc) {
      findings.push({
        severity: 'error',
        rule: 'malformed-xml',
        part: path,
        message: 'Part is not well-formed XML.'
      });
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
        findings.push({
          severity: 'error',
          rule: 'dangling-relationship-id',
          part: path,
          message: `References relationship ${id}, which is not declared in ${relsPath}.`
        });
      }
    }
  }

  return findings.sort((a, b) =>
    a.severity === b.severity ? a.part.localeCompare(b.part) : a.severity === 'error' ? -1 : 1
  );
};
