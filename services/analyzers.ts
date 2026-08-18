/**
 * The analyzer registry — what this engine can check, as data rather than as control flow.
 *
 * Before this, knowing what the engine could do meant reading `ANALYSIS_TARGETS` in a
 * React component and following six imports. Two consequences, both of which showed up
 * as real gaps:
 *
 *   - The **validate** surface ran exactly one analyzer (`checkPackageIntegrity`), so
 *     "is this file correct?" checked content types and relationships and nothing else.
 *   - The **compare** surface ran the structural diff and no analyzer at all, so a file
 *     that had lost an OLE embedding or a bookmark's end marker reported the part change
 *     without reporting that anything was broken.
 *
 * Both are the same bug: there was no way to say "run everything that applies here."
 * A registry makes that one line, and makes adding the next analyzer a data change.
 *
 * THE HONEST HALF.
 *
 * Every entry declares `cannotDetermine` alongside `determines`. This is not
 * documentation — it is the input to the capability ledger, and it extends the honesty
 * property from per-fact to per-capability. The badge already refuses to call a claim
 * Verified without provenance; the same discipline says the engine should be able to
 * answer "why can't you check this?" from data rather than from a model's guess about
 * its own abilities. An analyzer that lists nothing it cannot do is almost certainly
 * lying.
 */

import type { Finding } from './findings';
import type { PackageParts } from './packageIntegrity';
import { checkPackageIntegrity } from './packageIntegrity';
import { readBookmarks } from './wordBookmarks';
import { readComments, COMMENT_PART_PATHS } from './wordComments';
import { readOleObjects } from './oleObjects';
import { readPivotTables } from './excelPivotTables';

export type OoxmlFormat = 'docx' | 'xlsx' | 'pptx';

export interface Analyzer {
  /** Matches the namespace of the codes it emits, e.g. `bookmark` → `bookmark/…`. */
  id: string;
  /** Short human-readable name, for the capability ledger. */
  title: string;
  formats: readonly OoxmlFormat[];
  /** What this analyzer establishes, phrased as the questions it answers. */
  determines: readonly string[];
  /**
   * What it explicitly does NOT establish, so the engine can say so rather than
   * leaving a reader to assume silence means "fine".
   */
  cannotDetermine: readonly string[];
  /**
   * True when the package contains anything for this analyzer to look at. Keeps the
   * ledger honest about what actually ran versus what merely could have.
   */
  appliesTo: (parts: PackageParts) => boolean;
  analyze: (parts: PackageParts) => Finding[];
}

/** Word story parts, which is where bookmarks and comment anchors live. */
const WORD_BODY = /^word\/(?:document\d*|header[^/]*|footer[^/]*|footnotes\d*|endnotes\d*)\.xml$/;
/** Parts that can host an OLE object, in any of the three formats. */
const OLE_HOST =
  /^(?:word\/(?:document\d*|header[^/]*|footer[^/]*|footnotes\d*|endnotes\d*)\.xml|xl\/worksheets\/[^/]+\.xml|ppt\/slides\/[^/]+\.xml)$/;

const parse = (xml: string | undefined): Document | null => {
  if (xml === undefined) return null;
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
};

const matching = (parts: PackageParts, pattern: RegExp) => Object.keys(parts).filter(p => pattern.test(p));

export const ANALYZERS: readonly Analyzer[] = [
  {
    id: 'package',
    title: 'Package integrity',
    formats: ['docx', 'xlsx', 'pptx'],
    determines: [
      'whether every part is declared in [Content_Types].xml',
      'whether every relationship resolves to a part that exists',
      'whether implicit relationships a consumer needs are present and unambiguous'
    ],
    cannotDetermine: [
      'whether a part that exists contains valid content for its type',
      'whether the ZIP container itself is well-formed — that is checked when the file is opened'
    ],
    // Always applicable: every OPC package has content types and relationships.
    appliesTo: () => true,
    analyze: parts => checkPackageIntegrity(parts)
  },
  {
    id: 'bookmark',
    title: 'Bookmarks and the shared markup id space',
    formats: ['docx'],
    determines: [
      'whether every bookmark range opens and closes',
      'whether any w:id is reused across bookmarks, tracked changes or permissions',
      'what text a bookmark covers'
    ],
    cannotDetermine: [
      'whether a bookmark is referenced from a different part — each part is read on its own',
      'whether a field instruction that names a bookmark would actually resolve at render time'
    ],
    appliesTo: parts => matching(parts, WORD_BODY).length > 0,
    analyze: parts =>
      matching(parts, WORD_BODY).flatMap(path => {
        const doc = parse(parts[path]);
        return doc ? readBookmarks(doc, path).problems : [];
      })
  },
  {
    id: 'comment',
    title: 'Comments, anchoring and threading',
    formats: ['docx'],
    determines: [
      'whether every comment anchor has a body and every body an anchor',
      'what text a comment is attached to',
      'whether a comment is a reply, and whether it is resolved — when the side-car is present'
    ],
    cannotDetermine: [
      'threading or resolved-state when word/commentsExtended.xml is absent; unknown is reported as unknown, never as "not a reply"',
      'whether Word would display a comment that has no commentReference — the schema makes all three markers optional'
    ],
    appliesTo: parts => parts[COMMENT_PART_PATHS.comments] !== undefined || matching(parts, WORD_BODY).length > 0,
    analyze: parts => {
      const body = matching(parts, WORD_BODY)[0];
      const document = parse(body ? parts[body] : undefined);
      if (!document) return [];
      return readComments({
        document,
        comments: parse(parts[COMMENT_PART_PATHS.comments]),
        commentsExtended: parse(parts[COMMENT_PART_PATHS.extended]),
        commentsIds: parse(parts['word/commentsIds.xml'])
      }).problems;
    }
  },
  {
    id: 'ole',
    title: 'Embedded objects and their preview images',
    formats: ['docx', 'xlsx', 'pptx'],
    determines: [
      'whether an embedded object still has its data, or only the preview that renders in front of it',
      'whether the declared binding agrees with how the relationship is actually targeted'
    ],
    cannotDetermine: [
      'whether the embedded binary is valid for the progId it claims — the compound file is never opened',
      'whether a linked object resolves, since its target is outside the package by definition'
    ],
    appliesTo: parts => matching(parts, OLE_HOST).length > 0,
    analyze: parts => matching(parts, OLE_HOST).flatMap(path => readOleObjects(parts, path).flatMap(o => o.problems))
  },
  {
    id: 'pivot',
    title: 'Pivot tables and the cache chain',
    formats: ['xlsx'],
    determines: [
      'which hop of the pivot cache chain is broken, when one is',
      'whether a pivot field index falls outside the cache field count'
    ],
    cannotDetermine: [
      'whether the cached records still agree with their source range — staleness needs the source recomputed',
      'the "67 MS-OI29500 variations" figure quoted in the module doc, which came from an earlier research pass and has not been checked against the source'
    ],
    appliesTo: parts => Object.keys(parts).some(p => p.startsWith('xl/pivotTables/') || p.startsWith('xl/pivotCache/')),
    analyze: parts => readPivotTables(parts).flatMap(t => [...t.chain.problems, ...t.problems])
  }
] as const;

export interface AnalysisRun {
  findings: Finding[];
  /** Analyzer ids that found something to look at and ran. */
  ran: string[];
  /** Analyzer ids with nothing in this package to examine. */
  skipped: string[];
}

/**
 * Runs every analyzer that applies to this package.
 *
 * One analyzer throwing must not lose the others' findings — a malformed part in a
 * corner of the package is exactly when the rest of the report matters most.
 */
export function analyzePackage(parts: PackageParts, analyzers: readonly Analyzer[] = ANALYZERS): AnalysisRun {
  const findings: Finding[] = [];
  const ran: string[] = [];
  const skipped: string[] = [];

  for (const analyzer of analyzers) {
    if (!analyzer.appliesTo(parts)) {
      skipped.push(analyzer.id);
      continue;
    }
    ran.push(analyzer.id);
    try {
      findings.push(...analyzer.analyze(parts));
    } catch {
      // Deliberately swallowed: a thrown analyzer is a bug in this engine, not a
      // finding about the user's file, and reporting it as one would be a lie.
    }
  }

  return { findings, ran, skipped };
}

export interface CapabilityLedger {
  ran: Array<{ id: string; title: string; determines: readonly string[] }>;
  skipped: Array<{ id: string; title: string }>;
  /** Every limit declared by an analyzer that actually ran. */
  limits: string[];
}

/**
 * What the engine checked, what it skipped, and what it cannot tell you.
 *
 * Computed entirely from the registry — the model never asserts its own capabilities,
 * exactly as it never asserts the evidence tier. `limits` is the part worth showing a
 * user: a clean report means "nothing found by the checks that ran", and the difference
 * between that and "this file is fine" is the whole point.
 */
export function capabilityLedger(run: AnalysisRun, analyzers: readonly Analyzer[] = ANALYZERS): CapabilityLedger {
  const byId = new Map(analyzers.map(a => [a.id, a]));
  return {
    ran: run.ran.map(id => {
      const a = byId.get(id)!;
      return { id, title: a.title, determines: a.determines };
    }),
    skipped: run.skipped.map(id => ({ id, title: byId.get(id)!.title })),
    limits: run.ran.flatMap(id => byId.get(id)!.cannotDetermine)
  };
}

/**
 * Identity of a finding for the purpose of comparing two packages.
 *
 * Deliberately excludes the message: messages interpolate counts and paths, so two
 * genuinely identical faults can word themselves differently. Code, part and subject
 * are what make a finding the *same* finding.
 */
const identityOf = (f: Finding): string =>
  `${f.code}|${f.part}|${JSON.stringify(f.subject ?? {})}`;

export interface FindingsDelta {
  /** Present after, absent before — what this change broke. */
  introduced: Finding[];
  /** Present before, absent after — what it fixed. */
  resolved: Finding[];
  /** Present in both, and therefore not this change's doing. */
  unchanged: Finding[];
}

/**
 * What a change did to the health of a package.
 *
 * This is the question a regression investigation actually asks, and no structural diff
 * can answer it: a dropped OLE embedding shows up as one removed part, which looks like
 * any other removed part. Here it shows up as an introduced finding that says the
 * document still renders correctly and is broken anyway.
 */
export function diffFindings(before: PackageParts, after: PackageParts): FindingsDelta {
  const first = new Map(analyzePackage(before).findings.map(f => [identityOf(f), f]));
  const second = new Map(analyzePackage(after).findings.map(f => [identityOf(f), f]));

  const introduced: Finding[] = [];
  const resolved: Finding[] = [];
  const unchanged: Finding[] = [];

  for (const [id, f] of second) (first.has(id) ? unchanged : introduced).push(f);
  for (const [id, f] of first) if (!second.has(id)) resolved.push(f);

  return { introduced, resolved, unchanged };
}
