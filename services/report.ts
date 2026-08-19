/**
 * The engine's output as a document, not as a conversation.
 *
 * Everything upstream of here produces findings; this is where they become something
 * another program can hold onto. The stated direction for this project is that the
 * **engine is the product** and the side panel is its first consumer — *"based on those
 * we can build further agents for our use-cases."* An agent cannot parse prose back into
 * a fact, so the report is structured, sorted deterministically, and plain JSON.
 *
 * WHAT MAKES THIS A CONTRACT RATHER THAN A DUMP.
 *
 * `schemaVersion` is here because someone else's code will depend on these field names.
 * The rules, so they can be relied on:
 *
 *   - Adding a field, or adding a new `code` value, is a MINOR change.
 *   - Renaming or removing a field, or changing what an existing `code` means, is MAJOR.
 *   - Finding codes are part of the contract. `ole/data-part-missing` keeps that
 *     spelling or the version goes up.
 *
 * `generatedAt` is the only non-deterministic field. Everything else is a pure function
 * of the package, so two runs over the same file produce byte-identical reports apart
 * from that timestamp — which is what makes a report diffable in CI.
 *
 * WHAT IS DELIBERATELY NOT HERE.
 *
 * The *explain* path — the six-layer Word cascade, what a bookmark covers, where a
 * chart's numbers come from — stays prose and does not appear in this report. It
 * describes how markup **resolves**, which is a narrative answer to a question someone
 * asked, not a fault anyone can act on programmatically. Forcing it into findings would
 * mean inventing a fault for every fact. When an agent needs that too, it wants a
 * different shape, and inventing that shape before there is a caller would be guessing.
 */

import { analyzePackage, capabilityLedger, diffFindings, type CapabilityLedger } from './analyzers';
import { compareFindings, type Finding, type Severity } from './findings';
import type { PackageParts } from './packageIntegrity';

/** See the header for what a version bump means. */
export const REPORT_SCHEMA_VERSION = '1.0.0';

export interface FindingCounts {
  total: number;
  error: number;
  warning: number;
  note: number;
  /**
   * How many render correctly and are broken anyway. Called out at the top level
   * because it is the number that says whether looking at the document would have
   * found any of this.
   */
  silent: number;
}

export interface PackageReport {
  schemaVersion: string;
  /** ISO-8601. The only field that is not a pure function of the package. */
  generatedAt: string;
  parts: number;
  counts: FindingCounts;
  findings: Finding[];
  /** What ran, what was skipped, and what the checks that ran cannot establish. */
  capabilities: CapabilityLedger;
}

export interface ComparisonReport extends Omit<PackageReport, 'counts' | 'findings'> {
  /** Present after and not before — what this change broke. */
  introduced: Finding[];
  /** Present before and not after — what it fixed. */
  resolved: Finding[];
  /** Present in both, and therefore not this change's doing. */
  unchanged: Finding[];
  counts: { introduced: FindingCounts; resolved: FindingCounts };
}

const countBy = (findings: readonly Finding[]): FindingCounts => {
  const of = (severity: Severity) => findings.filter(f => f.severity === severity).length;
  return {
    total: findings.length,
    error: of('error'),
    warning: of('warning'),
    note: of('note'),
    silent: findings.filter(f => f.silent).length
  };
};

/** Sorted so two runs over the same package produce identical output. */
const ordered = (findings: readonly Finding[]): Finding[] => [...findings].sort(compareFindings);

/**
 * Everything the engine can say about one package.
 *
 * `now` is injectable so a caller — or a test — can pin the timestamp and compare two
 * reports byte for byte.
 */
export function reportPackage(parts: PackageParts, now: Date = new Date()): PackageReport {
  const run = analyzePackage(parts);
  const findings = ordered(run.findings);

  return {
    schemaVersion: REPORT_SCHEMA_VERSION,
    generatedAt: now.toISOString(),
    parts: Object.keys(parts).length,
    counts: countBy(findings),
    findings,
    capabilities: capabilityLedger(run)
  };
}

/**
 * What a change did to the health of a package.
 *
 * The report a regression investigation wants: not "these bytes differ" but "this
 * change introduced two faults, fixed one, and left four alone that were already
 * there". `parts` counts the *after* side, since that is the package under review.
 */
export function reportComparison(
  before: PackageParts,
  after: PackageParts,
  now: Date = new Date()
): ComparisonReport {
  const delta = diffFindings(before, after);
  const run = analyzePackage(after);

  return {
    schemaVersion: REPORT_SCHEMA_VERSION,
    generatedAt: now.toISOString(),
    parts: Object.keys(after).length,
    introduced: ordered(delta.introduced),
    resolved: ordered(delta.resolved),
    unchanged: ordered(delta.unchanged),
    counts: { introduced: countBy(delta.introduced), resolved: countBy(delta.resolved) },
    capabilities: capabilityLedger(run)
  };
}

/**
 * A one-line verdict for a human, derived from the counts.
 *
 * Kept deliberately careful: a report with no findings means the checks that ran found
 * nothing, which is not the same as the file being correct. The capability ledger says
 * what those checks were, and this sentence refuses to overstate them.
 */
export function summariseReport(report: PackageReport): string {
  const { counts } = report;
  if (counts.total === 0) {
    return `No problems found by the ${report.capabilities.ran.length} check(s) that ran. That is not the same as the file being correct — see what these checks cannot establish.`;
  }
  const parts = [`${counts.error} error(s)`, `${counts.warning} warning(s)`, `${counts.note} note(s)`];
  const silent =
    counts.silent > 0
      ? ` ${counts.silent} of these render correctly and are broken anyway, so no visual check would catch them.`
      : '';
  return `${parts.join(', ')} across ${report.parts} part(s).${silent}`;
}

/** The report as JSON, stable enough to commit to a repository and diff in review. */
export const serialiseReport = (report: PackageReport | ComparisonReport): string =>
  JSON.stringify(report, null, 2);
