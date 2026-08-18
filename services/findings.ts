/**
 * One shape for everything this engine concludes.
 *
 * Every analyzer used to define its own near-identical problem type — `BookmarkProblem`,
 * `OleProblem`, `CommentProblem`, `PivotProblem`, `IntegrityFinding` — and each one
 * discovered the fields it needed independently. They drifted: only two carried
 * `silent`, only two carried `severity`, only one carried `part`, and one had no
 * `remediation` at all. Five shapes is survivable; the next twenty analyzers are not.
 *
 * WHY STRUCTURED, AND NOT PROSE.
 *
 * The analyzers previously flattened to `string[]` at the panel boundary. English
 * sentences are the one representation that *cannot* be filtered, ranked, grouped,
 * suppressed, counted, or diffed — and the stated goal is to feed this to other agents,
 * which cannot parse a sentence back into a fact. So prose is now **rendered from
 * findings**, never authored as findings. `renderFinding` is the only place a sentence
 * gets built, and it is one function rather than sixteen.
 *
 * WHY NOT EMBEDDINGS.
 *
 * A finding is a claim with a traceable origin: this code, in this part, about this
 * subject. That traceability is what computes the Verified badge — the orchestrator can
 * say *why* it believes something because the finding names what was read. A vector has
 * no origin to name. Compressing findings into an embedding space would keep the size
 * saving and throw away the one property the whole design rests on.
 */

/**
 * How much this matters.
 *
 * `note` is not a weak error — it marks a state that is legitimate and merely deserves
 * explaining. A pivot cache set to refresh on load genuinely has no records part; that
 * is not damage, and calling it damage teaches a reader to ignore the list.
 */
export type Severity = 'error' | 'warning' | 'note';

/** Contextual identifiers — a bookmark name, a comment id, a broken chain hop. */
export type FindingSubject = Readonly<Record<string, string>>;

export interface Finding {
  /**
   * Stable, namespaced, machine-readable: `analyzer/kind`, e.g. `ole/data-part-missing`.
   *
   * The namespace is what lets a consumer group or suppress by analyzer without a
   * separate field, and what keeps two analyzers from colliding on a generic kind like
   * `duplicate-id`. Treat these as a published contract: renaming one breaks callers.
   */
  code: string;
  severity: Severity;
  /** Package part the finding is about, e.g. `word/document.xml`. */
  part: string;
  /** The consequence, in terms of what a person using the document would experience. */
  message: string;
  /** What to change. A finding without a fix is half a finding. */
  remediation: string;
  /**
   * True when the document still renders exactly as intended and is broken anyway.
   *
   * The most valuable bit in the record. These are the defects no screenshot test and
   * no human eye will ever catch — a dropped OLE embedding behind an intact preview, a
   * bookmark whose end vanished so every cross-reference to it resolves to nothing.
   */
  silent: boolean;
  subject?: FindingSubject;
}

const SEVERITY_RANK: Record<Severity, number> = { error: 0, warning: 1, note: 2 };

/**
 * Most-worth-reading first: severity, then silent before visible.
 *
 * Silent findings sort ahead of visible ones at equal severity because the visible ones
 * are already discoverable by looking at the document. The list should lead with what
 * looking cannot tell you.
 */
export const compareFindings = (a: Finding, b: Finding): number => {
  const bySeverity = SEVERITY_RANK[a.severity] - SEVERITY_RANK[b.severity];
  if (bySeverity !== 0) return bySeverity;
  if (a.silent !== b.silent) return a.silent ? -1 : 1;
  return a.code.localeCompare(b.code);
};

/** The analyzer that produced a finding — the part of the code before the slash. */
export const analyzerOf = (code: string): string => code.split('/')[0];

/** Every finding for one analyzer. */
export const findingsFrom = (findings: readonly Finding[], analyzer: string): Finding[] =>
  findings.filter(f => analyzerOf(f.code) === analyzer);

/**
 * The findings that render correctly and are broken anyway.
 *
 * This is the list to lead with when someone compares a before and after file: no
 * visual check will catch any of them.
 */
export const silentFindings = (findings: readonly Finding[]): Finding[] =>
  findings.filter(f => f.silent);

export const errorsOnly = (findings: readonly Finding[]): Finding[] =>
  findings.filter(f => f.severity === 'error');

/**
 * The single place a finding becomes a sentence.
 *
 * Deliberately plain: the message already states the consequence and the remediation
 * already states the fix, so this joins them rather than decorating them. Notes are
 * marked because an unmarked note reads as a defect.
 */
export const renderFinding = (finding: Finding): string => {
  const prefix = finding.severity === 'note' ? 'Note: ' : '';
  const silent = finding.silent ? ' (this renders correctly and is broken anyway)' : '';
  return `${prefix}${finding.message}${silent} ${finding.remediation}`.trim();
};

/**
 * Findings as evidence lines for the model, strongest first.
 *
 * The ordering matters more than it looks: when the prompt budget truncates, it
 * truncates the tail, so the sort decides what survives.
 */
export const renderFindings = (findings: readonly Finding[]): string[] =>
  [...findings].sort(compareFindings).map(renderFinding);

/**
 * Builds a finding, defaulting the fields most analyzers do not need to think about.
 *
 * `severity: 'error'` and `silent: false` are the safe defaults — an analyzer that
 * forgets to mark something silent under-claims, which is the direction to fail in.
 */
export const finding = (
  code: string,
  part: string,
  message: string,
  remediation: string,
  options: { severity?: Severity; silent?: boolean; subject?: FindingSubject } = {}
): Finding => ({
  code,
  severity: options.severity ?? 'error',
  part,
  message,
  remediation,
  silent: options.silent ?? false,
  ...(options.subject ? { subject: options.subject } : {})
});
