/**
 * Turning "this looks wrong" into a report an engineer can act on.
 *
 * The coverage gap log (`coverageGaps.ts`) records one thing automatically: a part opened
 * that **no analyzer matched at all**. That is a real signal, but it is narrow, and the
 * two most valuable reports fall outside it entirely:
 *
 *   - **A false positive.** The engine reported a fault in a file Office is perfectly
 *     happy with. The part *was* covered, so the gap log saw nothing.
 *   - **A miss.** Someone knows a document is broken and the engine said nothing. Again
 *     the part was covered, so nothing was logged.
 *
 * Both need a person to say so, and a person will only say so if it takes one click. This
 * builds the report for them.
 *
 * WHAT IS DELIBERATELY NOT IN THE REPORT.
 *
 * **No document content.** Not text, not cell values, not the selected markup, not the
 * file name. This tool defaults to on-device AI precisely so that confidential documents
 * never leave the machine, and a bug-reporting feature that quietly undoes that would be
 * worse than having none — the person clicking the button is trying to be helpful and has
 * no reason to suspect they are exfiltrating a contract.
 *
 * What it does carry is engine output: which analyzers ran, which finding codes fired,
 * and the *normalised* part path (`word/header#.xml`). All of that is ECMA-376 vocabulary
 * and this engine's own strings. The reporter adds their own description, and the
 * placeholder tells them not to paste content into it.
 *
 * The result is plain Markdown for the clipboard — no backend, no telemetry endpoint, no
 * account. It goes wherever the team already triages: an issue tracker, a chat channel.
 */

import { analyzePackage, capabilityLedger } from './analyzers';
import { analyzerOf, type Finding } from './findings';
import { normalisePartPath } from './coverageGaps';
import { REPORT_SCHEMA_VERSION } from './report';
import type { PackageParts } from './packageIntegrity';

/** What the reporter is telling us. */
export type IssueKind =
  /** The engine reported something that is not actually wrong. */
  | 'false-positive'
  /** The engine said nothing about something that is wrong. */
  | 'missed-fault'
  /** The engine is right but the wording is unclear or unhelpful. */
  | 'unclear-message'
  /** No analyzer covers this at all. */
  | 'not-covered';

const KIND_LABEL: Record<IssueKind, string> = {
  'false-positive': 'False positive — the engine flagged a file that is fine',
  'missed-fault': 'Missed fault — the engine said nothing about a real problem',
  'unclear-message': 'Unclear message — right finding, unhelpful wording',
  'not-covered': 'Not covered — no analyzer looks at this'
};

/**
 * What each kind of report most needs from the reporter.
 *
 * Generic prompts produce generic reports. A false positive is only actionable with the
 * markup pattern; a miss is only actionable if we know what should have been said.
 */
const KIND_PROMPT: Record<IssueKind, string> = {
  'false-positive':
    'Which finding was wrong, and why is the file actually fine? If Office opens it without a repair prompt, say so — that is the standard we hold the engine to.',
  'missed-fault':
    'What is wrong with the document, and how would someone notice? If it renders correctly and is broken anyway, say that explicitly — those are the ones worth catching.',
  'unclear-message':
    'What did you think it meant, and what would have been clearer? The message should say what a reader of the document would get wrong.',
  'not-covered':
    'What should this check look for, and what breaks silently when it is wrong?'
};

export interface IssueReportInput {
  kind: IssueKind;
  /** The part being looked at, if any. Normalised before it goes in the report. */
  partPath?: string;
  /** Element local name, e.g. `w:bookmarkStart`. Spec vocabulary, not content. */
  elementName?: string;
  /** The reporter's own words. */
  description: string;
  /** The package under analysis, used to summarise what the engine did. */
  parts?: PackageParts;
}

/** Groups findings by code so a repeated fault is one line, not fifty. */
const summariseFindings = (findings: readonly Finding[]): string[] => {
  const byCode = new Map<string, number>();
  for (const f of findings) byCode.set(f.code, (byCode.get(f.code) ?? 0) + 1);
  return [...byCode.entries()]
    .sort((a, b) => b[1] - a[1])
    .map(([code, n]) => `- \`${code}\` ×${n}`);
};

/**
 * Builds a Markdown issue report.
 *
 * Deterministic apart from the timestamp, so two reports about the same situation are
 * comparable rather than merely similar.
 */
export function buildIssueReport(input: IssueReportInput, now: Date = new Date()): string {
  const lines: string[] = [
    `## OOXML Engine issue: ${KIND_LABEL[input.kind]}`,
    '',
    `**Reported:** ${now.toISOString()}`,
    `**Report schema:** ${REPORT_SCHEMA_VERSION}`
  ];

  if (input.partPath) {
    // Normalised so word/header3.xml and word/header7.xml read as one situation, and so
    // a numbered part name cannot carry information about the document.
    lines.push(`**Part:** \`${normalisePartPath(input.partPath)}\``);
  }
  if (input.elementName) lines.push(`**Element:** \`${input.elementName}\``);

  lines.push('', '### What happened', '', input.description.trim() || '_(no description given)_');

  if (input.parts) {
    const run = analyzePackage(input.parts);
    const ledger = capabilityLedger(run);

    lines.push('', '### What the engine did', '');
    lines.push(`- Analyzers that ran: ${ledger.ran.map(a => `\`${a.id}\``).join(', ') || 'none'}`);
    lines.push(`- Skipped (nothing to look at): ${ledger.skipped.length}`);

    if (run.findings.length === 0) {
      lines.push('- Findings: **none**');
    } else {
      lines.push(`- Findings (${run.findings.length}):`, ...summariseFindings(run.findings));
    }

    // The analyzer most likely responsible, so triage starts in the right file.
    const owner = input.partPath
      ? [...new Set(run.findings.map(f => analyzerOf(f.code)))].join(', ')
      : '';
    if (owner) lines.push(`- Likely owner: \`${owner}\``);

    if (ledger.limits.length > 0) {
      lines.push('', '<details><summary>What these checks cannot establish</summary>', '');
      for (const limit of ledger.limits) lines.push(`- ${limit}`);
      lines.push('', '</details>');
    }
  }

  lines.push(
    '',
    '---',
    '',
    '_No document content is included in this report — only part paths, element names and',
    "engine output. If the file can be shared, attach it separately; if it cannot, a",
    'hand-made file showing the same markup pattern is just as useful._'
  );

  return lines.join('\n');
}

/** The prompt shown to someone filing this kind of report. */
export const promptFor = (kind: IssueKind): string => KIND_PROMPT[kind];

/** Every kind, for building a picker. */
export const ISSUE_KINDS: ReadonlyArray<{ kind: IssueKind; label: string }> = (
  Object.keys(KIND_LABEL) as IssueKind[]
).map(kind => ({ kind, label: KIND_LABEL[kind] }));
