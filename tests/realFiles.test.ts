import { describe, it, expect } from 'vitest';
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import { readPackageParts } from '../services/zipService';
import { analyzePackage, capabilityLedger } from '../services/analyzers';
import { compareFindings, type Finding } from '../services/findings';
import { reportPackage, summariseReport } from '../services/report';

/**
 * The engine, run against files Office actually produced.
 *
 * Every other test in this repo uses XML written by hand — by the same person who wrote
 * the code under test, which means both share the same assumptions about what real markup
 * looks like. **A false positive on a genuine document is invisible to all of them.**
 * This is the only test that can catch one.
 *
 * THE ASSERTION, AND WHY IT IS THE RIGHT WAY ROUND.
 *
 * A fixture is *known good* if Office opens it without a repair prompt and it renders as
 * intended. Against such a file the engine must report **nothing**. If it reports a
 * finding, the engine is wrong — not the document. That direction matters: it is easy to
 * write an analyzer that catches every real fault and also fires on half the corpus, and
 * such an analyzer is worse than none, because a report nobody trusts is a report nobody
 * reads.
 *
 * Name a file `*.expect-findings.docx` to opt out — for a document deliberately broken to
 * confirm the engine does catch it.
 *
 * SKIPPING IS DELIBERATE.
 *
 * The fixture binaries are gitignored (see `tests/fixtures/README.md`) so that pointing
 * that directory at a confidential work document cannot leak it. CI therefore has no
 * fixtures, and this suite **skips cleanly rather than failing**.
 *
 * ⚠️ The consequence, stated plainly because it is exactly the kind of thing that gets
 * forgotten: **a green test suite does not mean the engine has been checked against real
 * files.** It means it was checked against whatever fixtures were present, which in CI is
 * none. Run this locally with real documents and read the output.
 */

const FIXTURE_DIR = join(__dirname, 'fixtures');
const OFFICE_FILE = /\.(docx|xlsx|pptx|docm|xlsm|pptm)$/i;
/** A file deliberately broken to prove the engine catches it. */
const EXPECTS_FINDINGS = /\.expect-findings\./i;

const fixtures = existsSync(FIXTURE_DIR)
  ? readdirSync(FIXTURE_DIR).filter(name => OFFICE_FILE.test(name)).sort()
  : [];

const loadParts = async (name: string): Promise<Record<string, string>> => {
  const zip = await new JSZip().loadAsync(readFileSync(join(FIXTURE_DIR, name)));
  return readPackageParts(zip);
};

/** Groups findings by code so a repeated fault reads as one line, not fifty. */
const summarise = (findings: readonly Finding[]): string => {
  const byCode = new Map<string, number>();
  for (const f of findings) byCode.set(f.code, (byCode.get(f.code) ?? 0) + 1);
  return [...byCode.entries()]
    .sort((a, b) => b[1] - a[1])
    .map(([code, n]) => `${code} ×${n}`)
    .join(', ');
};

describe.skipIf(fixtures.length === 0)('real Office files', () => {
  it('found fixtures to check', () => {
    // Announces the corpus size, so a run against one file is not mistaken for a run
    // against a hundred.
    console.log(`\n  Checking ${fixtures.length} real file(s): ${fixtures.join(', ')}\n`);
    expect(fixtures.length).toBeGreaterThan(0);
  });

  for (const name of fixtures) {
    const expectsFindings = EXPECTS_FINDINGS.test(name);

    describe(name, () => {
      it('parses as an OPC package', async () => {
        const parts = await loadParts(name);

        expect(Object.keys(parts).length).toBeGreaterThan(0);
        expect(parts['[Content_Types].xml']).toBeDefined();
      });

      it('runs every applicable analyzer without throwing', async () => {
        const parts = await loadParts(name);
        const run = analyzePackage(parts);
        const ledger = capabilityLedger(run);

        console.log(`  ${name}: ${ledger.ran.length} analyzer(s) ran, ${ledger.skipped.length} skipped`);
        expect(run.ran.length).toBeGreaterThan(0);
      });

      if (expectsFindings) {
        it('reports the faults it was built to demonstrate', async () => {
          const parts = await loadParts(name);
          const { findings } = analyzePackage(parts);

          console.log(`  ${name}: ${summarise(findings)}`);
          expect(findings.length).toBeGreaterThan(0);
        });
      } else {
        it('reports NOTHING, because Office is happy with this file', async () => {
          // The assertion the whole fixture corpus exists for. A finding here is a false
          // positive: the engine contradicting Office about a file Office accepts.
          const parts = await loadParts(name);
          const { findings } = analyzePackage(parts);

          if (findings.length > 0) {
            const detail = [...findings]
              .sort(compareFindings)
              .map(f => `    [${f.code}] ${f.part} — ${f.message}`)
              .join('\n');
            console.error(`\n  FALSE POSITIVES in ${name}:\n${detail}\n`);
          }

          expect(findings, `false positive(s) in ${name}: ${summarise(findings)}`).toEqual([]);
        });
      }

      it('produces a serialisable report', async () => {
        // The export path, exercised against real input rather than a fixture built to
        // suit it. A report that cannot round-trip is useless to the agents downstream.
        const parts = await loadParts(name);
        const report = reportPackage(parts, new Date('2026-01-01T00:00:00.000Z'));

        console.log(`  ${name}: ${summariseReport(report)}`);
        expect(JSON.parse(JSON.stringify(report))).toEqual(report);
      });
    });
  }
});

describe.skipIf(fixtures.length > 0)('real Office files (none present)', () => {
  it('skips, and says so rather than passing silently', () => {
    // A passing suite with no fixtures must not read as "checked against real files".
    console.log(
      '\n  No fixtures in tests/fixtures/ — real-file checks were SKIPPED, not passed.' +
        '\n  Drop .docx/.xlsx/.pptx files there and re-run. See tests/fixtures/README.md.\n'
    );
    expect(fixtures).toEqual([]);
  });
});
