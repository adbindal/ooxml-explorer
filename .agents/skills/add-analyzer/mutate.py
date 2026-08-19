#!/usr/bin/env python3
"""Mutation-test one module against its test file.

Applies each mutation to the source one at a time, runs that file's tests, restores the
source, and reports how many tests each mutation killed.

**A mutation that kills zero tests is the finding.** On this project the cause has more
often been a weak test than missing coverage — a fixture that would pass either way, an
assertion decided by an earlier branch, an accidental tie-break, or a guard that is
provably unreachable. Suspect the test first. If the mutant is genuinely equivalent, say
so rather than inventing a test for it.

Usage
-----
    python3 .agents/skills/add-analyzer/mutate.py \
        --source services/wordNotes.ts \
        --test tests/wordNotes.test.ts \
        --mutations mutations.json

`mutations.json` is a list of objects:

    [
      {
        "label": "separators reported as orphans",
        "old": "if (n.structural || referenced.has(n.id)) continue;",
        "new": "if (referenced.has(n.id)) continue;"
      }
    ]

`old` must appear **exactly once**; the script refuses ambiguous or absent matches rather
than silently mutating the wrong line. Set `"new": ""` to delete.

Exit code is 0 when every mutation killed at least one test, 1 otherwise — so this can
gate a commit.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

TESTS_FAILED = re.compile(r"Tests\s+(\d+)\s+failed")


def run_tests(test_path: str) -> int:
    """Returns how many tests failed. Treats a crash as 'killed', since a mutation that
    breaks the build is certainly detected."""
    result = subprocess.run(
        ["npx", "vitest", "run", test_path],
        capture_output=True,
        text=True,
    )
    output = result.stdout + result.stderr
    match = TESTS_FAILED.search(output)
    if match:
        return int(match.group(1))
    # No "N failed" line: either everything passed, or the run never got that far.
    return 0 if result.returncode == 0 else -1


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Mutation-test a module.",
        epilog="A surviving mutant means the tests do not check that behaviour.",
    )
    parser.add_argument("--source", required=True, help="module to mutate")
    parser.add_argument("--test", required=True, help="test file to run")
    parser.add_argument("--mutations", required=True, help="JSON list of {label, old, new}")
    args = parser.parse_args()

    source = Path(args.source)
    if not source.is_file():
        print(f"error: {source} not found", file=sys.stderr)
        return 2

    mutations = json.loads(Path(args.mutations).read_text())
    original = source.read_text()

    # Keep the pristine copy outside the repo so a crash cannot leave a mutated file
    # staged. Restored in `finally` no matter how this exits.
    with tempfile.NamedTemporaryFile("w", suffix=source.suffix, delete=False) as backup:
        backup.write(original)
        backup_path = Path(backup.name)

    survivors: list[str] = []
    try:
        for mutation in mutations:
            label = mutation["label"]
            old, new = mutation["old"], mutation["new"]

            occurrences = original.count(old)
            if occurrences == 0:
                print(f"  SKIP  [{label}] — pattern not found; the code may have moved")
                survivors.append(f"{label} (pattern not found)")
                continue
            if occurrences > 1:
                print(f"  SKIP  [{label}] — pattern appears {occurrences}× ; make it unique")
                survivors.append(f"{label} (ambiguous pattern)")
                continue

            source.write_text(original.replace(old, new, 1))
            killed = run_tests(args.test)
            source.write_text(original)

            if killed == -1:
                print(f"  killed [{label}] — build/run failure (counts as detected)")
            elif killed == 0:
                print(f"  SURVIVED [{label}] — no test checks this")
                survivors.append(label)
            else:
                print(f"  killed [{label}] — {killed} test(s)")
    finally:
        shutil.copy(backup_path, source)
        backup_path.unlink(missing_ok=True)

    print()
    if survivors:
        print(f"{len(survivors)} of {len(mutations)} mutant(s) survived:")
        for s in survivors:
            print(f"  - {s}")
        print()
        print("Suspect the TEST before the code. Ask whether the fixture would have passed")
        print("either way, whether an earlier branch decided the assertion, or whether the")
        print("guard is reachable at all. If the mutant is truly equivalent, record that.")
        return 1

    print(f"All {len(mutations)} mutant(s) killed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
