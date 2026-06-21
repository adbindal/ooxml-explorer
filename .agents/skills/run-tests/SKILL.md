---
name: run-tests
description: Execute the full Vitest suite for the ooxml-explorer project to verify changes and code correctness.
---

# run-tests Skill

This skill allows any developer agent to programmatically execute the unit and integration test suite of the `ooxml-explorer` project.

## When to Use
Use this skill whenever you have:
1. Modified the Zustand store in `store/appStore.ts`.
2. Modified the ZIP parsing, diff, or packing engine in `services/zipService.ts`.
3. Modified rendering elements, file trees, or helper utilities.
4. Created new features and want to verify that existing features aren't broken (regression testing).

## How to Execute
To run the tests, execute the helper script `run_tests.sh` located in this skill folder:

```bash
./.agents/skills/run-tests/run_tests.sh
```

Alternatively, you can run the npm scripts directly:
- To run unit tests: `npm run test`
- To run coverage: `npm run test:coverage`

## Interpreting Results
- If all tests pass, the command will exit with code `0` and print a success message.
- If any test fails, it will exit with a non-zero exit code and output the precise assertion failures. Address these failures before declaring your task complete!
