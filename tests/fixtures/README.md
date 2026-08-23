# Real-file fixtures

Drop real `.docx`, `.xlsx` and `.pptx` files here. `tests/realFiles.test.ts` runs every
analyzer over each one and **fails if a known-good file produces a finding**.

That is the whole point: every other test in this repo uses hand-written XML that was
written by the same person who wrote the code being tested. A false positive on a real
document is invisible to all of them.

## The binaries are gitignored on purpose

`.gitignore` excludes the document files themselves so that dropping a confidential
document in here cannot leak it into version control. **This directory is safe to point
at your own work files.**

The consequence is that CI has no fixtures, so `realFiles.test.ts` **skips cleanly when
the directory is empty** rather than failing. A green suite therefore does not imply
real-file coverage — run `npm run test:real` locally and read what it prints.

## Where to get files

**Your own documents are the best corpus**, because they are the ones your engineers
actually process. Anything you can open in Office works; no conversion needed.

For a shared corpus, real Office output exists under permissive licences in projects that
test against it — Apache POI (Apache-2.0), LibreOffice (MPL-2.0), python-docx (MIT). Those
are genuine files produced by real Office versions, published so that tools can test
against them.

## What counts as "known good"

A file is known-good if **Office opens it without a repair prompt and it renders as
intended**. That is the standard the assertion encodes: if the engine reports a fault in a
file Office is happy with, the engine is wrong, not the file.

Name a file `*.expect-findings.*` to opt it out of that assertion — for a document you
have deliberately broken to confirm the engine *does* catch it.

## Licensing note

⚠️ This is unrelated to `../../docs/ooxml-expert-agent/LICENSING.md`. That research is
about reproducing **[MS-OI29500]'s prose**, Microsoft's written specification text. A
document file is not that text, and none of the four open questions there apply here.
