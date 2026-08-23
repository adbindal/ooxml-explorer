/**
 * Writes two small but genuine OPC packages into `tests/fixtures/`.
 *
 * These are **not a substitute for real Office output** — they are written here, so they
 * share the blind spots of the code that reads them, which is the whole limitation the
 * fixture corpus exists to escape. What they do give you is a **baseline that proves the
 * harness itself works in both directions**, available to anyone with a clone and no
 * access to Office:
 *
 *   smoke-valid.docx                    → the engine must report NOTHING
 *   smoke-broken.expect-findings.docx   → the engine must report SOMETHING
 *
 * A harness that only ever sees healthy files cannot tell "no faults" from "not looking",
 * which is why the broken one is here too.
 *
 * The broken file breaks three things a person cannot see by looking at the document:
 * an unclosed bookmark (every cross-reference to it dies silently), and a paragraph style
 * that is referenced but never defined (Word falls back to Normal and the heading renders
 * as body text).
 *
 *   node scripts/makeSmokeFixtures.mjs
 *
 * Both outputs are gitignored — see tests/fixtures/README.md for why.
 */

import JSZip from 'jszip';
import { writeFileSync, mkdirSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const FIXTURES = join(dirname(fileURLToPath(import.meta.url)), '..', 'tests', 'fixtures');

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

/** @param {{ closeBookmark: boolean }} options */
const documentXml = ({ closeBookmark }) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>A heading</w:t></w:r></w:p>
<w:p><w:bookmarkStart w:id="1" w:name="Intro"/><w:r><w:t>Body text.</w:t></w:r>${
  closeBookmark ? '<w:bookmarkEnd w:id="1"/>' : ''
}</w:p>
<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF Intro \\h </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>A heading</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;

/** @param {{ defineHeading: boolean }} options */
const stylesXml = ({ defineHeading }) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${
  defineHeading
    ? '\n<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style>'
    : ''
}
</w:styles>`;

/** @param {{ healthy: boolean }} options */
async function build({ healthy }) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CONTENT_TYPES);
  zip.file('_rels/.rels', ROOT_RELS);
  zip.file('word/_rels/document.xml.rels', DOC_RELS);
  zip.file('word/document.xml', documentXml({ closeBookmark: healthy }));
  zip.file('word/styles.xml', stylesXml({ defineHeading: healthy }));
  return zip.generateAsync({ type: 'nodebuffer' });
}

mkdirSync(FIXTURES, { recursive: true });

for (const [name, healthy] of [
  ['smoke-valid.docx', true],
  ['smoke-broken.expect-findings.docx', false]
]) {
  const buffer = await build({ healthy });
  writeFileSync(join(FIXTURES, name), buffer);
  console.log(`wrote tests/fixtures/${name} (${buffer.length} bytes)`);
}

console.log('\nRun `npm run test:real` to check the engine against them.');
console.log('Add your own .docx/.xlsx/.pptx alongside — they are gitignored.');
