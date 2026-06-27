import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

interface ReferenceDoc {
  tag: string;
  namespace: string;
  domain: string;
  definition: string;
  attributes: string[];
  parents: string[];
  citation: string;
  sdkClass: string;
}

const loadJson = (filePath: string): ReferenceDoc[] => {
  if (!fs.existsSync(filePath)) {
    console.error(`Error: File not found at ${filePath}`);
    process.exit(1);
  }
  return JSON.parse(fs.readFileSync(filePath, 'utf8'));
};

const main = () => {
  const goldenPath = path.join(__dirname, '../public/rag-data.json');
  const generatedPath = path.join(__dirname, '../public/generated-rag.json');

  const golden = loadJson(goldenPath);
  const generated = loadJson(generatedPath);

  console.log(`\n=== RAG Calibration Diff Report ===`);
  console.log(`Golden Reference: ${golden.length} tags`);
  console.log(`Generated Output: ${generated.length} tags\n`);

  let mismatchCount = 0;

  for (const goldDoc of golden) {
    const genDoc = generated.find(g => g.tag === goldDoc.tag && g.domain === goldDoc.domain);

    if (!genDoc) {
      console.log(`❌ Missing Tag: <${goldDoc.namespace}:${goldDoc.tag}> in domain ${goldDoc.domain}`);
      mismatchCount++;
      continue;
    }

    const diffs: string[] = [];

    // Compare namespace
    if (goldDoc.namespace !== genDoc.namespace) {
      diffs.push(`  - Namespace: Golden="${goldDoc.namespace}", Generated="${genDoc.namespace}"`);
    }

    // Compare citation
    if (goldDoc.citation !== genDoc.citation) {
      diffs.push(`  - Citation: Golden="${goldDoc.citation}", Generated="${genDoc.citation}"`);
    }

    // Compare sdkClass
    if (goldDoc.sdkClass !== genDoc.sdkClass) {
      diffs.push(`  - SDK Class: Golden="${goldDoc.sdkClass}", Generated="${genDoc.sdkClass}"`);
    }

    // Compare attributes (loose comparison of sets)
    const goldAttrs = new Set(goldDoc.attributes);
    const genAttrs = new Set(genDoc.attributes);
    const missingAttrs = [...goldAttrs].filter(x => !genAttrs.has(x));
    const extraAttrs = [...genAttrs].filter(x => !goldAttrs.has(x));
    if (missingAttrs.length > 0 || extraAttrs.length > 0) {
      diffs.push(`  - Attributes Mismatch:`);
      if (missingAttrs.length > 0) diffs.push(`    * Missing: [${missingAttrs.join(', ')}]`);
      if (extraAttrs.length > 0) diffs.push(`    * Extra: [${extraAttrs.join(', ')}]`);
    }

    // Compare parents
    const goldParents = new Set(goldDoc.parents);
    const genParents = new Set(genDoc.parents);
    const missingParents = [...goldParents].filter(x => !genParents.has(x));
    const extraParents = [...genParents].filter(x => !goldParents.has(x));
    if (missingParents.length > 0 || extraParents.length > 0) {
      diffs.push(`  - Parents Mismatch:`);
      if (missingParents.length > 0) diffs.push(`    * Missing: [${missingParents.join(', ')}]`);
      if (extraParents.length > 0) diffs.push(`    * Extra: [${extraParents.join(', ')}]`);
    }

    if (diffs.length > 0) {
      console.log(`⚠️ Mismatch in <${goldDoc.namespace}:${goldDoc.tag}> (${goldDoc.domain}):`);
      diffs.forEach(d => console.log(d));
      mismatchCount++;
    } else {
      console.log(`✅ Perfect Match: <${goldDoc.namespace}:${goldDoc.tag}>`);
    }
  }

  console.log(`\n===================================`);
  if (mismatchCount === 0) {
    console.log(`🎉 Success! 100% Alignment reached between Golden and Generated datasets!`);
  } else {
    console.log(`⚠️ Calibration Required: Found ${mismatchCount} mismatches/omissions.`);
  }
};

main();
