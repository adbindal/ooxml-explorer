import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const SCHEMAS_DIR = path.join(__dirname, '../schemas');

const SOURCES = [
  {
    domain: 'docx',
    url: 'https://raw.githubusercontent.com/dotnet/Open-XML-SDK/main/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_openxmlformats_org_wordprocessingml_2006_main.g.cs'
  },
  {
    domain: 'xlsx',
    url: 'https://raw.githubusercontent.com/dotnet/Open-XML-SDK/main/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_openxmlformats_org_spreadsheetml_2006_main.g.cs'
  },
  {
    domain: 'pptx',
    url: 'https://raw.githubusercontent.com/dotnet/Open-XML-SDK/main/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_openxmlformats_org_presentationml_2006_main.g.cs'
  }
];

interface ClassMap {
  [key: string]: string; // "domain:tag" -> "ClassName"
}

async function fetchAndParse(domain: string, url: string, map: ClassMap) {
  console.log(`Downloading ${domain} from ${url}...`);
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(`Failed to download ${domain}: ${res.statusText}`);
  }
  const text = await res.text();
  const lines = text.split('\n');

  let currentClass: string | null = null;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    // Match class declaration: public partial class ClassName :
    const classMatch = line.match(/public\s+partial\s+class\s+([a-zA-Z0-9_]+)\s*:/);
    if (classMatch) {
      currentClass = classMatch[1];
      continue;
    }

    // Match ElementQName definition: internal static readonly new OpenXmlQualifiedName ElementQName = new("...", "tag");
    const qnameMatch = line.match(/ElementQName\s*=\s*new\s*\([^,]+,\s*["']([^"']+)["']\)/);
    if (qnameMatch && currentClass) {
      const tag = qnameMatch[1];
      const key = `${domain}:${tag}`;
      map[key] = currentClass;
      currentClass = null; // reset
    }
  }
}

async function main() {
  if (!fs.existsSync(SCHEMAS_DIR)) {
    fs.mkdirSync(SCHEMAS_DIR, { recursive: true });
  }

  const map: ClassMap = {};

  for (const source of SOURCES) {
    try {
      await fetchAndParse(source.domain, source.url, map);
    } catch (e) {
      console.error(`Error processing ${source.domain}:`, e);
    }
  }

  // Write the map to schemas/sdkClassMap.json
  const destPath = path.join(SCHEMAS_DIR, 'sdkClassMap.json');
  fs.writeFileSync(destPath, JSON.stringify(map, null, 2), 'utf8');
  console.log(`Successfully compiled SDK class map! Saved ${Object.keys(map).length} mappings to ${destPath}`);
}

main().catch(console.error);
