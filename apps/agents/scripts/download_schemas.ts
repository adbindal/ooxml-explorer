import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const SCHEMAS_DIR = path.join(__dirname, '../schemas');

const SCHEMAS = [
  {
    name: 'wml.xsd',
    url: 'https://raw.githubusercontent.com/QtExcel/ecma-376-5th/master/ECMA-376/OfficeOpenXML-XMLSchema-Transitional/wml.xsd'
  },
  {
    name: 'sml.xsd',
    url: 'https://raw.githubusercontent.com/QtExcel/ecma-376-5th/master/ECMA-376/OfficeOpenXML-XMLSchema-Transitional/sml.xsd'
  },
  {
    name: 'pml.xsd',
    url: 'https://raw.githubusercontent.com/QtExcel/ecma-376-5th/master/ECMA-376/OfficeOpenXML-XMLSchema-Transitional/pml.xsd'
  },
  {
    name: 'dml-main.xsd',
    url: 'https://raw.githubusercontent.com/QtExcel/ecma-376-5th/master/ECMA-376/OfficeOpenXML-XMLSchema-Transitional/dml-main.xsd'
  },
  {
    name: 'shared-commonSimpleTypes.xsd',
    url: 'https://raw.githubusercontent.com/QtExcel/ecma-376-5th/master/ECMA-376/OfficeOpenXML-XMLSchema-Transitional/shared-commonSimpleTypes.xsd'
  },
  {
    name: 'opc-relationships.xsd',
    url: 'https://raw.githubusercontent.com/QtExcel/ecma-376-5th/master/ECMA-376/OpenPackagingConventions-XMLSchema/opc-relationships.xsd'
  }
];

async function downloadFile(name: string, url: string) {
  const destPath = path.join(SCHEMAS_DIR, name);
  
  console.log(`Downloading ${name} from ${url}...`);
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(`Failed to download ${name}: ${res.statusText}`);
  }
  const text = await res.text();
  fs.writeFileSync(destPath, text, 'utf8');
  console.log(`Saved ${name} to ${destPath}`);
}

async function main() {
  if (!fs.existsSync(SCHEMAS_DIR)) {
    fs.mkdirSync(SCHEMAS_DIR, { recursive: true });
  }
  
  for (const schema of SCHEMAS) {
    try {
      await downloadFile(schema.name, schema.url);
    } catch (e) {
      console.error(e);
    }
  }
}

main().catch(console.error);
