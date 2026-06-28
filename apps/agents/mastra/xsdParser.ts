import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const SCHEMAS_DIR = path.join(__dirname, '../schemas');

export interface XSDGrounding {
  tag: string;
  domain: string;
  namespace: string;
  attributes: string[];
  parents: string[];
}

export function getXSDGrounding(tag: string, domain: string): XSDGrounding {
  let schemaFile = '';
  let defaultNamespace = '';

  if (domain === 'docx') {
    schemaFile = 'wml.xsd';
    defaultNamespace = 'w';
  } else if (domain === 'xlsx') {
    schemaFile = 'sml.xsd';
    defaultNamespace = 'r';
  } else if (domain === 'pptx') {
    schemaFile = 'pml.xsd';
    defaultNamespace = 'p';
  } else if (domain === 'shared') {
    schemaFile = 'opc-relationships.xsd';
    defaultNamespace = 'r';
  } else {
    throw new Error(`Unsupported domain: ${domain}`);
  }

  const filePath = path.join(SCHEMAS_DIR, schemaFile);
  if (!fs.existsSync(filePath)) {
    throw new Error(`Schema file not found: ${filePath}`);
  }

  const content = fs.readFileSync(filePath, 'utf8');
  const lines = content.split('\n');

  // 1. Find the line declaring the element: name="tag"
  const elementDeclRegex = new RegExp(`name=["']${tag}["']`);
  let declLineIdx = -1;
  for (let i = 0; i < lines.length; i++) {
    if (elementDeclRegex.test(lines[i]) && (lines[i].includes('<xsd:element') || lines[i].includes('<element'))) {
      declLineIdx = i;
      break;
    }
  }

  if (declLineIdx === -1) {
    return {
      tag,
      domain,
      namespace: defaultNamespace,
      attributes: [],
      parents: []
    };
  }

  const declLine = lines[declLineIdx];

  // Extract type of the element, e.g. type="CT_OnOff"
  const typeMatch = declLine.match(/type=["']([^"']+)["']/);
  const typeName = typeMatch ? typeMatch[1] : null;

  // 2. Find attributes
  const attributes: string[] = [];
  if (typeName) {
    const typeDefRegex = new RegExp(`(complexType|simpleType)\\s+name=["']${typeName}["']`);
    let typeDefLineIdx = -1;
    for (let i = 0; i < lines.length; i++) {
      if (typeDefRegex.test(lines[i])) {
        typeDefLineIdx = i;
        break;
      }
    }

    if (typeDefLineIdx !== -1) {
      let j = typeDefLineIdx + 1;
      let openBraces = 1;
      while (j < lines.length && openBraces > 0) {
        const line = lines[j];
        if (line.includes('<xsd:complexType') || line.includes('<xsd:simpleType') || line.includes('<complexType') || line.includes('<simpleType')) {
          openBraces++;
        }
        if (line.includes('</xsd:complexType') || line.includes('</xsd:simpleType') || line.includes('</complexType') || line.includes('</simpleType')) {
          openBraces--;
        }
        
        const attrMatch = line.match(/(?:xsd:)?attribute\s+name=["']([^"']+)["']/);
        if (attrMatch) {
          attributes.push(attrMatch[1]);
        }
        j++;
      }
    }
  }

  // 3. Find parent elements by walking backwards to find the containing complexType
  const parents: string[] = [];
  let parentType: string | null = null;
  for (let i = declLineIdx - 1; i >= 0; i--) {
    const line = lines[i];
    const complexTypeMatch = line.match(/(?:xsd:)?complexType\s+name=["']([^"']+)["']/);
    if (complexTypeMatch) {
      parentType = complexTypeMatch[1];
      break;
    }
  }

  if (parentType) {
    const childTypes = [parentType];
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const extMatch = line.match(/(?:xsd:)?extension\s+base=["']([^"']+)["']/);
      if (extMatch && extMatch[1] === parentType) {
        for (let j = i - 1; j >= 0; j--) {
          const compMatch = lines[j].match(/(?:xsd:)?complexType\s+name=["']([^"']+)["']/);
          if (compMatch) {
            childTypes.push(compMatch[1]);
            break;
          }
        }
      }
    }

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (line.includes('<xsd:element') || line.includes('<element')) {
        const typeAttrMatch = line.match(/type=["']([^"']+)["']/);
        if (typeAttrMatch && childTypes.includes(typeAttrMatch[1])) {
          const nameAttrMatch = line.match(/name=["']([^"']+)["']/);
          if (nameAttrMatch) {
            parents.push(nameAttrMatch[1]);
          }
        }
      }
    }
  }

  return {
    tag,
    domain,
    namespace: defaultNamespace,
    attributes: Array.from(new Set(attributes)),
    parents: Array.from(new Set(parents))
  };
}
