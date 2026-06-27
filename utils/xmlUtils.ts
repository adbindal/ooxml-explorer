/**
 * Formats XML string with indentation for readability.
 * Resilient against malformed, binary, or malicious inputs.
 */
export const formatXml = (xml: string): string => {
  if (!xml || typeof xml !== 'string') return '';
  
  try {
      let formatted = '';
      const reg = /(>)(<)(\/*)/g;
      const xmlStr = xml.replace(reg, '$1\n$2$3');
      let pad = 0;
      
      xmlStr.split('\n').forEach((node) => {
        let indent = 0;
        if (node.match(/.+<\/\w[^>]*>$/)) {
          indent = 0;
        } else if (node.match(/^<\/\w/)) {
          if (pad !== 0) pad -= 1;
        } else if (node.match(/^<\w[^>]*[^/]>.*$/)) {
          indent = 1;
        } else {
          indent = 0;
        }

        let padding = '';
        for (let i = 0; i < pad; i++) {
          padding += '  ';
        }

        formatted += padding + node + '\n';
        pad += indent;
      });

      return formatted.trim();
  } catch (e) {
      console.error("[Formatter] Failed to format XML, returning raw content", e);
      return xml;
  }
};

/**
 * Checks if a file path represents an XML document.
 */
export const isXmlFile = (path: string): boolean => {
  if (!path) return false;
  const lower = path.toLowerCase();
  return lower.endsWith('.xml') || lower.endsWith('.rels');
};

/**
 * Checks if a file path represents an image asset.
 */
export const isImageFile = (path: string): boolean => {
  if (!path) return false;
  const lower = path.toLowerCase();
  return lower.endsWith('.png') || lower.endsWith('.jpeg') || lower.endsWith('.jpg') || lower.endsWith('.gif') || lower.endsWith('.emf') || lower.endsWith('.wmf');
};

/**
 * Checks if a file path represents a binary asset (not text, not XML, not a web image).
 */
export const isBinaryFile = (path: string): boolean => {
  if (!path) return false;
  const lower = path.toLowerCase();
  // Image files are technically binary, but we support visual previews for them.
  // This function returns true for other binary files where editing/preview is not supported.
  return lower.endsWith('.bin') || 
         lower.endsWith('.otf') || 
         lower.endsWith('.ttf') || 
         lower.endsWith('.woff') || 
         lower.endsWith('.woff2') || 
         lower.endsWith('.zip');
};

/**
 * Minifies XML by removing unnecessary whitespaces.
 */
export const minifyXml = (xml: string): string => {
  if (!xml || typeof xml !== 'string') return '';
  return xml.replace(/>\s+</g, '><').trim();
};

/**
 * Extracts tag name and raw XML snippet from a selected text range or tag name.
 */
export const extractTagAtSelection = (
  content: string,
  selectionText: string
): { tagName: string; rawXml: string } | null => {
  const cleanSel = selectionText.trim();
  if (!cleanSel) return null;

  // Case 1: User selected a full XML tag e.g. <w:tcW w:w="120" /> or <w:p>
  if (cleanSel.startsWith('<') && cleanSel.endsWith('>')) {
    const match = cleanSel.match(/^<([\w:-]+)/);
    if (match) {
      return { tagName: match[1], rawXml: cleanSel };
    }
  }

  // Case 2: User selected just a tag name e.g. "w:tcW" or "tcW"
  const tagRegex = /^[a-zA-Z0-9_:-]+$/;
  if (tagRegex.test(cleanSel)) {
    // Let's find the enclosing XML tag in the content
    // Look for <cleanSel ... > or <cleanSel> or </cleanSel>
    const escapedTag = cleanSel.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&');
    const enclosingRegex = new RegExp(`<(${escapedTag})([^>]*?)(/?>|>)`);
    const match = content.match(enclosingRegex);
    if (match) {
      return { tagName: match[1], rawXml: match[0] };
    }
    return { tagName: cleanSel, rawXml: `<${cleanSel} />` };
  }

  return null;
};
