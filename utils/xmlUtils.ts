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
 * Minifies XML by removing unnecessary whitespaces.
 */
export const minifyXml = (xml: string): string => {
  if (!xml || typeof xml !== 'string') return '';
  return xml.replace(/>\s+</g, '><').trim();
};
