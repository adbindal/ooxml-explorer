/**
 * Formats XML string with indentation for readability.
 */
export const formatXml = (xml: string): string => {
  if (!xml) return '';
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
    } else if (node.match(/^<\w[^>]*[^\/]>.*$/)) {
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
};

/**
 * Minifies XML string for saving back to OOXML format (removes extra whitespace).
 */
export const minifyXml = (xml: string): string => {
  if (!xml) return '';
  return xml.replace(/>\s+</g, '><').trim();
};

export const isXmlFile = (filename: string): boolean => {
  return /\.(xml|rels)$/i.test(filename);
};

export const isImageFile = (filename: string): boolean => {
  return /\.(png|jpeg|jpg|gif|bmp|svg|webp)$/i.test(filename);
};
