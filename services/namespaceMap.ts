/**
 * Resolves the official XML Namespace URI for a given prefix, domain, and tag.
 * Decouples the short prefixes used in the database/UI from the verbose URIs in the spec.
 */
export const getNamespaceUri = (prefix: string, domain: string, tag: string): string => {
  if (prefix === 'w') {
    return 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  }
  if (prefix === 'p') {
    return 'http://schemas.openxmlformats.org/presentationml/2006/main';
  }
  if (domain === 'xlsx') {
    return 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  }
  if (domain === 'shared') {
    if (tag === 'Relationships' || tag === 'Relationship') {
      return 'http://schemas.openxmlformats.org/package/2006/relationships';
    }
    if (tag === 'Types' || tag === 'Override' || tag === 'Default') {
      return 'http://schemas.openxmlformats.org/package/2006/content-types';
    }
  }
  return '';
};
