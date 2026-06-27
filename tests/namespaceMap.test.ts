import { describe, it, expect } from '../services/browserTestRunner';
import { getNamespaceUri } from '../services/namespaceMap';

describe('Namespace Map Utilities', () => {
    it('should resolve correct URI for WordprocessingML (w)', () => {
        const uri = getNamespaceUri('w', 'docx', 'document');
        expect(uri).toBe('http://schemas.openxmlformats.org/wordprocessingml/2006/main');
    });

    it('should resolve correct URI for SpreadsheetML (xlsx)', () => {
        const uri = getNamespaceUri('r', 'xlsx', 'row');
        expect(uri).toBe('http://schemas.openxmlformats.org/spreadsheetml/2006/main');
    });

    it('should resolve correct URI for PresentationML (p)', () => {
        const uri = getNamespaceUri('p', 'pptx', 'presentation');
        expect(uri).toBe('http://schemas.openxmlformats.org/presentationml/2006/main');
    });

    it('should resolve correct URI for OPC Relationships (shared)', () => {
        const uri = getNamespaceUri('r', 'shared', 'Relationship');
        expect(uri).toBe('http://schemas.openxmlformats.org/package/2006/relationships');
    });

    it('should resolve correct URI for OPC Content Types (shared)', () => {
        const uri = getNamespaceUri('r', 'shared', 'Override');
        expect(uri).toBe('http://schemas.openxmlformats.org/package/2006/content-types');
    });

    it('should return empty string for unknown prefixes/domains', () => {
        const uri = getNamespaceUri('unknown', 'docx', 'document');
        expect(uri).toBe('');
    });
});
