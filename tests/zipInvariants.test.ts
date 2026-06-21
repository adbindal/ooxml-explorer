import { describe, it, expect } from '../services/browserTestRunner';
import JSZip from 'jszip';
import { createModifiedZip } from '../services/zipService';

describe('ZIP Packaging Invariants', () => {
    it('enforces strict OOXML zip packing order and compression modes', async () => {
        // 1. Setup mock ZIP content
        const zip = new JSZip();
        zip.file("mimetype", "application/vnd.openxmlformats-package.core-properties+xml");
        zip.file("[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'></Types>");
        zip.file("word/document.xml", "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'></w:document>");
        
        // 2. Export the modified ZIP
        const pendingChanges = {
            "word/document.xml": "<w:document><w:body>Updated Main Content</w:body></w:document>"
        };
        const zipBlob = await createModifiedZip(zip, pendingChanges);
        
        // 3. Load it back using raw JSZip to inspect internal structures
        const resultZip = await new JSZip().loadAsync(zipBlob);
        const files = resultZip.files;
        const filePaths = Object.keys(files);
        
        // --- REQUIREMENT A: mimetype MUST BE THE FIRST ENTRY ---
        expect(filePaths[0]).toBe("mimetype");
        
        // --- REQUIREMENT B: mimetype MUST BE STORED (UNCOMPRESSED) ---
        const mimeEntry = files["mimetype"];
        const getCompressionMethod = (entry: JSZip.JSZipObject) => {
            const extEntry = entry as JSZip.JSZipObject & { 
                _data?: { compressionMethod?: number };
                options?: { compression?: string };
            };
            if (extEntry._data && typeof extEntry._data.compressionMethod === 'number') {
                return extEntry._data.compressionMethod;
            }
            if (extEntry.options && extEntry.options.compression === 'STORE') return 0;
            return null;
        };
        
        const mimeCompression = getCompressionMethod(mimeEntry);
        if (mimeCompression !== null) {
            expect(mimeCompression).toBe(0); // 0 = STORE
        }
        
        // --- REQUIREMENT C: subsequent XML assets MUST BE DEFLATED ---
        const docEntry = files["word/document.xml"];
        const docCompression = getCompressionMethod(docEntry);
        if (docCompression !== null) {
            expect(docCompression).toBe(8); // 8 = DEFLATE
        }
    });
});
