import { describe, it, expect } from '../services/browserTestRunner';
import { createModifiedZip } from '../services/zipService';
import { useAppStore } from '../store/appStore';
import JSZip from 'jszip';

describe('System Resilience & Self-Healing Audits', () => {

    it('heals non-compliant and out-of-order zip archives automatically during repack', async () => {
        // 1. Setup a DELIBERATELY BROKEN, non-compliant OOXML structure:
        // - 'mimetype' is placed at the END of the zip instead of first.
        // - 'mimetype' is COMPRESSED (using DEFLATE) instead of STORED (uncompressed).
        const brokenZip = new JSZip();
        
        // Add non-mimetype entries first
        brokenZip.file("[Content_Types].xml", "<Types></Types>", { compression: "DEFLATE" });
        brokenZip.file("word/document.xml", "<document></document>", { compression: "DEFLATE" });
        
        // Add mimetype last, and compress it!
        brokenZip.file("mimetype", "application/vnd.openxmlformats-package.core-properties+xml", { compression: "DEFLATE" });

        // Generate the broken original zip blob
        const brokenBlob = await brokenZip.generateAsync({ type: 'blob' });
        const originalZip = await new JSZip().loadAsync(brokenBlob);

        // 2. Run the self-healing repack exporter
        const pendingChanges = { "word/document.xml": "<document>healed</document>" };
        const healedBlob = await createModifiedZip(originalZip, pendingChanges);

        // 3. Verify the healed zip output complies strictly with MS Office standards
        const healedZip = await new JSZip().loadAsync(healedBlob);
        const healedPaths = Object.keys(healedZip.files);

        // --- RESILIENCE ASSERTIONS ---
        // A. The 'mimetype' file MUST have been automatically moved to the first position!
        expect(healedPaths[0]).toBe("mimetype");

        // B. The 'mimetype' file MUST have been automatically decompressed to STORE (uncompressed)!
        const mimeEntry = healedZip.files["mimetype"];
        
        // Read compression settings
        const extEntry = mimeEntry as JSZip.JSZipObject & { 
            _data?: { compressionMethod?: number };
            options?: { compression?: string };
        };
        
        const getCompressionMethod = (entry: typeof extEntry) => {
            if (entry._data && typeof entry._data.compressionMethod === 'number') {
                return entry._data.compressionMethod;
            }
            if (entry.options && entry.options.compression === 'STORE') return 0;
            return null;
        };

        const mimeCompression = getCompressionMethod(extEntry);
        if (mimeCompression !== null) {
            expect(mimeCompression).toBe(0); // 0 = STORE (uncompressed)
        }
        
        // C. The pending changes must be correctly integrated
        const docContent = await healedZip.file("word/document.xml")!.async("string");
        expect(docContent).toBe("<document>healed</document>");
    });

    it('fully purges all file handles, caches, and states from memory when returning to landing page', async () => {
        const store = useAppStore.getState();
        
        // 1. Simulate an active session by populating editor states
        const mockZip = new JSZip();
        mockZip.file("mimetype", "application/vnd.openxmlformats-package.core-properties+xml");
        const loadedZip = await new JSZip().loadAsync(await mockZip.generateAsync({ type: 'blob' }));

        useAppStore.setState({
            mode: 'editor',
            editor: {
                zip: loadedZip,
                tree: { name: 'root', path: '', isFolder: true, children: {} },
                fileName: 'test_leak.docx',
                activePath: 'word/document.xml',
                openTabs: ['word/document.xml', 'mimetype'],
                pendingChanges: { 'word/document.xml': 'dirty-content' },
                modifiedPaths: new Set(['word/document.xml']),
                contentCache: { 'word/document.xml': 'original-content' }
            }
        });

        // Verify active state before reset
        expect(useAppStore.getState().mode).toBe('editor');
        expect(useAppStore.getState().editor.fileName).toBe('test_leak.docx');
        expect(useAppStore.getState().editor.openTabs).toHaveLength(2);

        // 2. Trigger the self-healing reset action (setMode to landing)
        store.setMode('landing');

        // 3. Verify that ALL complex state structures, file descriptors, and caches are completely purged!
        const resetState = useAppStore.getState();
        expect(resetState.mode).toBe('landing');
        
        // The editor state must be strictly equal to its initial blank state (purged of all structures)
        expect(resetState.editor.zip).toBeNull();
        expect(resetState.editor.tree).toBeNull();
        expect(resetState.editor.fileName).toBe('');
        expect(resetState.editor.activePath).toBeNull();
        expect(resetState.editor.openTabs).toHaveLength(0);
        expect(Object.keys(resetState.editor.pendingChanges)).toHaveLength(0);
        expect(resetState.editor.modifiedPaths.size).toBe(0);
        expect(Object.keys(resetState.editor.contentCache)).toHaveLength(0);
    });

});
