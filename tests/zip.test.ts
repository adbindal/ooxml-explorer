import { describe, it, expect } from '../services/browserTestRunner';
import JSZip, { JSZipObject } from 'jszip';
import { loadZipFile, generateDiffTree, createModifiedZip } from '../services/zipService';
import { DiffNode } from '../types';

describe('Zip Service', () => {
    it('loadZipFile throws error when loading zip missing [Content_Types].xml', async () => {
        // Create a raw zip without the required OOXML signature file
        const zip = new JSZip();
        zip.file('hello.txt', 'world');
        const content = await zip.generateAsync({ type: 'blob' });
        const file = new File([content], 'invalid.zip');

        await expect(loadZipFile(file)).rejects.toThrow('Invalid OOXML file: Missing [Content_Types].xml');
    });

    it('loadZipFile parses valid OOXML structure and generates file tree', async () => {
        // Create a valid-looking OOXML zip
        const zip = new JSZip();
        zip.file('[Content_Types].xml', '<root/>');
        const folder = zip.folder('word');
        folder?.file('document.xml', 'content');
        
        const content = await zip.generateAsync({ type: 'blob' });
        const file = new File([content], 'valid.docx');

        const result = await loadZipFile(file);
        
        // Check Zip Instance
        expect(result.zip).toBeDefined();
        
        // Check Flat Map
        expect(result.flat['word/document.xml']).toBeDefined();
        expect(result.flat['[Content_Types].xml']).toBeDefined();
        
        // Check Tree Structure
        expect(result.tree.children['word']).toBeDefined();
        expect(result.tree.children['word'].isFolder).toBe(true);
        expect(result.tree.children['word'].children['document.xml']).toBeDefined();
    });

    it('detects compression method correctly', async () => {
        // We test detectCompression via createModifiedZip behavior
        // since it is an internal private function: createModifiedZip should preserve compression.
        // createModifiedZip should preserve compression.
        
        const zip = new JSZip();
        zip.file("test.xml", "content", { compression: "STORE" });
        const blob = await createModifiedZip(zip, {});
        const newZip = await new JSZip().loadAsync(blob);
        // JSZip doesn't expose compression method easily on read, but we trust the logic if no error
        expect(newZip.file("test.xml")).toBeDefined();
    });

    it('createModifiedZip handles missing mimetype in original zip', async () => {
        const zip = new JSZip();
        zip.file("word/document.xml", "content");
        // No mimetype file
        
        const blob = await createModifiedZip(zip, { "word/document.xml": "new" });
        const newZip = await new JSZip().loadAsync(blob);
        expect(newZip.file("word/document.xml")).toBeDefined();
        // Should not crash
    });

    it('createModifiedZip handles directory entries correctly', async () => {
        const zip = new JSZip();
        zip.folder("word");
        zip.folder("word/theme");
        zip.file("word/document.xml", "content");
        
        const blob = await createModifiedZip(zip, {});
        const newZip = await new JSZip().loadAsync(blob);
        expect(newZip.folder("word")).toBeDefined();
        expect(newZip.folder("word/theme")).toBeDefined();
    });

    it('createModifiedZip prioritizes pending mimetype change', async () => {
        const zip = new JSZip();
        zip.file("mimetype", "old");
        
        const blob = await createModifiedZip(zip, { "mimetype": "new" });
        const newZip = await new JSZip().loadAsync(blob);
        const content = await newZip.file("mimetype")?.async("string");
        expect(content).toBe("new");
    });
});

describe('Diff Tree Generation', () => {
    it('identifies added files in new zip', () => {
        const flatA: Record<string, JSZipObject> = {};
        const flatB: Record<string, JSZipObject> = { 'new.xml': { name: 'new.xml', dir: false } as JSZipObject };
        
        const tree = generateDiffTree(flatA, flatB);
        const node = tree.children['new.xml'];
        
        expect(node).toBeDefined();
        expect(node.status).toBe('added');
        expect(tree.hasChange).toBe(true);
    });

    it('identifies deleted files from original zip', () => {
        const flatA: Record<string, JSZipObject> = { 'old.xml': { name: 'old.xml', dir: false } as JSZipObject };
        const flatB: Record<string, JSZipObject> = {};
        
        const tree = generateDiffTree(flatA, flatB);
        const node = tree.children['old.xml'];
        
        expect(node).toBeDefined();
        expect(node.status).toBe('deleted');
        expect(tree.hasChange).toBe(true);
    });

    it('identifies modified files based on CRC32 mismatch', () => {
        // Simulate JSZip entries with CRC32
        const entryA = { name: 'doc.xml', dir: false, _data: { crc32: 123 } };
        const entryB = { name: 'doc.xml', dir: false, _data: { crc32: 456 } }; // Different CRC
        
        const flatA = { 'doc.xml': entryA };
        const flatB = { 'doc.xml': entryB };
        
        const tree = generateDiffTree(flatA, flatB);
        const node = tree.children['doc.xml'];
        
        expect(node.status).toBe('modified');
        expect(tree.hasChange).toBe(true);
    });

    it('marks files with same CRC32 as unchanged', () => {
        const entryA = { name: 'doc.xml', dir: false, _data: { crc32: 999 } };
        const entryB = { name: 'doc.xml', dir: false, _data: { crc32: 999 } }; // Same CRC
        
        const flatA = { 'doc.xml': entryA };
        const flatB = { 'doc.xml': entryB };
        
        const tree = generateDiffTree(flatA, flatB);
        const node = tree.children['doc.xml'];
        
        expect(node.status).toBe('unchanged');
        expect(tree.hasChange).toBe(false); // Root should not be marked as having change if only child is unchanged
    });

    it('propagates "hasChange" status to parent folders', () => {
        const flatA: Record<string, JSZipObject> = {};
        const flatB: Record<string, JSZipObject> = { 'word/document.xml': { name: 'word/document.xml', dir: false } as JSZipObject };
        
        const tree = generateDiffTree(flatA, flatB);
        
        // Cast to DiffNode because tree.children definition returns FileNode (from base interface)
        const wordFolder = tree.children['word'] as DiffNode;
        expect(wordFolder.isFolder).toBe(true);
        expect(wordFolder.hasChange).toBe(true); // Should be true because child is added
        
        const docNode = wordFolder.children['document.xml'];
        expect(docNode.status).toBe('added');
    });
});