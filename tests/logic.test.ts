import { describe, it, expect } from '../services/browserTestRunner';
import { formatXml, minifyXml, isXmlFile, isImageFile, isBinaryFile } from '../utils/xmlUtils';
import { calculateMatches, cycleIndex, replaceContent } from '../utils/searchUtils';
import { selectFileWithPicker } from '../utils/filePicker';
import { isSaveHotkey, isFindHotkey, isSaveAllHotkey, isSidebarHotkey } from '../utils/hotkeyUtils';
import { parseInlineStyles, parseMarkdownSegments } from '../utils/markdownUtils';
import { getModifiedPaths } from '../utils/treeUtils';
import { shouldAutoRunDiff } from '../utils/diffUtils';
import { DiffNode } from '../types';

describe('Diff Auto-Run Logic', () => {
    it('returns FALSE when original or modified files are missing', () => {
        expect(shouldAutoRunDiff('diff-view', false, true, false)).toBe(false);
        expect(shouldAutoRunDiff('diff-view', true, false, false)).toBe(false);
    });

    it('returns FALSE if diff tree already exists (avoid re-calc)', () => {
        expect(shouldAutoRunDiff('diff-view', true, true, true)).toBe(false);
    });

    it('returns FALSE when in setup mode (e.g. Landing Page Drag & Drop)', () => {
        expect(shouldAutoRunDiff('diff-setup', true, true, false)).toBe(false);
    });

    it('returns TRUE when in diff-view mode with files present and no existing tree', () => {
        expect(shouldAutoRunDiff('diff-view', true, true, false)).toBe(true);
    });
});

describe('XML Formatting Utilities', () => {
    
    it('formatXml adds newlines and indentation to nested tags', () => {
        const input = '<root><child>Text</child></root>';
        const result = formatXml(input);
        expect(result).toContain('\n');
        expect(result).toContain('  <child');
    });

    it('minifyXml removes whitespace and newlines between tags', () => {
        const input = '<root>\n  <child>Text</child>\n</root>';
        const expected = '<root><child>Text</child></root>';
        const result = minifyXml(input);
        expect(result).not.toContain('\n');
        expect(result).toBe(expected);
    });

    it('formatXml correctly handles self-closing tags', () => {
        const input = '<root><item id="1"/><item id="2"/></root>';
        const result = formatXml(input);
        expect(result).toContain('<item id="1"/>');
        expect(result).toContain('  <item id="1"/>');
    });

    it('formatXml handles basic comments (though regex is limited)', () => {
        const input = '<root><!-- comment --><child/></root>';
        const result = formatXml(input);
        expect(result).toContain('<!-- comment -->');
    });
});

describe('File Extension Helpers', () => {
    it('isXmlFile returns true for .xml and .rels extensions', () => {
        expect(isXmlFile('document.xml')).toBe(true);
        expect(isXmlFile('styles.xml')).toBe(true);
        expect(isXmlFile('_rels/.rels')).toBe(true);
    });

    it('isXmlFile returns false for non-xml files', () => {
        expect(isXmlFile('image.png')).toBe(false);
    });

    it('isImageFile returns true for common image extensions', () => {
        expect(isImageFile('image.png')).toBe(true);
        expect(isImageFile('photo.jpg')).toBe(true);
    });
    
    it('isImageFile returns false for non-image files', () => {
        expect(isImageFile('doc.xml')).toBe(false);
    });

    it('isBinaryFile returns true for binary assets', () => {
        expect(isBinaryFile('word/embeddings/oleObject1.bin')).toBe(true);
        expect(isBinaryFile('word/media/image.bin')).toBe(true);
        expect(isBinaryFile('word/fonts/font.otf')).toBe(true);
    });

    it('isBinaryFile returns false for xml or images', () => {
        expect(isBinaryFile('word/document.xml')).toBe(false);
        expect(isBinaryFile('word/media/image.png')).toBe(false);
    });
});

describe('Search Utilities', () => {
    const sampleText = "Hello world. Hello code. hello test.";

    it('calculateMatches finds all occurrences case-insensitively', () => {
        const matches = calculateMatches(sampleText, 'hello');
        expect(matches).toHaveLength(3);
        expect(matches[0].start).toBe(0); // Hello
        expect(matches[1].start).toBe(13); // Hello
        expect(matches[2].start).toBe(25); // hello
    });

    it('calculateMatches handles special regex characters in search term', () => {
        const text = "Look at this [special] text.";
        const matches = calculateMatches(text, '[special]');
        expect(matches).toHaveLength(1);
        expect(matches[0].start).toBe(13);
    });

    it('cycleIndex wraps around correctly (Next)', () => {
        expect(cycleIndex(0, 3, 'next')).toBe(1);
        expect(cycleIndex(2, 3, 'next')).toBe(0); // Wrap around
    });

    it('cycleIndex wraps around correctly (Prev)', () => {
        expect(cycleIndex(0, 3, 'prev')).toBe(2); // Wrap around
        expect(cycleIndex(2, 3, 'prev')).toBe(1);
    });

    it('replaceContent substitutes text at correct indices', () => {
        const text = "Hello world";
        const matches = calculateMatches(text, "world");
        const result = replaceContent(text, matches[0], "universe");
        expect(result).toBe("Hello universe");
    });
});

describe('Keyboard Shortcut Logic', () => {
    it('isSaveHotkey returns TRUE for Ctrl+S or Cmd+S', () => {
        const eventCtrl = { ctrlKey: true, key: 's', shiftKey: false } as KeyboardEvent;
        const eventCmd = { metaKey: true, key: 's', shiftKey: false } as KeyboardEvent;
        
        expect(isSaveHotkey(eventCtrl)).toBe(true);
        expect(isSaveHotkey(eventCmd)).toBe(true);
    });

    it('isSaveHotkey returns FALSE for "s" without modifiers', () => {
        const eventPlain = { key: 's' } as KeyboardEvent;
        expect(isSaveHotkey(eventPlain)).toBe(false);
    });

    it('isSaveAllHotkey returns TRUE for Ctrl+Shift+S or Cmd+Shift+S', () => {
        const eventCtrl = { ctrlKey: true, key: 's', shiftKey: true } as KeyboardEvent;
        const eventCmd = { metaKey: true, key: 's', shiftKey: true } as KeyboardEvent;
        
        expect(isSaveAllHotkey(eventCtrl)).toBe(true);
        expect(isSaveAllHotkey(eventCmd)).toBe(true);
    });
    
    it('isSaveAllHotkey returns FALSE without Shift key', () => {
         const eventNoShift = { ctrlKey: true, key: 's', shiftKey: false } as KeyboardEvent;
         expect(isSaveAllHotkey(eventNoShift)).toBe(false);
    });

    it('isFindHotkey returns TRUE for Ctrl+F or Cmd+F', () => {
        const event = { ctrlKey: true, key: 'f' } as KeyboardEvent;
        expect(isFindHotkey(event)).toBe(true);
    });

    it('isSidebarHotkey returns TRUE for Ctrl+B or Cmd+B', () => {
        const event = { ctrlKey: true, key: 'b' } as KeyboardEvent;
        const eventCmd = { metaKey: true, key: 'b' } as KeyboardEvent;
        expect(isSidebarHotkey(event)).toBe(true);
        expect(isSidebarHotkey(eventCmd)).toBe(true);
    });
});

describe('Markdown Utilities', () => {
    it('parseInlineStyles converts **bold** to HTML', () => {
        const input = "This is **bold** text";
        const expected = 'This is <strong class="text-blue-500 font-bold">bold</strong> text';
        expect(parseInlineStyles(input)).toBe(expected);
    });

    it('parseInlineStyles converts `code` to HTML', () => {
        const input = "Use the `code` format";
        const expected = 'Use the <code class="bg-blue-500/10 px-1 py-0.5 rounded text-[10px] font-mono text-blue-500 border border-blue-500/20">code</code> format';
        expect(parseInlineStyles(input)).toBe(expected);
    });

    it('parseMarkdownSegments splits text and code blocks', () => {
        const input = "Here is some code:\n```xml\n<root/>\n```\nEnd.";
        const segments = parseMarkdownSegments(input);
        
        expect(segments).toHaveLength(3);
        expect(segments[0].type).toBe('text');
        
        expect(segments[1].type).toBe('code');
        expect(segments[1].language).toBe('xml');
        expect(segments[1].content).toBe('<root/>');
        
        expect(segments[2].type).toBe('text');
        expect(segments[2].content).toContain('End.');
    });
});

describe('Tree Traversal Utilities', () => {
    const mockFileAdded: DiffNode = { name: 'new.xml', path: 'new.xml', isFolder: false, children: {}, status: 'added', hasChange: false };
    const mockFileModified: DiffNode = { name: 'mod.xml', path: 'mod.xml', isFolder: false, children: {}, status: 'modified', hasChange: false };
    const mockFileUnchanged: DiffNode = { name: 'same.xml', path: 'same.xml', isFolder: false, children: {}, status: 'unchanged', hasChange: false };

    const mockFolder: DiffNode = { 
        name: 'folder', 
        path: 'folder', 
        isFolder: true, 
        hasChange: true,
        children: {
            'new.xml': mockFileAdded,
            'mod.xml': mockFileModified,
            'same.xml': mockFileUnchanged
        }
    };

    it('getModifiedPaths recursively extracts paths of added or modified files', () => {
        const paths = getModifiedPaths(mockFolder);
        expect(paths).toContain('new.xml');
        expect(paths).toContain('mod.xml');
        expect(paths).not.toContain('same.xml');
        expect(paths).toHaveLength(2);
    });
});

describe('File Picker Utility', () => {
    it('uses showOpenFilePicker when available and returns the file', async () => {
        const mockFile = new File(['content'], 'test.docx');
        const mockHandle = { getFile: async () => mockFile };
        
        // Mock showOpenFilePicker on window
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const originalShowOpenFilePicker = (window as any).showOpenFilePicker;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (window as any).showOpenFilePicker = async () => [mockHandle];
        
        const result = await selectFileWithPicker(['.docx']);
        expect(result).toBe(mockFile);
        
        // Restore
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (window as any).showOpenFilePicker = originalShowOpenFilePicker;
    });
});