import { describe, it, expect } from '../services/browserTestRunner';
import { isNodeVisible } from '../utils/treeUtils';
import { DiffNode } from '../types';

describe('Tree Visibility Logic', () => {

    const mockFileUnchanged: DiffNode = { name: 'same.xml', path: 'same.xml', isFolder: false, children: {}, status: 'unchanged', hasChange: false };
    const mockFileAdded: DiffNode = { name: 'new.xml', path: 'new.xml', isFolder: false, children: {}, status: 'added', hasChange: false };
    const mockFileModified: DiffNode = { name: 'mod.xml', path: 'mod.xml', isFolder: false, children: {}, status: 'modified', hasChange: false };
    
    // Folder containing mixed content
    const mockFolder: DiffNode = { 
        name: 'word', 
        path: 'word', 
        isFolder: true, 
        hasChange: true,
        children: {
            'same.xml': mockFileUnchanged,
            'new.xml': mockFileAdded
        }
    };

    it('hides unchanged files when showUnchanged is false (Diff Mode)', () => {
        expect(isNodeVisible(mockFileUnchanged, '', true, false)).toBe(false);
        expect(isNodeVisible(mockFileAdded, '', true, false)).toBe(true);
        expect(isNodeVisible(mockFileModified, '', true, false)).toBe(true);
    });

    it('shows unchanged files when showUnchanged is true (Diff Mode)', () => {
        expect(isNodeVisible(mockFileUnchanged, '', true, true)).toBe(true);
    });

    it('filters visible nodes by search query string', () => {
        // Normal mode (no diff)
        expect(isNodeVisible(mockFileUnchanged, 'same', false, true)).toBe(true);
        expect(isNodeVisible(mockFileUnchanged, 'foobar', false, true)).toBe(false);
    });

    it('enforces Diff Filters (hide unchanged) even if node matches search query', () => {
        // User searched for "same", but also toggled "Diffs Only".
        // The file "same.xml" matches "same", BUT it is 'unchanged'.
        // It should NOT be visible.
        expect(isNodeVisible(mockFileUnchanged, 'same', true, false)).toBe(false);
    });

    it('shows folders if they contain matching children (Recursive Search)', () => {
        // Search "new" in Diff Mode (Diffs Only)
        // 'word' folder matches because it contains 'new.xml' which is added.
        expect(isNodeVisible(mockFolder, 'new', true, false)).toBe(true);
    });

    it('hides folders if children match search but are hidden by diff filter', () => {
        // Search "same" in Diff Mode (Diffs Only)
        // 'word' contains 'same.xml'. 'same.xml' matches name, BUT is hidden by diff filter.
        // Therefore 'word' should also be hidden (assuming it has no other visible matches).
        expect(isNodeVisible(mockFolder, 'same', true, false)).toBe(false);
    });
});