
import { FileNode, DiffNode } from '../types';

/**
 * Checks if a specific node passes the "Diffs Only" filter.
 * - Files: Must be added, modified, or deleted.
 * - Folders: Must have `hasChange: true`.
 */
export const passesDiffFilter = (
    node: FileNode | DiffNode, 
    isDiffMode: boolean, 
    showUnchanged: boolean
): boolean => {
    if (!isDiffMode || showUnchanged) return true;
    
    // In "Diffs Only" mode:
    if (node.isFolder) {
        // Folders are visible if they contain changes (DiffNode property)
        // We cast to DiffNode safely because hasChange is optional on FileNode but present on DiffNode
        return (node as DiffNode).hasChange === true;
    } else {
        // Files are visible if status is not unchanged
        return node.status !== 'unchanged';
    }
};

/**
 * Determines if a node should be visible based on both Diff Filters AND Search Query.
 * This implements the AND condition: (Passes Diff Filter) AND (Matches Search OR Has Matching Children)
 */
export const isNodeVisible = (
    node: FileNode | DiffNode,
    searchQuery: string,
    isDiffMode: boolean,
    showUnchanged: boolean
): boolean => {
    // 1. Diff Filter Check (Strict Pre-requisite)
    if (!passesDiffFilter(node, isDiffMode, showUnchanged)) {
        return false;
    }

    // 2. Search Check
    if (!searchQuery) return true;

    const nameMatches = node.name.toLowerCase().includes(searchQuery.toLowerCase());
    
    // If it's a file, it must match name (since we already passed diff filter)
    if (!node.isFolder) {
        return nameMatches;
    }

    // If it's a folder:
    // It is visible if NAME matches OR if ANY child matches (recursively)
    if (nameMatches) return true;

    if (node.children) {
        return Object.values(node.children).some(child => 
            isNodeVisible(child, searchQuery, isDiffMode, showUnchanged)
        );
    }

    return false;
};

/**
 * Recursively traverses a DiffNode tree to collect paths of all modified files.
 * Used for "Explain Changes" feature to get context files.
 */
export const getModifiedPaths = (node: DiffNode): string[] => {
    let paths: string[] = [];
    if (!node.isFolder && node.status !== 'unchanged') {
        paths.push(node.path);
    }
    if (node.children) {
        Object.values(node.children).forEach(child => {
            paths = [...paths, ...getModifiedPaths(child as DiffNode)];
        });
    }
    return paths;
};
