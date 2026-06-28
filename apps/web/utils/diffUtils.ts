import { AppMode } from '../types';

/**
 * Determines whether the Diff View should automatically start the comparison process.
 *
 * Rules:
 * 1. Must have both files (Original & Modified).
 * 2. Must NOT have an existing Diff Tree (already computed).
 * 3. Mode must be 'diff-view' (triggered by Validator or direct action).
 *    - 'diff-setup' (triggered by Landing Drag & Drop) should NOT auto-run.
 */
export const shouldAutoRunDiff = (
    mode: AppMode,
    hasOriginal: boolean,
    hasModified: boolean,
    hasTree: boolean
): boolean => {
    if (!hasOriginal || !hasModified || hasTree) return false;
    
    // Only 'diff-view' implies an intent to run immediately.
    // 'diff-setup' implies the user is in the setup phase (swapping files, etc).
    return mode === 'diff-view';
};
