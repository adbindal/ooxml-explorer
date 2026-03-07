
export interface SearchMatch {
    start: number;
    end: number;
}

/**
 * reliable search that handles case-insensitivity and special regex characters safely.
 */
export const calculateMatches = (content: string, term: string): SearchMatch[] => {
    if (!content || !term) return [];
    
    try {
        // Escape special regex characters to prevent crashes or unintended matching
        const escapedTerm = term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const regex = new RegExp(escapedTerm, 'gi');
        
        const matches: SearchMatch[] = [];
        let match;
        
        // Use exec in a loop to find all occurrences with indices
        while ((match = regex.exec(content)) !== null) {
            matches.push({ 
                start: match.index, 
                end: match.index + match[0].length 
            });
        }
        return matches;
    } catch (e) {
        console.error("Search regex error", e);
        return [];
    }
};

/**
 * Calculates the next index in a circular list (Next/Previous navigation).
 */
export const cycleIndex = (currentIndex: number, total: number, direction: 'next' | 'prev'): number => {
    if (total === 0) return -1;
    
    if (direction === 'next') {
        return (currentIndex + 1) % total;
    } else {
        return (currentIndex - 1 + total) % total;
    }
};

/**
 * Performs a string replacement on content (Logic for future Replace feature).
 */
export const replaceContent = (content: string, match: SearchMatch, replacement: string): string => {
    const before = content.slice(0, match.start);
    const after = content.slice(match.end);
    return before + replacement + after;
};
