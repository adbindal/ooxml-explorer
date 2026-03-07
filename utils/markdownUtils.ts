
export interface MarkdownSegment {
    type: 'code' | 'text';
    content: string;
    language?: string;
}

/**
 * Parses inline markdown styles (bold, inline code) into HTML strings.
 * This handles the transformation of **text** and `code` into specific tailwind classes.
 */
export const parseInlineStyles = (text: string): string => {
    let result = text;
    // Bold: **text** -> strong
    result = result.replace(/\*\*(.*?)\*\*/g, '<strong class="text-blue-500 font-bold">$1</strong>');
    // Inline Code: `text` -> code
    result = result.replace(/`([^`]+)`/g, '<code class="bg-blue-500/10 px-1 py-0.5 rounded text-[10px] font-mono text-blue-500 border border-blue-500/20">$1</code>');
    return result;
};

/**
 * Splits raw markdown content into structured segments (Code Blocks vs Text).
 * This makes rendering logic in React much cleaner.
 */
export const parseMarkdownSegments = (content: string): MarkdownSegment[] => {
    if (!content) return [];
    
    // Split by code block regex, capturing the delimiter to preserve order
    const parts = content.split(/(```[\s\S]*?```)/g);
    
    return parts.map(part => {
        if (part.startsWith('```')) {
            // Extract language and code content
            const match = part.match(/```(\w*)\n([\s\S]*?)```/);
            if (match) {
                return {
                    type: 'code',
                    language: match[1] || '',
                    content: match[2].trim()
                };
            }
            // Fallback for malformed or simple blocks
            return {
                type: 'code',
                language: '',
                content: part.replace(/```/g, '').trim()
            };
        }
        // Text block (may contain whitespace, filtered later if needed)
        return { type: 'text', content: part };
    });
};
