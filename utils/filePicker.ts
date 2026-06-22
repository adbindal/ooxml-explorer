/**
 * Triggers the browser's file picker.
 * Uses the modern File System Access API (showOpenFilePicker) if available,
 * which often bypasses corporate DLP upload scans on Chromium browsers.
 * Falls back to a programmatically clicked hidden input element if showOpenFilePicker is unavailable or fails.
 */
export const selectFileWithPicker = async (extensions: string[]): Promise<File | null> => {
    // 1. Try modern File System Access API (Chromium only)
    if (typeof window !== 'undefined' && 'showOpenFilePicker' in window) {
        try {
            // Map extensions to exact MIME types
            const accept: Record<string, string[]> = {};
            
            if (extensions.includes('.docx')) {
                accept['application/vnd.openxmlformats-officedocument.wordprocessingml.document'] = ['.docx'];
            }
            if (extensions.includes('.xlsx')) {
                accept['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'] = ['.xlsx'];
            }
            if (extensions.includes('.pptx')) {
                accept['application/vnd.openxmlformats-officedocument.presentationml.presentation'] = ['.pptx'];
            }

            // Fallback for general case if no match
            if (Object.keys(accept).length === 0) {
                accept['application/octet-stream'] = extensions;
            }

            const pickerOptions = {
                types: [
                    {
                        description: 'Office Open XML Documents',
                        accept
                    }
                ],
                excludeAcceptAllOption: false,
                multiple: false
            };
            
            // @ts-expect-error - showOpenFilePicker is a modern API not yet in standard lib types
            const [handle] = await window.showOpenFilePicker(pickerOptions);
            if (handle) {
                const file = await handle.getFile();
                return file;
            }
        } catch (e) {
            // If user cancelled, return null immediately
            if ((e as Error).name === 'AbortError') {
                return null;
            }
            // Log other errors (e.g. security block) and fall back
            console.warn("[FilePicker] showOpenFilePicker failed, falling back to input", e);
        }
    }
    
    // 2. Fallback to standard input element (Firefox, Safari, or blocked Chromium)
    return new Promise((resolve) => {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = extensions.join(',');
        input.style.display = 'none';
        
        input.onchange = (e) => {
            const file = (e.target as HTMLInputElement).files?.[0] || null;
            resolve(file);
        };
        
        document.body.appendChild(input);
        input.click();
        
        // Cleanup after click
        setTimeout(() => {
            try {
                if (document.body.contains(input)) {
                    document.body.removeChild(input);
                }
            } catch {
                // Ignore cleanup error
            }
        }, 10000);
    });
};
