
/**
 * Checks if the event matches standard "Save" shortcuts (Ctrl+S / Cmd+S).
 */
export const isSaveHotkey = (e: KeyboardEvent): boolean => {
    const isCmdOrCtrl = !!(e.metaKey || e.ctrlKey);
    return isCmdOrCtrl && e.key?.toLowerCase() === 's' && !e.shiftKey;
};

/**
 * Checks if the event matches "Save All" (Ctrl+Shift+S / Cmd+Shift+S).
 */
export const isSaveAllHotkey = (e: KeyboardEvent): boolean => {
    const isCmdOrCtrl = !!(e.metaKey || e.ctrlKey);
    return isCmdOrCtrl && e.key?.toLowerCase() === 's' && e.shiftKey;
};

/**
 * Checks if the event matches "Find" (Ctrl+F / Cmd+F).
 */
export const isFindHotkey = (e: KeyboardEvent): boolean => {
    const isCmdOrCtrl = !!(e.metaKey || e.ctrlKey);
    return isCmdOrCtrl && e.key?.toLowerCase() === 'f';
};

/**
 * Checks if the event matches "Toggle Sidebar" (Ctrl+B / Cmd+B).
 */
export const isSidebarHotkey = (e: KeyboardEvent): boolean => {
    const isCmdOrCtrl = !!(e.metaKey || e.ctrlKey);
    return isCmdOrCtrl && e.key?.toLowerCase() === 'b';
};
