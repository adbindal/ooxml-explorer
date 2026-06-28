import { useMemo } from 'react';
import type { Monaco } from '@monaco-editor/react';
import type { ThemeClasses } from '../types';

export const getThemeClasses = (theme: 'dark' | 'light'): ThemeClasses => {
    const isDark = theme === 'dark';
    
    // Updated Dark Palette
    // Main BG: #0B1221 (Deepest Navy)
    // Secondary BG: #152238 (Surface Navy)
    // Brand/Accents: #4A89DC (Light Blue) & #1F3F70 (Classic Blue)

    return {
        bg: isDark ? 'bg-[#0B1221]' : 'bg-[#FFFFFF]',
        bgSec: isDark ? 'bg-[#152238]' : 'bg-[#FFFFFF]',
        bgSidebar: isDark ? 'bg-[#152238]' : 'bg-[#FFFFFF]',
        bgPanel: isDark ? 'bg-[#152238]' : 'bg-[#FFFFFF]',
        
        fg: isDark ? 'text-[#E2E8F0]' : 'text-[#1F3F70]', // Light slate text for dark mode
        fgMuted: isDark ? 'text-[#94A3B8]' : 'text-[#1F3F70]/70',
        
        border: isDark ? 'border-[#1F3F70]/50' : 'border-[#1F3F70]/20',
        
        hover: isDark ? 'hover:bg-[#1F3F70]/40' : 'hover:bg-[#4A89DC]/10',
        hoverText: isDark ? 'hover:text-[#4A89DC]' : 'hover:text-[#4A89DC]',
        
        activeTree: isDark ? 'bg-[#1F3F70]/60 text-[#FFFFFF]' : 'bg-[#4A89DC]/20 text-[#1F3F70] font-medium',
        
        activeTab: isDark ? 'bg-[#0B1221] text-[#4A89DC] border-t-2 border-t-[#4A89DC]' : 'bg-[#FFFFFF] text-[#4A89DC] border-t-2 border-t-[#4A89DC] shadow-sm',
        inactiveTab: isDark ? 'bg-[#152238] text-[#94A3B8] hover:bg-[#1F3F70]/40' : 'bg-[#FFFFFF] text-[#1F3F70]/60 hover:bg-[#4A89DC]/10',
        
        card: isDark ? 'bg-[#152238] border-[#1F3F70]' : 'bg-[#FFFFFF] border-[#1F3F70]/20 shadow-sm',
        
        input: isDark ? 'bg-[#0B1221] border-[#1F3F70] text-[#E2E8F0] focus:border-[#4A89DC]' : 'bg-[#FFFFFF] border-[#1F3F70]/30 text-[#1F3F70] focus:border-[#4A89DC]',
        
        icon: isDark ? 'text-[#94A3B8]' : 'text-[#1F3F70]/60',
        
        monaco: isDark ? 'ooxml-dark' : 'light'
    };
};

export const useThemeClasses = (theme: 'dark' | 'light'): ThemeClasses => {
  return useMemo(() => getThemeClasses(theme), [theme]);
};

/**
 * Defines the custom 'ooxml-dark' and 'ooxml-light' themes on the provided Monaco instance.
 * Shared between EditorView and DiffView for consistency.
 */
export const defineMonacoThemes = (monaco: Monaco) => {
    // Detailed Dark Theme to match app "Deep Navy"
    monaco.editor.defineTheme('ooxml-dark', {
        base: 'vs-dark',
        inherit: true,
        rules: [],
        colors: {
            'editor.background': '#0B1221',
            'editor.foreground': '#E2E8F0',
            'editor.lineHighlightBackground': '#152238',
            'minimap.background': '#0B1221',
            'editorLineNumber.foreground': '#1F3F70',
            'editorLineNumber.activeForeground': '#4A89DC',
            
            // Widget Styling for Dark Mode Consistency
            'editorWidget.background': '#152238',
            'editorWidget.border': '#1F3F70',
            'editorWidget.foreground': '#E2E8F0',
            'editorWidget.resizeBorder': '#4A89DC',

            // Find Widget Inputs
            'input.background': '#0B1221',
            'input.border': '#1F3F70',
            'input.foreground': '#E2E8F0',
            'inputOption.activeBorder': '#4A89DC',
            'focusBorder': '#4A89DC',
            
            // Lists & Scrollbars
            'list.hoverBackground': '#1F3F70',
            'list.activeSelectionBackground': '#4A89DC30',
            'list.activeSelectionForeground': '#FFFFFF',
            'scrollbarSlider.background': '#4A89DC20',
            'scrollbarSlider.hoverBackground': '#4A89DC50',
            'scrollbarSlider.activeBackground': '#4A89DC',
        }
    });

    // Detailed Light Theme
    monaco.editor.defineTheme('ooxml-light', {
        base: 'vs',
        inherit: true,
        rules: [],
        colors: {
             'editor.background': '#FFFFFF',
             'editor.foreground': '#1F3F70',
             'editor.lineHighlightBackground': '#F0F9FF',
             'editorLineNumber.foreground': '#94A3B8',
             'editorLineNumber.activeForeground': '#4A89DC',
             'minimap.background': '#FFFFFF',

             'editorWidget.background': '#F8FAFC',
             'editorWidget.border': '#E2E8F0',
             'editorWidget.foreground': '#1F3F70',
             
             'input.background': '#FFFFFF',
             'input.border': '#CBD5E1',
             'input.foreground': '#1F3F70',
             'inputOption.activeBorder': '#4A89DC',
             'focusBorder': '#4A89DC',

             'list.hoverBackground': '#F1F5F9',
             'list.activeSelectionBackground': '#E0F2FE',
             'list.activeSelectionForeground': '#1F3F70',
        }
    });
};