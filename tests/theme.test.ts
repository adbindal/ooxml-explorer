
import { describe, it, expect, vi } from '../services/browserTestRunner';
import { defineMonacoThemes, getThemeClasses } from '../utils/theme';

describe('Theme Utilities', () => {
    it('getThemeClasses returns correct classes for dark mode', () => {
        const classes = getThemeClasses('dark');
        expect(classes.bg).toBe('bg-[#0B1221]');
        expect(classes.fg).toBe('text-[#E2E8F0]');
    });

    it('getThemeClasses returns correct classes for light mode', () => {
        const classes = getThemeClasses('light');
        expect(classes.bg).toBe('bg-[#FFFFFF]');
        expect(classes.fg).toBe('text-[#1F3F70]');
    });

    it('defineMonacoThemes registers both light and dark themes', () => {
        // Mock the Monaco API structure
        const defineThemeSpy = vi.fn();
        const mockMonaco = {
            editor: {
                defineTheme: defineThemeSpy
            }
        };

        // Execute
        defineMonacoThemes(mockMonaco);

        // Verify it was called twice (once for dark, once for light)
        expect(defineThemeSpy.mock.calls.length).toBe(2);

        // Extract arguments to verify specific theme details
        const calls = defineThemeSpy.mock.calls;
        const themeNames = calls.map(c => c[0]);
        
        // Check names
        expect(themeNames).toContain('ooxml-dark');
        expect(themeNames).toContain('ooxml-light');

        // Check Dark Theme Configuration
        const darkDef = calls.find(c => c[0] === 'ooxml-dark')[1];
        expect(darkDef.base).toBe('vs-dark');
        expect(darkDef.inherit).toBe(true);
        // Verify a specific brand color override exists
        expect(darkDef.colors['editor.background']).toBe('#0B1221');
        expect(darkDef.colors['editorLineNumber.activeForeground']).toBe('#4A89DC');

        // Check Light Theme Configuration
        const lightDef = calls.find(c => c[0] === 'ooxml-light')[1];
        expect(lightDef.base).toBe('vs');
        expect(lightDef.inherit).toBe(true);
        expect(lightDef.colors['editor.background']).toBe('#FFFFFF');
        expect(lightDef.colors['focusBorder']).toBe('#4A89DC');
    });
});
