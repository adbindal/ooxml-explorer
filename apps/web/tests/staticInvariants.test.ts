import { describe, it, expect } from '../services/browserTestRunner';
import fs from 'fs';
import path from 'path';

describe('Static Architecture Invariants', () => {
    it('enforces Monaco editor configurations statically', () => {
        const workspaceRoot = process.cwd();
        const filesToInspect = [
            path.join(workspaceRoot, 'views/EditorView.tsx'),
            path.join(workspaceRoot, 'views/DiffView.tsx')
        ];
        
        for (const filePath of filesToInspect) {
            if (fs.existsSync(filePath)) {
                const content = fs.readFileSync(filePath, 'utf8');
                // Invariant: Monaco instances MUST enable word wrap
                expect(content.includes("wordWrap: 'on'")).toBe(true);
            }
        }
    });

    it('enforces theme compliance statically (no hardcoded tailwind colors in components)', () => {
        const workspaceRoot = process.cwd();
        const directoriesToScan = [
            path.join(workspaceRoot, 'components'),
            path.join(workspaceRoot, 'views')
        ];
        
        // Files allowed to have custom static/hardcoded colors
        const whitelist = [
            'AIPanel.tsx',       // AI Assistant custom blue-accented branding and sparkles
            'ErrorBoundary.tsx', // Crash recovery fallback UI (must remain isolated from dynamic theme state)
            'FileTree.tsx',      // Recurse file tree requiring explicit green/red/amber diff status markers
            'LandingView.tsx',   // Specific branding gradient buttons and decorative elements
            'Logo.tsx',          // Core logo SVG path colors
            'ValidatorView.tsx'  // QA dashboard custom black console output styling
        ];
        
        // Regex to catch raw tailwind color classes like bg-red-500, text-blue-600, border-slate-700
        const rawColorRegex = /\b(bg|text|border|ring)-(red|blue|green|slate|zinc|gray|neutral|orange|amber|yellow|emerald|teal|cyan|sky|indigo|violet|purple|fuchsia|pink|rose)-(50|100|200|300|400|500|600|700|800|900|950)\b/g;
        
        const scanDirectory = (dir: string) => {
            if (!fs.existsSync(dir)) return;
            const entries = fs.readdirSync(dir, { withFileTypes: true });
            
            for (const entry of entries) {
                const fullPath = path.join(dir, entry.name);
                if (entry.isDirectory()) {
                    scanDirectory(fullPath);
                } else if (entry.isFile() && (entry.name.endsWith('.tsx') || entry.name.endsWith('.ts'))) {
                    if (whitelist.includes(entry.name)) continue;
                    
                    const content = fs.readFileSync(fullPath, 'utf8');
                    const matches = content.match(rawColorRegex);
                    
                    if (matches && matches.length > 0) {
                        throw new Error(
                            `Architectural Invariant Violation: Found hardcoded Tailwind colors in ${entry.name}. ` +
                            `All colors must be theme-aware or mapped via HSL tokens. Matches: ${matches.slice(0, 5).join(', ')}`
                        );
                    }
                }
            }
        };
        
        for (const dir of directoriesToScan) {
            scanDirectory(dir);
        }
    });

    it('prevents components from duplicating Zustand store state locally', () => {
        const workspaceRoot = process.cwd();
        const componentsDir = path.join(workspaceRoot, 'components');
        
        if (fs.existsSync(componentsDir)) {
            const entries = fs.readdirSync(componentsDir);
            for (const entry of entries) {
                if (entry.endsWith('.tsx')) {
                    const content = fs.readFileSync(path.join(componentsDir, entry), 'utf8');
                    // Invariant: Components should not manage local duplicate states like activePath or sidebarOpen
                    // that must remain globally synchronized via Zustand
                    const hasLocalActivePathState = /const\s+\[\s*activePath\s*,\s*\w+\s*\]\s*=\s*useState/.test(content);
                    if (hasLocalActivePathState) {
                        throw new Error(
                            `Architectural Invariant Violation in ${entry}: ` +
                            `Local state 'activePath' found. Shared paths MUST be managed in the Zustand store.`
                        );
                    }
                }
            }
        }
    });
});
