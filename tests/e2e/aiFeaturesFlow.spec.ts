import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

test.use({
    contextOptions: {
        extraHTTPHeaders: {
            'Cache-Control': 'no-cache, no-store, must-revalidate',
            'Pragma': 'no-cache',
            'Expires': '0'
        }
    }
});

// Helper to generate a minimal, valid OOXML package dynamically
async function createMockDocx(): Promise<Buffer> {
    const zip = new JSZip();
    zip.file("mimetype", "application/vnd.openxmlformats-package.core-properties+xml", { compression: "STORE" });
    zip.file("[Content_Types].xml", 
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>`
    );
    zip.file("word/document.xml", 
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:r>
                        <w:t>Hello World from E2E Test!</w:t>
                    </w:r>
                </w:p>
            </w:body>
        </w:document>`
    );
    return await zip.generateAsync({ type: 'nodebuffer' });
}

test.describe('OOXML Explorer E2E AI Features & DLP Flow', () => {
    let tempDocxPath: string;

    test.beforeEach(async () => {
        const buffer = await createMockDocx();
        tempDocxPath = path.join(__dirname, `temp_ai_features_${Math.random().toString(36).substring(7)}.docx`);
        fs.writeFileSync(tempDocxPath, buffer);
    });

    test.afterEach(() => {
        if (fs.existsSync(tempDocxPath)) {
            fs.unlinkSync(tempDocxPath);
        }
    });

    test('should enable DLP Mode by default and show security warning if Local AI is unsupported', async ({ page }) => {
        page.on('console', msg => console.log('BROWSER CONSOLE:', msg.text()));
        page.on('pageerror', err => console.log('BROWSER ERROR:', err.message));
        page.on('response', async response => {
            if (response.url().includes('appStore.ts') || response.url().includes('appStore?')) {
                console.log('STORE FILE CONTENT:', await response.text().catch(() => 'CANNOT READ'));
            }
            if (response.url().includes('AIPanel.tsx') || response.url().includes('AIPanel?')) {
                console.log('AIPANEL FILE CONTENT:', await response.text().catch(() => 'CANNOT READ'));
            }
        });

        // Simulating unsupported Local AI by not injecting window.LanguageModel
        await page.goto('/');
        const fileInput = page.locator('input[type="file"]');
        await fileInput.setInputFiles(tempDocxPath);

        // Open AI Panel
        const sparklesButton = page.locator('button:has([data-lucide="sparkles"]), button[title*="AI"]').first();
        await sparklesButton.click();

        // Print the HTML of the AI Panel header
        const headerHtml = await page.locator('.h-12.border-b').first().innerHTML().catch(() => 'NOT FOUND');
        console.log('AI PANEL HEADER HTML:', headerHtml);

        // Verify the header pill shows "DLP Shield" (since dlpMode is true, it forces local)
        await expect(page.locator('text=DLP Shield')).toBeVisible();

        // Trigger any AI action (e.g., Technical Analysis or double-click to explain)
        const explainBtn = page.locator('button:has-text("Technical Analysis")').first();
        if (await explainBtn.isVisible()) {
            await explainBtn.click();
            // Verify that the output area displays the DLP block error
            await expect(page.locator('text=Analysis Failed')).toBeVisible();
            await expect(page.locator('text=DLP_BLOCK')).toBeVisible();
        }
    });

    test('should explain selected tag when Local AI is active and tag is double-clicked', async ({ page }) => {
        // 1. Inject mock LanguageModel before page loads
        await page.addInitScript(() => {
            window.LanguageModel = {
                availability: async () => 'available',
                create: async () => ({
                    prompt: async () => "Mocked local explanation of the paragraph element.",
                    destroy: () => {}
                })
            };
        });

        // 2. Navigate and Upload
        await page.goto('/');
        const fileInput = page.locator('input[type="file"]');
        await fileInput.setInputFiles(tempDocxPath);

        // 3. Open AI Panel
        const sparklesButton = page.locator('button:has([data-lucide="sparkles"]), button[title*="AI"]').first();
        await sparklesButton.click();

        // Switch to Local AI (Nano) in Settings to bypass API key screen
        const settingsBtn = page.locator('button[title="Settings"]').first();
        await settingsBtn.click();
        const localAiToggleBtn = page.locator('button:has-text("Local AI (Nano)")');
        await localAiToggleBtn.click();
        await settingsBtn.click(); // Close settings view

        // 4. Select 'document.xml' in the sidebar
        const fileNode = page.locator('text=document.xml');
        await fileNode.click();

        // 5. Select/Double-click a tag name inside Monaco Editor
        // We'll look for the "body" tag text inside Monaco and double-click it
        const tagText = page.locator('.monaco-editor >> text=body').first();
        await expect(tagText).toBeVisible();
        await tagText.dblclick();

        // 6. Wait for the debounce (300ms) and verify the "Explain Selected Tag" button appears
        const explainTagBtn = page.locator('button:has-text("Explain Selected Tag <body")').first();
        await expect(explainTagBtn).toBeVisible();

        // 7. Click the button and verify the mocked local explanation is rendered
        await explainTagBtn.click();
        await expect(page.locator('text=XML Element Explanation')).toBeVisible();
        await expect(page.locator('text=Mocked local explanation')).toBeVisible();
    });
});
