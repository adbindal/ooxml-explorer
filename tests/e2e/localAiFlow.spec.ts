import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

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
                        <w:t>Hello World from E2E Local AI Test!</w:t>
                    </w:r>
                </w:p>
            </w:body>
        </w:document>`
    );

    return await zip.generateAsync({ type: 'nodebuffer' });
}

test.describe('OOXML Explorer E2E Local AI & Fallback Flow', () => {
    let tempDocxPath: string;

    test.beforeEach(async () => {
        const buffer = await createMockDocx();
        tempDocxPath = path.join(__dirname, `temp_local_ai_${Math.random().toString(36).substring(7)}.docx`);
        fs.writeFileSync(tempDocxPath, buffer);
    });

    test.afterEach(() => {
        if (fs.existsSync(tempDocxPath)) {
            fs.unlinkSync(tempDocxPath);
        }
    });

    test('should show Local AI active pill when Prompt API is available in browser', async ({ page }) => {
        // 1. Inject mock LanguageModel before page loads
        await page.addInitScript(() => {
            window.LanguageModel = {
                availability: async () => 'available',
                create: async () => ({
                    prompt: async () => JSON.stringify({
                        summary: "Mock Local AI Summary from E2E!",
                        criticalIssues: [],
                        keyElements: []
                    }),
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

        // 4. Verify AI Panel opened and Cloud badge is shown by default
        await expect(page.locator('text=Cloud')).toBeVisible();

        // 5. Open AI Settings (inside the Assistant panel)
        const settingsButton = page.locator('button[title*="Settings"]').first();
        await settingsButton.click();

        // 6. Select 'Local AI (Nano)' in the segmented toggle
        const localAiToggleBtn = page.locator('button:has-text("Local AI (Nano)")');
        await localAiToggleBtn.click();

        // 7. Verify Local AI engine status guide shows "Active"
        await expect(page.locator('text=Local AI Engine')).toBeVisible();
        await expect(page.locator('text=Active')).toBeVisible();

        // 8. Go back to Assistant
        const backBtn = page.locator('button:has-text("Back to Assistant")');
        await backBtn.click();

        // 9. Verify the header pill updated to "Local"
        await expect(page.locator('text=⚡ Local')).toBeVisible();
    });

    test('should show Cloud Fallback pill in E2E when Prompt API is unsupported', async ({ page }) => {
        // Simulating Safari/Firefox by not injecting window.LanguageModel
        await page.goto('/');
        const fileInput = page.locator('input[type="file"]');
        await fileInput.setInputFiles(tempDocxPath);

        // Open AI Panel
        const sparklesButton = page.locator('button:has([data-lucide="sparkles"]), button[title*="AI"]').first();
        await sparklesButton.click();

        // Open AI Settings
        const settingsButton = page.locator('button[title*="Settings"]').first();
        await settingsButton.click();

        // Select 'Local AI (Nano)' in the segmented toggle
        const localAiToggleBtn = page.locator('button:has-text("Local AI (Nano)")');
        await localAiToggleBtn.click();

        // Verify Local AI engine status guide shows "Unsupported"
        await expect(page.locator('text=Unsupported')).toBeVisible();

        // Go back to Assistant
        const backBtn = page.locator('button:has-text("Back to Assistant")');
        await backBtn.click();

        // Verify the header pill updated to "Cloud Fallback"
        await expect(page.locator('text=Cloud Fallback')).toBeVisible();
    });
});
