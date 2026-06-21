import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Helper to generate a mock OOXML package dynamically with custom text
async function createMockDocx(text: string): Promise<Buffer> {
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
                        <w:t>${text}</w:t>
                    </w:r>
                </w:p>
            </w:body>
        </w:document>`
    );
    return await zip.generateAsync({ type: 'nodebuffer' });
}

test.describe('OOXML Explorer E2E Diff Flow', () => {
    let originalDocxPath: string;
    let modifiedDocxPath: string;

    test.beforeAll(async () => {
        // Generate original and modified mock documents
        const bufA = await createMockDocx("Hello from original document!");
        const bufB = await createMockDocx("Hello from modified document! Added text.");

        originalDocxPath = path.join(__dirname, 'temp_original.docx');
        modifiedDocxPath = path.join(__dirname, 'temp_modified.docx');

        fs.writeFileSync(originalDocxPath, bufA);
        fs.writeFileSync(modifiedDocxPath, bufB);
    });

    test.afterAll(() => {
        // Clean up temp files
        if (fs.existsSync(originalDocxPath)) fs.unlinkSync(originalDocxPath);
        if (fs.existsSync(modifiedDocxPath)) fs.unlinkSync(modifiedDocxPath);
    });

    test('should navigate to diff, upload two files, run comparison, and exercise diff views', async ({ page }) => {
        // 1. Navigate to landing page
        await page.goto('/');

        // 2. Go to Compare Documents page
        const diffCard = page.locator('text=Diff Files');
        await expect(diffCard).toBeVisible();
        await diffCard.click();

        // Verify transition to diff setup page
        await expect(page.locator('h2')).toContainText('Compare Documents');

        // 3. Upload original and modified files
        // We select by index (0 = original input, 1 = modified input)
        const fileInputs = page.locator('input[type="file"]');
        await expect(fileInputs).toHaveCount(2);

        await fileInputs.nth(0).setInputFiles(originalDocxPath);
        await fileInputs.nth(1).setInputFiles(modifiedDocxPath);

        // Verify that filenames are rendered on the boxes
        await expect(page.locator('text=temp_original.docx')).toBeVisible();
        await expect(page.locator('text=temp_modified.docx')).toBeVisible();

        // 4. Click the Compare button to run comparison
        const compareButton = page.locator('button:has-text("Run Comparison")').first();
        await expect(compareButton).toBeVisible();
        await compareButton.click();

        // 5. Select document.xml in the sidebar file tree (already expanded by default)
        const docFileNode = page.locator('text=document.xml');
        await expect(docFileNode).toBeVisible();

        // 6. Select document.xml to open it in Monaco Diff Editor
        await docFileNode.click();

        // 7. Verify Monaco Diff Editor displays original (left) and modified (right) contents
        const diffEditor = page.locator('.monaco-diff-editor');
        await expect(diffEditor).toBeVisible();
        await expect(diffEditor).toContainText('Hello from original');
        await expect(diffEditor).toContainText('Hello from modified');

        // 8. Toggle Diff view modes (Split / Inline)
        const splitButton = page.locator('button[title*="Split"], button:has-text("Split")').first();
        const inlineButton = page.locator('button[title*="Inline"], button:has-text("Inline")').first();
        
        if (await inlineButton.isVisible()) {
            await inlineButton.click();
            // Verify no crashes occur, layout transitions
            if (await splitButton.isVisible()) {
                await splitButton.click();
            }
        }

        // 9. Cycle through diff differences
        const nextDiffButton = page.locator('button[title*="Next"], button:has-text("Next")').first();
        const prevDiffButton = page.locator('button[title*="Prev"], button:has-text("Prev")').first();

        if (await nextDiffButton.isVisible() && await nextDiffButton.isEnabled()) {
            await nextDiffButton.click();
            // Clicks successfully
            if (await prevDiffButton.isVisible()) {
                await prevDiffButton.click();
            }
        }

        // 10. Click swap files button to swap original/modified and re-run
        // First go back to setup by clicking back button
        const backButton = page.locator('button:has([data-lucide="arrow-left"]), button[title*="Back"]').first();
        if (await backButton.isVisible()) {
            await backButton.click();
            // Verify we are back on the compare page
            await expect(page.locator('h2')).toContainText('Compare Documents');
        }
    });
});
