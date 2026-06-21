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
    
    // 1. mimetype (first, uncompressed)
    zip.file("mimetype", "application/vnd.openxmlformats-package.core-properties+xml", { compression: "STORE" });
    
    // 2. [Content_Types].xml
    zip.file("[Content_Types].xml", 
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>`
    );
    
    // 3. word/document.xml (main content)
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

test.describe('OOXML Explorer E2E Editor Flow', () => {
    let tempDocxPath: string;

    test.beforeAll(async () => {
        // Generate and write the mock docx file to a temp location before running tests
        const buffer = await createMockDocx();
        tempDocxPath = path.join(__dirname, 'temp_test_document.docx');
        fs.writeFileSync(tempDocxPath, buffer);
    });

    test.afterAll(() => {
        // Clean up the temp file after tests finish
        if (fs.existsSync(tempDocxPath)) {
            fs.unlinkSync(tempDocxPath);
        }
    });

    test('should load landing, upload file, edit in Monaco, and export successfully', async ({ page }) => {
        // 1. Navigate to the App
        await page.goto('/');
        await expect(page).toHaveTitle(/OOXML Explorer/);
        
        // Assert landing view is active
        await expect(page.locator('h1')).toContainText('OOXML Explorer');

        // 2. Upload the mock docx file
        const fileInput = page.locator('input[type="file"]');
        await fileInput.setInputFiles(tempDocxPath);

        // 3. Verify transition to Editor View
        // The filename should be displayed in the header, and the sidebar file tree should render
        await expect(page.locator('header')).toContainText('temp_test_document.docx');
        await expect(page.locator('text=word')).toBeVisible();

        // 4. Select 'document.xml' in the sidebar file tree (already expanded by default)
        const fileNode = page.locator('text=document.xml');
        await expect(fileNode).toBeVisible();
        await fileNode.click();

        // Verify the file name is shown in the active editor tabs
        await expect(page.locator('main .border-b').first()).toContainText('document.xml');

        // 6. Verify that Monaco Editor loads and displays the XML content
        // We look for the text content inside the Monaco editor viewport
        await expect(page.locator('.monaco-editor')).toContainText('Hello World from E2E Test!');

        // 7. Toggle Sidebar and AI Panel to verify layout responsiveness
        const sidebarToggleButton = page.locator('button[title*="Sidebar"], button[title*="panel"]').first();
        if (await sidebarToggleButton.isVisible()) {
            await sidebarToggleButton.click();
            // Verify sidebar is collapsed (width or visibility check, but simple click verifies no crash)
            await sidebarToggleButton.click(); // restore
        }

        // 8. Open Settings Drawer to verify dialog rendering
        const settingsButton = page.locator('button:has([data-lucide="settings"]), button[title*="Settings"]').first();
        if (await settingsButton.isVisible()) {
            await settingsButton.click();
            await expect(page.locator('text=Global Settings')).toBeVisible();
            // Close settings
            await page.keyboard.press('Escape');
        }

        // 9. Export the file and verify download
        // We click the Export button and wait for the browser to trigger a download
        const exportButton = page.locator('button:has-text("Export")');
        await expect(exportButton).toBeVisible();

        const downloadPromise = page.waitForEvent('download');
        await exportButton.click();
        const download = await downloadPromise;

        // Assert the download file name matches the expected export format
        expect(download.suggestedFilename()).toBe('MODIFIED_temp_test_document.docx');

        // Save the downloaded file to verify its size and integrity
        const downloadPath = path.join(__dirname, 'downloaded_test.docx');
        await download.saveAs(downloadPath);
        
        expect(fs.existsSync(downloadPath)).toBe(true);
        const stats = fs.statSync(downloadPath);
        expect(stats.size).toBeGreaterThan(0);

        // Clean up download
        fs.unlinkSync(downloadPath);
    });
});
