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
                <w:p><w:r><w:t>Live Post-Deployment Validator Check</w:t></w:r></w:p>
            </w:body>
        </w:document>`
    );
    return await zip.generateAsync({ type: 'nodebuffer' });
}

test.describe('OOXML Explorer Live Post-Deployment Validator Check', () => {
    let tempDocxPath: string;

    test.beforeAll(async () => {
        const buffer = await createMockDocx();
        tempDocxPath = path.join(__dirname, 'temp_validator_doc.docx');
        fs.writeFileSync(tempDocxPath, buffer);
    });

    test.afterAll(() => {
        if (fs.existsSync(tempDocxPath)) {
            fs.unlinkSync(tempDocxPath);
        }
    });

    test('should load the website, navigate to the validator, upload file, run tests, and assert 100% success', async ({ page, baseURL }) => {
        // Capture browser console logs and page errors for diagnostic debugging
        page.on('console', msg => {
            console.log(`[Browser Console] ${msg.type()}: ${msg.text()}`);
        });
        page.on('pageerror', err => {
            console.error(`[Browser PageError] ${err.message}`);
        });

        // Use the configured baseURL (will be the live Cloudflare URL in production, or localhost in development)
        const targetUrl = baseURL || 'http://127.0.0.1:5173';
        console.log(`🚀 Executing Live Post-Deployment Smoke Test on target: ${targetUrl}`);

        // 1. Load the target URL
        await page.goto(targetUrl);
        await expect(page).toHaveTitle(/OOXML Explorer/);

        // 2. Click the link to transition to the Validator View
        const validatorButton = page.locator('text=Go to Validator & Test Runner');
        await expect(validatorButton).toBeVisible();
        await validatorButton.click();

        // Verify transition to Validator View
        await expect(page.locator('h2')).toContainText('QA Validation Suite');

        // 3. Upload the dynamically generated mock docx file to the Primary File input
        // There are two file inputs (0 = primary file, 1 = comparison file)
        const fileInputs = page.locator('input[type="file"]');
        await expect(fileInputs).toHaveCount(2);
        
        await fileInputs.nth(0).setInputFiles(tempDocxPath);

        // 4. Click the "Run Unit Tests" button (should now be enabled!)
        const runButton = page.locator('button:has-text("Run Unit Tests")');
        await expect(runButton).toBeEnabled();
        await runButton.click();

        // 5. Wait for the browser runner to execute the test suite
        // We look for the final success message in the console output pane
        // Timeout is set to 10 seconds to allow Monaco loading and full suite execution
        const consolePane = page.locator('div.font-mono');
        await expect(consolePane).toBeVisible();
        
        // Assert that the console output displays that the unit tests passed successfully
        await expect(consolePane).toContainText('Unit Tests Passed', { timeout: 10000 });

        // Assert that no errors were printed to the browser console logs
        const containsError = await consolePane.locator('.text-red-400').count();
        expect(containsError).toBe(0);

        console.log('✅ Live Post-Deployment Smoke Test Passed Perfectly!');
    });
});
