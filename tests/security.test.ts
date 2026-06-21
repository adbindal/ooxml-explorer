import { describe, it, expect } from '../services/browserTestRunner';
import { formatXml } from '../utils/xmlUtils';
import { loadZipFile } from '../services/zipService';
import { setApiKey, clearApiKey } from '../services/geminiService';
import { getLogString, setDebugMode, initDebugService } from '../services/debugService';
import JSZip from 'jszip';

describe('Security & Vulnerability Audits', () => {

    it('sanitizes and neutralizes path traversal sequences in zip entry paths', async () => {
        // Setup a mock zip containing path traversal sequences
        const zip = new JSZip();
        
        // standard files
        zip.file("mimetype", "application/vnd.openxmlformats-package.core-properties+xml");
        zip.file("[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'></Types>");
        
        // malicious files trying to breakout or corrupt layout
        zip.file("word/../../malicious_breakout.xml", "<malicious/>");
        zip.file("word/./current_dir.xml", "<current/>");
        zip.file("../outside_root.xml", "<outside/>");

        const blob = await zip.generateAsync({ type: 'blob' });
        const file = new File([blob], 'security_test.docx');

        // Load the zip file
        const { tree, flat } = await loadZipFile(file);

        // --- ASSERTIONS ---
        // 1. Traverse paths must NOT create outside root structures.
        // 2. The flat map should only contain normalized, safe paths, or traverse paths should be neutralized.
        
        const flatPaths = Object.keys(flat);
        
        // Verify that traversal files are either rejected, normalized, or safely structured.
        // In our sanitized zip service, any entry containing '../' or './' must be ignored or bound.
        flatPaths.forEach(path => {
            expect(path.includes('../')).toBe(false);
            expect(path.includes('./')).toBe(false);
        });

        // The virtual tree root should never have children with traversal names like '..' or '.'
        expect(tree.children['..']).toBeUndefined();
        expect(tree.children['.']).toBeUndefined();
    });

    it('formats malformed or invalid XML gracefully without throwing uncaught exceptions', () => {
        // Malformed XML strings that would cause strict parsers to crash
        const malformedXml1 = "<root><child>Unclosed tag";
        const malformedXml2 = "<root xmlns:h='http://www.w3.org/TR/html4/'><h:table><h:tr>Missing namespace declarations or mismatch";
        const binaryDataInXml = "<bin>\x00\x01\x02\x03\x04\x05Malicious Binary Infiltration</bin>";

        // Verify they fail gracefully and return valid non-empty strings instead of throwing unhandled errors
        expect(formatXml(malformedXml1)).toBeDefined();
        expect(formatXml(malformedXml1).length).toBeGreaterThan(0);

        expect(formatXml(malformedXml2)).toBeDefined();
        expect(formatXml(malformedXml2).length).toBeGreaterThan(0);

        expect(formatXml(binaryDataInXml)).toBeDefined();
        expect(formatXml(binaryDataInXml).length).toBeGreaterThan(0);
        
        // Edge cases
        expect(formatXml('')).toBe('');
        expect(formatXml(null as unknown as string)).toBe('');
    });

    it('resists Billion Laughs recursive entity expansion attacks without locking threads', () => {
        // Billion Laughs attack payload
        const billionLaughsPayload = `<?xml version="1.0"?>
        <!DOCTYPE lolz [
         <!ENTITY lol "lol">
         <!ELEMENT lolz (#PCDATA)>
         <!ENTITY lol1 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">
         <!ENTITY lol2 "&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;">
         <!ENTITY lol3 "&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;">
        ]>
        <lolz>&lol3;</lolz>`;

        // Verify that formatting this does not crash the app, and fails safely or strips entities
        const startTime = Date.now();
        const formatted = formatXml(billionLaughsPayload);
        const duration = Date.now() - startTime;

        // Must complete instantly (should not cause exponential expansion lockup)
        expect(duration < 500).toBe(true);
        expect(formatted).toBeDefined();
    });

    it('actively scrubs the Gemini API Key from all debug logs and dumps', () => {
        const originalLog = console.log;
        const originalWarn = console.warn;
        const originalError = console.error;

        // 1. Enable debug capture and set a mock API Key
        setDebugMode(true);
        const mockSecretKey = "secret-gemini-key-xyz-12345-fort-knox";
        setApiKey(mockSecretKey);

        try {
            // 2. Initialize debug service to patch console methods
            initDebugService();

            // 3. Log some error messages containing the secret key
            console.error("Failed to connect to Gemini API with key: " + mockSecretKey);
            console.warn("Retrying request with authentication header Bearer " + mockSecretKey);
            console.log("Debug dump context: ", { apiKey: mockSecretKey, host: "api.gemini.google.com" });

            // 4. Retrieve the debug log dump
            const logDump = getLogString();

            // --- ASSERTIONS ---
            // The secret key must NEVER be present in the log dump!
            const containsSecretKey = logDump.includes(mockSecretKey);
            expect(containsSecretKey).toBe(false);

            // Verify the scrubbing placeholder is present
            expect(logDump.includes("[SCRUBBED_API_KEY]")).toBe(true);
        } finally {
            // 5. Clean up keys and restore original console methods to isolate tests
            clearApiKey();
            setDebugMode(false);
            console.log = originalLog;
            console.warn = originalWarn;
            console.error = originalError;
        }
    });

});
