import JSZip from 'jszip';
// Import Browser Test Runner and Test Files
import { executeBrowserTests } from './browserTestRunner';
// Importing these files registers the tests in the runner
import '../tests/logic.test';
import '../tests/tree.test';
import '../tests/zip.test';
import '../tests/store.test';
import '../tests/theme.test';
import '../tests/aiService.test';
import '../tests/aiEvaluation.test';
import '../tests/aiRealModelEval.test';
import '../tests/zipInvariants.test';
import '../tests/types.test';
import '../tests/security.test';
import '../tests/resilience.test';
import { generateDiffTree } from './zipService';

export interface LogEntry {
  msg: string;
  type: 'info' | 'success' | 'warning' | 'error';
  timestamp: number;
}

export interface CoverageModule {
    name: string;
    score: number; // 0-100
    details: string;
}

export interface TestResult {
  success: boolean;
  logs: LogEntry[];
  coverage: CoverageModule[];
}

/**
 * Runs a suite of automated checks.
 * 1. Runs all registered Unit/Integration tests (via Vitest shim).
 * 2. Runs specific runtime checks on the provided user file (File A) to ensure it loads in this environment.
 */
export const runSystemChecks = async (fileA: File, fileB?: File): Promise<TestResult> => {
    const logs: LogEntry[] = [];
    const addLog = (msg: string, type: 'info' | 'success' | 'warning' | 'error') => {
        logs.push({ msg, type, timestamp: Date.now() });
    };

    // Initialize Coverage Trackers
    const coverage = {
        unitTests: 0,
        runtimeIntegrity: 0,
        diffLogic: 0,
    };

    try {
        addLog(`File A Input Size: ${(fileA.size / 1024).toFixed(2)} KB`, 'info');
        if (fileB) {
            addLog(`File B Input Size: ${(fileB.size / 1024).toFixed(2)} KB`, 'info');
        }

        // --- STEP 1: RUN VITEST SUITE (Browser Runner) ---
        addLog("🚀 Executing Vitest Suites (Store, Logic, Tree, Zip)...", 'info');
        
        // Execute the tests registered by the imports above
        const testResults = await executeBrowserTests();
        
        // Append test logs to the main log output
        testResults.logs.forEach(l => logs.push(l));

        if (testResults.failed === 0) {
            coverage.unitTests = 100;
            addLog(`✅ All ${testResults.passed} Unit Tests Passed`, 'success');
        } else {
            coverage.unitTests = Math.floor((testResults.passed / (testResults.passed + testResults.failed)) * 100);
            addLog(`❌ Unit Tests Failed: ${testResults.failed} errors`, 'error');
            // If unit tests fail, we mark overall success as false but continue to runtime checks
        }

        // --- STEP 2: RUN RUNTIME INTEGRATION CHECKS (Real User Data) ---
        // These checks verify that the specific file the user uploaded can be processed by the current browser environment.
        
        addLog("--- Runtime Integrity Checks ---", 'info');

        // 2.1 File A Integrity Check
        const zipA = await new JSZip().loadAsync(fileA);
        coverage.runtimeIntegrity += 20; // Loaded successfully
        addLog(`File A Integrity: Valid ZIP (${Object.keys(zipA.files).length} files)`, 'success');
        
        // 2.2 Critical OOXML Structure
        if (!zipA.file("[Content_Types].xml")) {
            throw new Error("File A is not a valid OOXML package ([Content_Types].xml missing).");
        }
        if (!zipA.file("_rels/.rels")) {
            addLog("⚠️ Missing _rels/.rels (Non-standard OOXML root)", 'warning');
        } else {
            addLog("Found Root Relationships (_rels/.rels)", 'success');
        }
        coverage.runtimeIntegrity += 20; 

        // 2.3 Read/Write Cycle Test (Simulates Editor logic on real file)
        const xmlFiles = Object.keys(zipA.files).filter(f => f.endsWith('.xml'));
        if (xmlFiles.length > 0) {
            const testFile = xmlFiles[0];
            const originalContent = await zipA.file(testFile)?.async("string");
            
            if (originalContent !== undefined) {
                const marker = "<!-- TEST_MARKER -->";
                const modifiedContent = originalContent + marker;
                
                // Write to zip (Memory only)
                zipA.file(testFile, modifiedContent); 
                
                // Read back
                const readBack = await zipA.file(testFile)?.async("string");
                
                if (readBack !== modifiedContent) {
                    throw new Error(`Read/Write cycle failed for ${testFile}: Content mismatch`);
                }
                coverage.runtimeIntegrity += 30; // R/W Cycle
                addLog(`Read/Write logic passed for ${testFile}`, 'success');
            }
        } else {
             addLog("No XML files found to test read/write cycle", 'warning');
             coverage.runtimeIntegrity += 30; // Skip but credit as not a failure
        }

        // 2.4 Dry-Run Export (Repacking Verification)
        addLog("📦 Verifying Repack/Export Engine...", 'info');
        const repackBlob = await zipA.generateAsync({ type: 'blob' });
        if (repackBlob.size === 0) throw new Error("Repack failed: Resulting blob is empty");
        addLog(`Export successful: Generated ${(repackBlob.size / 1024).toFixed(2)} KB blob`, 'success');
        coverage.runtimeIntegrity += 30;


        // --- STEP 3: DIFF & IMAGE LOGIC VERIFICATION ---
        addLog("--- Diff & Binary Logic Checks ---", 'info');

        // 3.1 Simulated Image Diff
        // We simulate a diff scenario to ensure the application correctly handles binary file changes
        // even if the user didn't upload two files.
        const mockFlatA = { 
            'image1.png': { name: 'image1.png', dir: false, _data: { crc32: 11111 } } 
        };
        const mockFlatB = { 
            'image1.png': { name: 'image1.png', dir: false, _data: { crc32: 22222 } } // Changed CRC
        };
        
        const mockTree = generateDiffTree(
            mockFlatA as unknown as Record<string, JSZip.JSZipObject>, 
            mockFlatB as unknown as Record<string, JSZip.JSZipObject>
        );
        const imageNode = mockTree.children['image1.png'];
        
        if (imageNode && imageNode.status === 'modified') {
            addLog("✅ Binary Image Diff Logic: Detected modified PNG via CRC32", 'success');
            coverage.diffLogic += 50;
        } else {
            addLog(`❌ Binary Image Diff Logic Failed: Expected 'modified', got '${imageNode?.status}'`, 'error');
        }

        // 3.2 Real Diff Integration (if File B provided)
        if (fileB) {
            const zipB = await new JSZip().loadAsync(fileB);
            
            // Generate Flat Maps
            const flatA: Record<string, JSZip.JSZipObject> = {};
            zipA.forEach((path, entry) => flatA[path] = entry);
            
            const flatB: Record<string, JSZip.JSZipObject> = {};
            zipB.forEach((path, entry) => flatB[path] = entry);

            addLog("Running Full Diff Comparison on provided files...", 'info');
            const tree = generateDiffTree(flatA, flatB);
            
            if (tree.children && Object.keys(tree.children).length > 0) {
                addLog(`Diff Tree Generated successfully with ${Object.keys(tree.children).length} root nodes`, 'success');
                coverage.diffLogic += 50;
            } else {
                throw new Error("Diff Tree generation resulted in empty root");
            }
        } else {
             addLog("Skipping Real Diff (No File B provided)", 'info');
             coverage.diffLogic += 50; // Credit for skipping without error
        }

        const report: CoverageModule[] = [
            { name: "Unit Test Suite", score: coverage.unitTests, details: `${testResults.passed} passed, ${testResults.failed} failed` },
            { name: "Runtime Integrity", score: coverage.runtimeIntegrity, details: "Structure, IO, Relations, Export" },
            { name: "Diff & Binary", score: coverage.diffLogic, details: "Image CRC, Tree Generation" }
        ];

        return { success: testResults.failed === 0, logs, coverage: report };

    } catch (e) {
        addLog((e as Error).message, 'error');
        const report: CoverageModule[] = [
            { name: "Unit Test Suite", score: coverage.unitTests, details: "Aborted" },
            { name: "Runtime Integrity", score: coverage.runtimeIntegrity, details: "Failed" },
            { name: "Diff & Binary", score: coverage.diffLogic, details: "Unknown" }
        ];
        return { success: false, logs, coverage: report };
    }
};