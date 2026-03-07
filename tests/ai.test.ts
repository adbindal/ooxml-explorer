import { describe, it, expect, vi, beforeEach } from '../services/browserTestRunner';
import { getApiKey, setApiKey, clearApiKey, analyzeFile, analyzeDiff } from '../services/geminiService';

describe('Gemini AI Service', () => {
    beforeEach(() => {
        clearApiKey();
        // Mock localStorage if needed, but browserTestRunner might not handle it.
        // In a real browser environment, localStorage exists.
    });

    it('manages API key in storage', () => {
        expect(getApiKey()).toBeUndefined();
        setApiKey('test-key');
        expect(getApiKey()).toBe('test-key');
        clearApiKey();
        expect(getApiKey()).toBeUndefined();
    });

    it('analyzeFile throws error if API key is missing', async () => {
        const mockFiles = [{ fileName: 'test.xml', content: '<root/>' }];
        await expect(analyzeFile(mockFiles, 'explain')).rejects.toThrow('API_KEY_MISSING');
    });

    it('analyzeDiff throws error if API key is missing', async () => {
        const mockFiles = [{ fileName: 'test.xml', original: '<root/>', modified: '<root/>' }];
        await expect(analyzeDiff(mockFiles, 'summary')).rejects.toThrow('API_KEY_MISSING');
    });
});
