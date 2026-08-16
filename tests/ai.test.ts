import { describe, it, expect, beforeEach } from '../services/browserTestRunner';
import { getApiKey, setApiKey, clearApiKey, analyzeFile, analyzeDiff } from '../services/geminiService';
import { useAppStore } from '../store/appStore';

describe('Gemini AI Service', () => {
    beforeEach(() => {
        clearApiKey();
        // Mock localStorage if needed, but browserTestRunner might not handle it.
        // In a real browser environment, localStorage exists.

        // DLP Mode defaults to on; these tests exercise the cloud/API-key path
        // directly, so opt out of DLP for them explicitly.
        useAppStore.setState(state => ({
            ui: { ...state.ui, aiProvider: 'gemini-cloud', dlpMode: false }
        }));
        window.LanguageModel = undefined;
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

    describe('DLP Mode Security Shield', () => {
        it('analyzeFile is blocked by DLP mode instead of silently falling back to cloud', async () => {
            useAppStore.setState(state => ({ ui: { ...state.ui, dlpMode: true } }));
            setApiKey('test-key'); // Even with a valid key, cloud must not be reachable under DLP mode.

            const mockFiles = [{ fileName: 'test.xml', content: '<root/>' }];
            await expect(analyzeFile(mockFiles, 'explain')).rejects.toThrow('DLP_BLOCK');
        });

        it('analyzeDiff is blocked by DLP mode instead of silently falling back to cloud', async () => {
            useAppStore.setState(state => ({ ui: { ...state.ui, dlpMode: true } }));
            setApiKey('test-key');

            const mockFiles = [{ fileName: 'test.xml', original: '<root/>', modified: '<root/>' }];
            await expect(analyzeDiff(mockFiles, 'summary')).rejects.toThrow('DLP_BLOCK');
        });
    });
});
