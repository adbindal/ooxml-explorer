import { describe, it, expect, beforeEach, vi } from '../services/browserTestRunner';
import { getActiveAIProvider, explainElement, getAiClient } from '../services/aiService';
import { useAppStore } from '../store/appStore';

describe('AI Service Layer 1', () => {
  // Test suite setup

  beforeEach(() => {
    // Reset store state
    useAppStore.setState({
      ui: {
        ...useAppStore.getState().ui,
        aiProvider: 'gemini-cloud'
      }
    });

    // Reset window.LanguageModel
    window.LanguageModel = undefined;
    
    // Clear all mocks
    const ai = getAiClient('mock-api-key');
    if (ai && ai.models) {
      vi.spyOn(ai.models, 'generateContent').mockRestore?.();
    }
  });

  describe('getActiveAIProvider', () => {
    it('returns cloud when preferred provider is cloud', async () => {
      useAppStore.setState(state => ({
        ui: { ...state.ui, aiProvider: 'gemini-cloud' }
      }));
      const provider = await getActiveAIProvider();
      expect(provider).toBe('cloud');
    });

    it('returns cloud when preferred is local but LanguageModel is not in window', async () => {
      useAppStore.setState(state => ({
        ui: { ...state.ui, aiProvider: 'chrome-local' }
      }));
      window.LanguageModel = undefined;
      const provider = await getActiveAIProvider();
      expect(provider).toBe('cloud');
    });

    it('returns cloud when preferred is local but LanguageModel availability is not available', async () => {
      useAppStore.setState(state => ({
        ui: { ...state.ui, aiProvider: 'chrome-local' }
      }));
      window.LanguageModel = {
        availability: async () => 'downloadable',
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        create: async () => ({} as any)
      };
      const provider = await getActiveAIProvider();
      expect(provider).toBe('cloud');
    });

    it('returns local when preferred is local and LanguageModel availability is available', async () => {
      useAppStore.setState(state => ({
        ui: { ...state.ui, aiProvider: 'chrome-local' }
      }));
      window.LanguageModel = {
        availability: async () => 'available',
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        create: async () => ({} as any)
      };
      const provider = await getActiveAIProvider();
      expect(provider).toBe('local');
    });
  });

  describe('explainElement', () => {
    it('routes to local AI when provider is local', async () => {
      useAppStore.setState(state => ({
        ui: { ...state.ui, aiProvider: 'chrome-local' }
      }));

      const mockPrompt = vi.fn(async () => 'Mocked local tag explanation');
      const mockDestroy = vi.fn();
      
      window.LanguageModel = {
        availability: async () => 'available',
        create: async (options?: { systemPrompt?: string }) => {
          expect(options?.systemPrompt).toContain('expert');
          return {
            prompt: mockPrompt,
            destroy: mockDestroy
          };
        }
      };

      const explanation = await explainElement('tcW', '<w:tcW w:w="120" />', 'docx');
      expect(explanation).toBe('Mocked local tag explanation');
      expect(mockPrompt.mock.calls.length).toBe(1);
      expect(mockDestroy.mock.calls.length).toBe(1);
    });

    it('routes to cloud AI when provider is cloud', async () => {
      useAppStore.setState(state => ({
        ui: { ...state.ui, aiProvider: 'gemini-cloud' }
      }));

      // Mock the cloud API key
      localStorage.setItem('ooxml_explorer_api_key', 'mock-api-key');

      const ai = getAiClient('mock-api-key');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      vi.spyOn(ai.models, 'generateContent').mockImplementation(async (args: any) => {
        expect(args.model).toBe('gemini-2.5-flash');
        expect(args.contents).toContain('tcW');
        return {
          text: 'Mocked cloud tag explanation'
        };
      });

      const explanation = await explainElement('tcW', '<w:tcW w:w="120" />', 'docx');
      expect(explanation).toBe('Mocked cloud tag explanation');
      
      localStorage.removeItem('ooxml_explorer_api_key');
    });
  });
});
