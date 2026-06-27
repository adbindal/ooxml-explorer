export interface ChromeLanguageModelSession {
  prompt: (input: string) => Promise<string>;
  promptStreaming?: (input: string) => AsyncIterable<string>;
  destroy: () => void;
}

export interface ChromeLanguageModel {
  availability: (options?: { languages?: string[] }) => Promise<'available' | 'downloadable' | 'downloading' | 'unavailable'>;
  create: (options?: { systemPrompt?: string; temperature?: number; topK?: number }) => Promise<ChromeLanguageModelSession>;
}

declare global {
  interface Window {
    LanguageModel?: ChromeLanguageModel;
  }
}
