import { useAppStore } from "../store/appStore";

export type AIProvider = 'chrome-local' | 'gemini-cloud';

const isLocalAiAvailable = async (): Promise<boolean> => {
  if (typeof window === 'undefined' || !window.LanguageModel) return false;
  try {
    const availability = await window.LanguageModel.availability();
    return availability === 'available';
  } catch (e) {
    console.warn("[AI Provider] Error checking local AI availability:", e);
    return false;
  }
};

/**
 * Single source of truth for choosing between Chrome's built-in local model
 * and the Gemini Cloud API. Every AI action in the app (whole-file explain,
 * diff explain, and the selected-tag explainer) must route through this so
 * DLP Mode is enforced consistently everywhere, not just in some call sites.
 *
 * - DLP Mode (default on): cloud is never allowed. If local AI isn't
 *   available, the request fails with a DLP_BLOCK error instead of silently
 *   falling back to the cloud.
 * - Otherwise: honors the user's preferred provider, falling back to cloud
 *   if local AI was requested but isn't available/ready.
 */
export const getActiveAIProvider = async (): Promise<AIProvider> => {
  const { aiProvider: preferredProvider, dlpMode } = useAppStore.getState().ui;

  if (dlpMode) {
    if (await isLocalAiAvailable()) {
      return 'chrome-local';
    }
    throw new Error("DLP_BLOCK: Cloud AI is disabled under DLP mode, and Local AI is unavailable.");
  }

  if (preferredProvider === 'chrome-local' && await isLocalAiAvailable()) {
    return 'chrome-local';
  }

  return 'gemini-cloud';
};
