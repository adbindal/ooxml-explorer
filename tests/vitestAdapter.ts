import { describe, it, expect, vi, beforeEach, afterEach, beforeAll, afterAll } from 'vitest';

// Re-export everything needed by the tests
export { describe, it, expect, vi, beforeEach, afterEach, beforeAll, afterAll };

// Mock executeBrowserTests to do nothing or throw in Node environment
// since we are running tests directly via Vitest
export const executeBrowserTests = async () => {
    return { logs: [], passed: 0, failed: 0 };
};
