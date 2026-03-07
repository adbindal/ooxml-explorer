
export interface AssertionResult {
    pass: boolean;
    message: string;
}

export interface TestDefinition {
    name: string;
    fn: () => void | Promise<void>;
    suiteName: string;
}

const registeredTests: TestDefinition[] = [];
let currentSuite = 'Global';

export const describe = (name: string, fn: () => void) => {
    const prevSuite = currentSuite;
    currentSuite = name;
    try {
        fn();
    } finally {
        currentSuite = prevSuite;
    }
};

export const it = (name: string, fn: () => void | Promise<void>) => {
    registeredTests.push({
        name,
        fn,
        suiteName: currentSuite
    });
};

export const beforeEach = (fn: () => void) => {
    // Basic support: run immediately for setup in this simplistic runner
    // In a full runner, this would run before every test in the suite.
    // For now, we assume tests are largely independent or setup is safe to run once per suite definition.
    try { fn(); } catch(e) { console.error("beforeEach failed", e); }
};

// Mocking 'vi' object
export const vi = {
    fn: (impl?: Function) => {
        const mockFn = (...args: any[]) => {
            mockFn.mock.calls.push(args);
            return impl ? impl(...args) : undefined;
        };
        mockFn.mock = { calls: [] as any[][] };
        return mockFn;
    },
    mock: (moduleName: string, factory: any) => {
        console.warn(`[BrowserRunner] vi.mock('${moduleName}') ignored. Tests will run against REAL implementations.`);
    },
    clearAllMocks: () => {
        // No-op in browser runner context
    }
};

// Assertion Logic
export const expect = (actual: any) => ({
    toBe: (expected: any) => {
        if (actual !== expected) throw new Error(`Expected ${expected}, received ${actual}`);
    },
    toEqual: (expected: any) => {
        const strActual = JSON.stringify(actual);
        const strExpected = JSON.stringify(expected);
        if (strActual !== strExpected) throw new Error(`Expected ${strExpected}, received ${strActual}`);
    },
    toContain: (expected: any) => {
        if (Array.isArray(actual) || typeof actual === 'string') {
            if (!actual.includes(expected)) throw new Error(`Expected collection to contain ${expected}`);
        } else {
            throw new Error(`expect(actual).toContain() expected array or string`);
        }
    },
    toHaveLength: (expected: number) => {
        if (actual.length !== expected) throw new Error(`Expected length ${expected}, received ${actual.length}`);
    },
    not: {
        toContain: (expected: any) => {
            if (Array.isArray(actual) || typeof actual === 'string') {
                if (actual.includes(expected)) throw new Error(`Expected collection NOT to contain ${expected}`);
            }
        },
        toBe: (expected: any) => {
            if (actual === expected) throw new Error(`Expected NOT to be ${expected}`);
        }
    },
    toBeDefined: () => {
        if (actual === undefined) throw new Error(`Expected value to be defined`);
    },
    toBeNull: () => {
        if (actual !== null) throw new Error(`Expected null, received ${actual}`);
    },
    toBeTruthy: () => {
        if (!actual) throw new Error(`Expected truthy, received ${actual}`);
    },
    rejects: {
        toThrow: async (msg?: string) => {
            try {
                await actual;
                throw new Error("Expected promise to reject, but it resolved");
            } catch (e: any) {
                if (msg && !e.message.includes(msg)) {
                    throw new Error(`Expected error containing "${msg}", got "${e.message}"`);
                }
            }
        }
    }
});

/**
 * Executes all registered tests and returns logs.
 */
export const executeBrowserTests = async (): Promise<{ logs: any[], passed: number, failed: number }> => {
    const logs: any[] = [];
    let passed = 0;
    let failed = 0;

    console.log(`[BrowserRunner] Starting execution of ${registeredTests.length} tests...`);

    for (const test of registeredTests) {
        const startTime = Date.now();
        try {
            await test.fn();
            passed++;
            logs.push({ 
                msg: `✅ [${test.suiteName}] ${test.name}`, 
                type: 'success', 
                timestamp: startTime 
            });
        } catch (e: any) {
            failed++;
            logs.push({ 
                msg: `❌ [${test.suiteName}] ${test.name}: ${e.message}`, 
                type: 'error', 
                timestamp: startTime 
            });
            console.error(e);
        }
    }

    return { logs, passed, failed };
};
