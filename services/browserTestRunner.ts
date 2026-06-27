export interface AssertionResult {
    pass: boolean;
    message: string;
}

export interface TestDefinition {
    name: string;
    fn: () => void | Promise<void>;
    suiteName: string;
}

export interface TestLogEntry {
    msg: string;
    type: 'info' | 'success' | 'warning' | 'error';
    timestamp: number;
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

const beforeEachCallbacks: (() => void | Promise<void>)[] = [];

export const beforeEach = (fn: () => void | Promise<void>) => {
    beforeEachCallbacks.push(fn);
};

// Mocking 'vi' object
export const vi = {
    fn: <T extends (...args: unknown[]) => unknown>(impl?: T) => {
        const mockFn = (...args: unknown[]) => {
            mockFn.mock.calls.push(args);
            return impl ? impl(...args) : undefined;
        };
        mockFn.mock = { calls: [] as unknown[][] };
        return mockFn;
    },
    spyOn: <T, K extends keyof T>(obj: T, method: K) => {
        const original = obj[method];
        const mockFn = (...args: unknown[]) => {
            mockFn.mock.calls.push(args);
            return typeof original === 'function' ? original.apply(obj, args) : undefined;
        };
        mockFn.mock = { calls: [] as unknown[][] };
        mockFn.mockRestore = () => {
            obj[method] = original;
        };
        mockFn.mockImplementation = (impl: (...args: unknown[]) => unknown) => {
            // @ts-expect-error - bypass index signature assignment
            obj[method] = impl;
            return mockFn;
        };
        // @ts-expect-error - bypass index signature assignment for mock
        obj[method] = mockFn;
        return mockFn;
    },
    mock: (moduleName: string) => {
        console.warn(`[BrowserRunner] vi.mock('${moduleName}') ignored. Tests will run against REAL implementations.`);
    },
    clearAllMocks: () => {
        // No-op in browser runner context
    }
};

// Assertion Logic
export const expect = (actual: unknown) => ({
    toBe: (expected: unknown) => {
        if (actual !== expected) throw new Error(`Expected ${expected}, received ${actual}`);
    },
    toEqual: (expected: unknown) => {
        const strActual = JSON.stringify(actual);
        const strExpected = JSON.stringify(expected);
        if (strActual !== strExpected) throw new Error(`Expected ${strExpected}, received ${strActual}`);
    },
    toContain: (expected: unknown) => {
        if (Array.isArray(actual) || typeof actual === 'string') {
            const arr = actual as unknown[];
            if (!arr.includes(expected)) throw new Error(`Expected collection to contain ${expected}`);
        } else {
            throw new Error(`expect(actual).toContain() expected array or string`);
        }
    },
    toHaveLength: (expected: number) => {
        const arr = actual as { length: number };
        if (arr.length !== expected) throw new Error(`Expected length ${expected}, received ${arr.length}`);
    },
    not: {
        toContain: (expected: unknown) => {
            if (Array.isArray(actual) || typeof actual === 'string') {
                const arr = actual as unknown[];
                if (arr.includes(expected)) throw new Error(`Expected collection NOT to contain ${expected}`);
            }
        },
        toBe: (expected: unknown) => {
            if (actual === expected) throw new Error(`Expected NOT to be ${expected}`);
        }
    },
    toBeDefined: () => {
        if (actual === undefined) throw new Error(`Expected value to be defined`);
    },
    toBeUndefined: () => {
        if (actual !== undefined) throw new Error(`Expected undefined, received ${actual}`);
    },
    toBeGreaterThan: (expected: number) => {
        if (typeof actual !== 'number' || actual <= expected) {
            throw new Error(`Expected greater than ${expected}, received ${actual}`);
        }
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
                await (actual as Promise<unknown>);
                throw new Error("Expected promise to reject, but it resolved");
            } catch (e: unknown) {
                const err = e as Error;
                if (msg && !err.message.includes(msg)) {
                    throw new Error(`Expected error containing "${msg}", got "${err.message}"`);
                }
            }
        }
    }
});

/**
 * Executes all registered tests and returns logs.
 */
export const executeBrowserTests = async (): Promise<{ logs: TestLogEntry[], passed: number, failed: number }> => {
    const logs: TestLogEntry[] = [];
    let passed = 0;
    let failed = 0;

    console.log(`[BrowserRunner] Starting execution of ${registeredTests.length} tests...`);

    for (const test of registeredTests) {
        // Run all registered beforeEach callbacks
        for (const cb of beforeEachCallbacks) {
            try {
                await cb();
            } catch (e) {
                console.error("[BrowserRunner] beforeEach failed:", e);
            }
        }

        const startTime = Date.now();
        try {
            await test.fn();
            passed++;
            logs.push({ 
                msg: `✅ [${test.suiteName}] ${test.name}`, 
                type: 'success', 
                timestamp: startTime 
            });
        } catch (e: unknown) {
            const err = e as Error;
            failed++;
            logs.push({ 
                msg: `❌ [${test.suiteName}] ${test.name}: ${err.message}`, 
                type: 'error', 
                timestamp: startTime 
            });
            console.error(err);
        }
    }

    return { logs, passed, failed };
};
