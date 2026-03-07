
interface LogEntry {
    timestamp: string;
    level: 'info' | 'warn' | 'error' | 'debug';
    message: string;
    data?: any[];
}

const MAX_LOGS = 200;
const logHistory: LogEntry[] = [];
const STORAGE_KEY = 'ooxml_debug_mode';

// Default to FALSE (off) unless explicitly enabled in storage
let isCapturing = typeof localStorage !== 'undefined' && localStorage.getItem(STORAGE_KEY) === 'true';

const formatArgs = (args: any[]): string => {
    return args.map(arg => {
        try {
            if (arg instanceof Error) return `${arg.name}: ${arg.message}\n${arg.stack}`;
            if (typeof arg === 'object') return JSON.stringify(arg);
            return String(arg);
        } catch (e) {
            return '[Circular/Unserializable]';
        }
    }).join(' ');
};

const addLog = (level: LogEntry['level'], args: any[]) => {
    // Optimization: Do not incur overhead if capturing is disabled
    if (!isCapturing) return;

    const message = formatArgs(args);
    const entry: LogEntry = {
        timestamp: new Date().toISOString(),
        level,
        message,
        data: args // Keep raw refs if needed for immediate inspection, though JSON export uses message
    };
    
    logHistory.push(entry);
    if (logHistory.length > MAX_LOGS) {
        logHistory.shift();
    }
};

export const setDebugMode = (enabled: boolean) => {
    isCapturing = enabled;
    localStorage.setItem(STORAGE_KEY, String(enabled));
    // We use a direct console call here, but since we patched it, we rely on the patch logic.
    // However, if we just turned it OFF, the patch logic suppresses it. 
    // If we turned it ON, the patch logic allows it.
    console.log(`[DebugService] Detailed Logging set to: ${enabled ? 'ON' : 'OFF'}`);
    
    if (enabled && logHistory.length === 0) {
        addLog('info', ['[DebugService] Logging Enabled']);
    }
};

export const getDebugMode = () => isCapturing;

export const initDebugService = () => {
    // 1. Capture Environment Info
    const envInfo = {
        userAgent: navigator.userAgent,
        screen: `${window.screen.width}x${window.screen.height}`,
        url: window.location.href,
        loggingEnabled: isCapturing
    };
    
    // Capture original methods before overwriting
    const originalLog = console.log;
    const originalWarn = console.warn;
    const originalError = console.error;
    const originalInfo = console.info;
    const originalDebug = console.debug;

    // Initial Log: Only show if enabled to keep console clean
    if (isCapturing) {
        originalInfo("Initializing Debug Service...", envInfo);
        addLog('info', ["Initializing Debug Service...", envInfo]);
    }

    // 2. Patch Console methods
    // Policy: 
    // - Error/Warn: Always print to console (critical), capture if enabled.
    // - Log/Info/Debug: Only print AND capture if enabled (noise reduction).

    console.log = (...args) => { 
        if (isCapturing) {
            addLog('info', args); 
            originalLog.apply(console, args); 
        }
    };

    console.warn = (...args) => { 
        // Warnings are important, always show in console
        originalWarn.apply(console, args);
        if (isCapturing) addLog('warn', args); 
    };

    console.error = (...args) => { 
        // Errors are critical, always show in console
        originalError.apply(console, args);
        if (isCapturing) addLog('error', args); 
    };

    console.info = (...args) => { 
        if (isCapturing) {
            addLog('info', args); 
            originalInfo.apply(console, args); 
        }
    };

    console.debug = (...args) => { 
        if (isCapturing) {
            addLog('debug', args); 
            originalDebug.apply(console, args); 
        }
    };

    // 3. Global Error Handlers
    window.addEventListener('error', (event) => {
        const msg = event.message ? String(event.message) : 'Unknown Error';
        
        // Deep inspection for Script Error
        const debugInfo = {
            message: msg,
            filename: event.filename,
            lineno: event.lineno,
            colno: event.colno,
            error: event.error ? {
                name: event.error.name,
                message: event.error.message,
                stack: event.error.stack
            } : 'null',
            type: event.type,
            target: event.target ? (event.target as HTMLElement).tagName : 'unknown'
        };

        // Always capture errors if we can, or just rely on console.error patch?
        // Global handler errors usually print to console automatically by browser.
        // We just ensure they get into our log history if capturing.
        addLog('error', [`[Global Error]`, debugInfo]);

        // Ignore ResizeObserver errors
        if (msg.includes('ResizeObserver loop') || msg.includes('ResizeObserver')) {
            event.stopImmediatePropagation();
            event.preventDefault();
            return;
        }

        // Suppress generic Script Errors but keep them in log
        if (msg.toLowerCase().includes('script error')) {
            // We use originalWarn to ensure this helps debugging if needed, but doesn't crash app
            if (isCapturing) {
                originalWarn.apply(console, ['Suppressed generic Script Error (likely Cross-Origin/CDN resource issue). Debug Info:', JSON.stringify(debugInfo)]);
            }
            event.stopImmediatePropagation();
            event.preventDefault();
            return;
        }
    });

    window.addEventListener('unhandledrejection', (event) => {
        const reason = event.reason;
        const msg = reason instanceof Error ? reason.message : String(reason);

        addLog('error', [`[Unhandled Rejection] ${msg}`, reason]);

        if (msg.includes('ResizeObserver') || msg.includes('ResizeObserver loop')) {
            event.preventDefault();
            event.stopImmediatePropagation();
            return;
        }
        
        if (msg.toLowerCase().includes('script error')) {
             if (isCapturing) {
                originalWarn.apply(console, ['Suppressed generic Script Error in promise.']);
            }
            event.preventDefault();
            event.stopImmediatePropagation();
            return;
        }
    });
};

export const getLogDump = () => {
    return {
        environment: {
            userAgent: navigator.userAgent,
            timestamp: new Date().toISOString(),
            url: window.location.href,
            loggingEnabled: isCapturing
        },
        logs: logHistory
    };
};

export const getLogString = () => {
    const dump = getLogDump();
    return JSON.stringify(dump, null, 2);
};
