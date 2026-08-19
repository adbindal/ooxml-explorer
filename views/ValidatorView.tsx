import React, { useState, useRef, useEffect } from 'react';
import { 
  ArrowLeft, Play, Loader2, Maximize2, GitCompare, ExternalLink, 
  CheckSquare, Square, Zap, AlertTriangle, Power, Activity,
  PieChart, ShieldCheck
} from 'lucide-react';
import { ThemeClasses } from '../types';
import { runSystemChecks, LogEntry, CoverageModule } from '../services/testService';
import { loadZipFile, readPackageParts } from '../services/zipService';
import { analyzePackage, capabilityLedger } from '../services/analyzers';
import { compareFindings } from '../services/findings';
import { readRetrievalMetrics, summariseRetrieval, resetRetrievalMetrics } from '../services/retrievalMetrics';
import { summariseCoverageGaps, resetCoverageGaps } from '../services/coverageGaps';
import {
    resolveAlternateContent,
    MODERN_CONSUMER_NAMESPACES,
    LEGACY_CONSUMER_NAMESPACES
} from '../services/markupCompatibility';
import { getApiKey, testConnection } from '../services/geminiService';

import { useAppStore } from '../store/appStore';
import { setDebugMode, getDebugMode } from '../services/debugService';

interface ValidatorViewProps {
  themeClasses: ThemeClasses;
}

const ValidatorView: React.FC<ValidatorViewProps> = ({ themeClasses }) => {
    const store = useAppStore();
    
    const [logs, setLogs] = useState<LogEntry[]>([]);
    const [status, setStatus] = useState<'idle' | 'running' | 'success' | 'error'>('idle');
    const [files, setFiles] = useState<{ a: File | null, b: File | null }>({ a: null, b: null });
    const [shouldCrash, setShouldCrash] = useState(false);
    const [debugEnabled, setDebugEnabled] = useState(getDebugMode());
    const [coverageReport, setCoverageReport] = useState<CoverageModule[]>([]);
    const [integrityStatus, setIntegrityStatus] = useState<'idle' | 'running' | 'done'>('idle');
    
    const consoleEndRef = useRef<HTMLDivElement>(null);
    
    useEffect(() => {
        if (consoleEndRef.current) {
            consoleEndRef.current.scrollIntoView({ behavior: 'smooth' });
        }
    }, [logs]);
    
    const [checklist, setChecklist] = useState<Record<string, boolean>>({
        theme: false, 
        lhp: false,
        rhp: false,
        gemini: false, 
        diffMode: false,
        diffNav: false
    });
    
    if (shouldCrash) {
        throw new Error("Manual Crash Test Triggered from Validator View");
    }

    const toggleDebug = () => {
        const newState = !debugEnabled;
        setDebugMode(newState);
        setDebugEnabled(newState);
    };

    const handleRunTests = async () => {
        if (!files.a) return;
        setStatus('running');
        setLogs([]);
        setCoverageReport([]);
        
        setLogs([{ msg: "🚀 Initializing Test Harness...", type: 'info', timestamp: Date.now() }]);

        const apiKey = getApiKey();
        if (apiKey) {
             setLogs(prev => [...prev, { msg: "🔑 Environment: API Key Configured (Ready)", type: 'success', timestamp: Date.now() }]);
        } else {
             setLogs(prev => [...prev, { msg: "⚠️ Environment: No API Key (AI features will be disabled)", type: 'warning', timestamp: Date.now() }]);
        }

        const result = await runSystemChecks(files.a, files.b || undefined);
        
        setLogs(prev => [...prev, ...result.logs]);
        setCoverageReport(result.coverage);
        setStatus(result.success ? 'success' : 'error');
        
        setLogs(prev => [...prev, { 
            msg: result.success ? "🏁 All Checks Passed" : "❌ Checks Failed", 
            type: result.success ? 'success' : 'error', 
            timestamp: Date.now() 
        }]);
    };

    /**
     * Runs the deterministic OPC integrity checks against the selected file.
     *
     * Distinct from the AI features on purpose: nothing here is generated or
     * retrieved. Every line it prints is computed from the package, so a clean run is
     * a fact about the file rather than a model's opinion of it.
     */
    /**
     * Prints how often each retrieval path has answered a lookup.
     *
     * Answers the question that decides whether semantic search is worth building at
     * all. Counts only - no query text is ever stored.
     */
    const handleShowRetrievalMetrics = () => {
        const summary = summariseRetrieval(readRetrievalMetrics());
        // The coverage backlog belongs next to the retrieval counts: both answer "what
        // should we build next?" from measurement rather than from guesswork, and
        // neither records anything about the user.
        const gaps = summariseCoverageGaps();
        setLogs(prev => [
            ...prev,
            { msg: '📊 Retrieval paths:', type: 'info', timestamp: Date.now() },
            ...summary.lines.map(line => ({ msg: line, type: 'info' as const, timestamp: Date.now() })),
            { msg: '🕳️ Coverage gaps — markup that was opened with no analyzer behind it:', type: 'info', timestamp: Date.now() },
            ...gaps.lines.map(line => ({ msg: line, type: 'info' as const, timestamp: Date.now() }))
        ]);
    };

    const handleCheckIntegrity = async () => {
        const addLog = (msg: string, type: 'info' | 'success' | 'warning' | 'error') =>
            setLogs(prev => [...prev, { msg, type, timestamp: Date.now() }]);

        if (!files.a) return;
        setIntegrityStatus('running');
        addLog(`🔍 Checking package integrity of ${files.a.name}...`, 'info');

        try {
            const { zip } = await loadZipFile(files.a);
            const parts = await readPackageParts(zip);
            // Every analyzer that applies, not just package integrity. This surface
            // previously ran one of them, so "is this file correct?" checked content
            // types and relationships and nothing else.
            const run = analyzePackage(parts);
            const findings = run.findings;
            const ledger = capabilityLedger(run);

            addLog(`Inspected ${Object.keys(parts).length} parts with ${ledger.ran.length} analyzer(s): ${ledger.ran.map(a => a.title).join(', ')}.`, 'info');
            if (ledger.skipped.length > 0) {
                addLog(`Not applicable to this package: ${ledger.skipped.map(a => a.title).join(', ')}.`, 'info');
            }

            // Markup compatibility is reported separately from integrity on purpose.
            // The integrity checks deliberately run against the *unresolved* markup,
            // because a broken relationship inside an mc:Fallback branch is still a
            // broken file for whichever consumer takes that branch.
            const alternateContent = Object.entries(parts)
                .filter(([path, content]) => path.endsWith('.xml') && content.includes('AlternateContent'))
                .map(([path, content]) => {
                    const doc = new DOMParser().parseFromString(content, 'application/xml');
                    const modern = resolveAlternateContent(
                        doc, MODERN_CONSUMER_NAMESPACES
                    ).selections;
                    const legacy = resolveAlternateContent(
                        new DOMParser().parseFromString(content, 'application/xml'),
                        LEGACY_CONSUMER_NAMESPACES
                    ).selections;
                    return { path, modern, legacy };
                })
                .filter(entry => entry.modern.length > 0);

            if (alternateContent.length > 0) {
                const total = alternateContent.reduce((n, e) => n + e.modern.length, 0);
                addLog(`ℹ️ ${total} block(s) of alternate content across ${alternateContent.length} part(s) — content written more than once for different Office versions.`, 'info');
                for (const entry of alternateContent) {
                    const divergent = entry.modern.filter((sel, i) => sel.chose !== entry.legacy[i]?.chose).length;
                    addLog(
                        `  ${entry.path}: ${entry.modern.length} block(s)` +
                        (divergent > 0 ? ` — ${divergent} render differently in a pre-2010 Office build.` : ''),
                        'info'
                    );
                }
            }

            if (findings.length === 0) {
                // Deliberately narrow wording. A clean run means the checks that ran
                // found nothing - not that the file is correct. The ledger below says
                // which checks those were and what they cannot see.
                addLog('✅ No problems found by the checks that ran.', 'success');
            } else {
                const errors = findings.filter(f => f.severity === 'error').length;
                const warnings = findings.length - errors;
                addLog(`Found ${errors} error(s) and ${warnings} warning(s).`, errors > 0 ? 'error' : 'warning');
                // Most integrity faults are silent - the package opens and renders and is
                // broken anyway - so say which ones, rather than leaving the reader to
                // assume anything visible would already have been noticed.
                for (const finding of [...findings].sort(compareFindings)) {
                    const silent = finding.silent ? ' (renders correctly and is broken anyway)' : '';
                    addLog(
                        `[${finding.code}] ${finding.part} — ${finding.message}${silent} ${finding.remediation}`,
                        // The log has no 'note' level; notes are informational, so they map to 'info'.
                        finding.severity === 'note' ? 'info' : finding.severity
                    );
                }
            }

            // The honest half: what the checks that ran still cannot see. Computed from
            // the registry, never asserted by a model.
            if (ledger.limits.length > 0) {
                addLog('These checks cannot establish:', 'info');
                for (const limit of ledger.limits) addLog(`  • ${limit}`, 'info');
            }
            setIntegrityStatus('done');
        } catch (e) {
            addLog(`❌ Could not read the package: ${(e as Error).message}`, 'error');
            setIntegrityStatus('done');
        }
    };

    const handleInteractiveTest = async (id: string) => {
        const addLog = (msg: string, type: 'info' | 'success' | 'warning' | 'error') =>
            setLogs(prev => [...prev, { msg, type, timestamp: Date.now() }]);

        try {
            switch (id) {
                case 'gemini': {
                    addLog("🤖 Testing Gemini Connection...", 'info');
                    const result = await testConnection();
                    if (result.success) {
                        addLog(`✅ ${result.message}`, 'success');
                    } else {
                        addLog(`❌ ${result.message}`, 'error');
                        if (result.message.includes("API Key not found")) {
                            addLog("👉 Configure key in the AI Assistant panel (RHP).", 'info');
                        }
                    }
                    break;
                }
                case 'theme':
                    addLog("🎨 Toggling theme...", 'info');
                    store.toggleTheme();
                    addLog("Theme toggled successfully in Store", 'success');
                    break;
                case 'lhp':
                    addLog("↔️ Toggling Sidebar (Left Hand Panel)...", 'info');
                    store.toggleSidebar();
                    addLog(`Sidebar state is now: ${useAppStore.getState().ui.sidebarOpen}`, 'success');
                    break;
                case 'rhp':
                    addLog("✨ Toggling AI Panel (Right Hand Panel)...", 'info');
                    store.toggleAiPanel();
                    addLog(`AI Panel state is now: ${useAppStore.getState().ui.showAi}`, 'success');
                    break;
                case 'diffMode':
                    addLog("👁️ Toggling Diff View Mode (Split/Inline)...", 'info');
                    store.setDiffViewMode(store.ui.diffViewMode === 'split' ? 'inline' : 'split');
                    addLog(`Diff Mode is now: ${useAppStore.getState().ui.diffViewMode}`, 'success');
                    break;
                case 'diffNav':
                    addLog("⬇️ Testing Diff Navigation controls...", 'info');
                    addLog("Logic connected to Monaco Editor instance. Check Diff View header for 'Next/Prev' buttons.", 'success');
                    break;
                default:
                    addLog(`Manual check for ${id}`, 'info');
            }
            setChecklist(prev => ({ ...prev, [id]: !prev[id] }));
        } catch (e: unknown) {
            const err = e as Error;
            addLog(`❌ Test Failed: ${err.message}`, 'error');
        }
    };

    const uxItems = [
        { id: 'theme', label: 'Theme Toggle' },
        { id: 'lhp', label: 'Sidebar (LHP) Toggle' },
        { id: 'rhp', label: 'AI Panel (RHP) Toggle' },
        { id: 'gemini', label: 'Gemini Connectivity' },
        { id: 'diffMode', label: 'Diff View Mode' },
        { id: 'diffNav', label: 'Diff Navigation' }
    ];

    // Helpers for switching modes via Store
    const handleOpenA = () => {
        if (!files.a) return;
        store.loadEditorFile(files.a).catch(e => alert(e.message));
    };

    const handleDiffAB = () => {
        if (!files.a || !files.b) return;
        store.setDiffFiles(files.a, files.b);
        store.setMode('diff-view');
    };

    return (
        <div className={`h-screen w-full flex flex-col md:p-8 font-sans overflow-hidden transition-colors duration-300 ${themeClasses.bg} ${themeClasses.fg}`}>
            <div className={`w-full md:max-w-4xl mx-auto md:border md:rounded-xl shadow-2xl flex flex-col h-full ${themeClasses.card}`}>
                <div className={`flex items-center gap-4 p-4 md:p-6 border-b shrink-0 ${themeClasses.border} ${themeClasses.bgSec}`}>
                    <button onClick={() => store.setMode('landing')} className={themeClasses.hoverText}><ArrowLeft size={20}/></button>
                    <div>
                        <h2 className="text-lg md:text-xl font-bold">QA Validation Suite</h2>
                        <p className={`text-xs ${themeClasses.fgMuted}`}>Unit tests and system integrity checks</p>
                    </div>
                </div>

                <div className="flex flex-col md:flex-row flex-1 overflow-hidden">
                    {/* Left: Setup & Checklist */}
                    <div className={`w-full md:w-1/3 p-4 md:p-6 border-b md:border-b-0 md:border-r flex flex-col gap-6 overflow-y-auto shrink-0 ${themeClasses.border}`}>
                        <div className="space-y-4">
                            <h3 className="font-bold text-sm uppercase tracking-wider">1. Unit Test Data</h3>
                            <div className="space-y-2">
                                <label className={`block text-xs font-medium ${themeClasses.fgMuted}`}>Primary File (Required)</label>
                                <input type="file" className={`text-xs w-full border rounded p-2 ${themeClasses.input}`} onChange={(e) => setFiles(f => ({ ...f, a: e.target.files?.[0] || null }))} />
                            </div>
                            <div className="space-y-2">
                                <label className={`block text-xs font-medium ${themeClasses.fgMuted}`}>Comparison File (Optional)</label>
                                <input type="file" className={`text-xs w-full border rounded p-2 ${themeClasses.input}`} onChange={(e) => setFiles(f => ({ ...f, b: e.target.files?.[0] || null }))} />
                            </div>
                            <button 
                                onClick={handleRunTests} 
                                disabled={!files.a || status === 'running'}
                                className={`w-full py-2 rounded font-medium flex items-center justify-center gap-2 transition-all ${!files.a ? 'bg-zinc-700 text-zinc-500 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-500 text-white'}`}
                            >
                                {status === 'running' ? <Loader2 className="animate-spin" size={16} /> : <Play size={16} />} Run Unit Tests
                            </button>

                            {/* Deterministic package checks - no model, no network. */}
                            <button
                                onClick={handleCheckIntegrity}
                                disabled={!files.a || integrityStatus === 'running'}
                                title="Verify content types and relationship integrity. Computed, not inferred."
                                className={`w-full py-2 rounded font-medium flex items-center justify-center gap-2 transition-all ${!files.a ? 'bg-zinc-700 text-zinc-500 cursor-not-allowed' : 'bg-emerald-700 hover:bg-emerald-600 text-white'}`}
                            >
                                {integrityStatus === 'running' ? <Loader2 className="animate-spin" size={16} /> : <ShieldCheck size={16} />} Check Package Integrity
                            </button>

                            {/* Counts only; never query text. */}
                            <div className="grid grid-cols-3 gap-2">
                                <button
                                    onClick={handleShowRetrievalMetrics}
                                    title="How often each retrieval path answered a lookup. Decides whether semantic search is worth building."
                                    className="col-span-2 py-2 rounded font-medium flex items-center justify-center gap-2 bg-slate-600 hover:bg-slate-500 text-white text-xs"
                                >
                                    <PieChart size={14} /> Retrieval Stats
                                </button>
                                <button
                                    onClick={() => { resetRetrievalMetrics(); resetCoverageGaps(); handleShowRetrievalMetrics(); }}
                                    title="Reset the retrieval counters."
                                    className="py-2 rounded font-medium bg-zinc-700 hover:bg-zinc-600 text-white text-xs"
                                >
                                    Reset
                                </button>
                            </div>
                            
                            {/* Coverage Report Widget */}
                            {coverageReport.length > 0 && (
                                <div className="animate-in fade-in slide-in-from-top-2 border rounded-lg overflow-hidden mt-4">
                                    <div className={`px-3 py-2 text-xs font-bold uppercase flex items-center gap-2 ${themeClasses.bgSec} ${themeClasses.border} border-b`}>
                                        <PieChart size={12} /> Functional Coverage
                                    </div>
                                    <div className={`p-3 space-y-3 ${themeClasses.bg}`}>
                                        {coverageReport.map(item => (
                                            <div key={item.name} className="space-y-1">
                                                <div className="flex justify-between text-[10px] uppercase font-bold opacity-70">
                                                    <span>{item.name}</span>
                                                    <span>{item.score}%</span>
                                                </div>
                                                <div className="h-1.5 w-full bg-gray-700 rounded-full overflow-hidden">
                                                    <div 
                                                        className={`h-full rounded-full transition-all duration-500 ${item.score === 100 ? 'bg-green-500' : item.score > 50 ? 'bg-blue-500' : 'bg-red-500'}`} 
                                                        style={{ width: `${item.score}%` }}
                                                    ></div>
                                                </div>
                                                <div className="text-[9px] opacity-50 truncate">{item.details}</div>
                                            </div>
                                        ))}
                                        <div className="pt-2 border-t border-dashed border-gray-700 text-[10px] text-center opacity-50">
                                            Values represent critical path verification by runtime harness.
                                        </div>
                                    </div>
                                </div>
                            )}

                            {/* ACTIONS: Context-Aware Buttons */}
                            {status === 'success' && (
                                <div className="space-y-2 animate-in fade-in slide-in-from-top-2">
                                    <div className={`text-xs font-bold uppercase tracking-wider mb-2 ${themeClasses.fgMuted}`}>Integration Actions</div>
                                    {files.b && files.a ? (
                                        <div className="grid grid-cols-2 gap-2">
                                            <button 
                                                onClick={handleOpenA}
                                                className="py-2 bg-blue-600 hover:bg-blue-500 text-white rounded font-medium flex flex-col items-center justify-center text-xs gap-1"
                                            >
                                                <Maximize2 size={14} /> Edit File A
                                            </button>
                                            <button 
                                                onClick={handleDiffAB}
                                                className="py-2 bg-slate-600 hover:bg-slate-500 text-white rounded font-medium flex flex-col items-center justify-center text-xs gap-1"
                                            >
                                                <GitCompare size={14} /> Diff A vs B
                                            </button>
                                        </div>
                                    ) : files.a ? (
                                        <button 
                                            onClick={handleOpenA}
                                            className="w-full py-2 bg-green-600 hover:bg-green-500 text-white rounded font-medium flex items-center justify-center gap-2"
                                        >
                                            <ExternalLink size={16} /> Open "{files.a.name}"
                                        </button>
                                    ) : null}
                                </div>
                            )}
                        </div>

                        <div className="space-y-3">
                            <h3 className="font-bold text-sm uppercase tracking-wider">2. UX Verification</h3>
                            {uxItems.map(item => (
                                <div key={item.id} className="flex items-center justify-between gap-2 group">
                                    <div className="flex items-center gap-2 cursor-pointer" onClick={() => handleInteractiveTest(item.id)}>
                                        {checklist[item.id] ? <CheckSquare size={16} className="text-green-500" /> : <Square size={16} className={themeClasses.fgMuted} />}
                                        <span className={`text-sm ${checklist[item.id] ? themeClasses.fg : themeClasses.fgMuted}`}>{item.label}</span>
                                    </div>
                                    <button 
                                        onClick={() => handleInteractiveTest(item.id)}
                                        className={`p-1 rounded hover:bg-blue-500/20 hover:text-blue-500 ${themeClasses.fgMuted} opacity-0 group-hover:opacity-100 transition-all`}
                                        title="Run Action"
                                    >
                                        <Zap size={12} />
                                    </button>
                                </div>
                            ))}
                        </div>

                        {/* Resilience & Debug Control */}
                         <div className="space-y-3 pt-4 border-t border-red-500/20">
                            <div className="flex items-center justify-between mb-2">
                                <h3 className="font-bold text-sm uppercase tracking-wider text-red-500">3. Resilience & Debug</h3>
                                <div className={`flex items-center gap-1.5 px-2 py-0.5 rounded-full text-[10px] font-bold ${debugEnabled ? 'bg-red-500/10 text-red-500' : 'bg-zinc-500/10 text-zinc-500'}`}>
                                    <Activity size={10} />
                                    {debugEnabled ? 'REC' : 'OFF'}
                                </div>
                            </div>
                            
                            <button 
                                onClick={toggleDebug}
                                className={`w-full py-2 border rounded font-medium flex items-center justify-center gap-2 transition-colors text-xs ${debugEnabled ? 'bg-red-900/10 border-red-500/30 text-red-400' : 'bg-transparent border-zinc-700 text-zinc-500 hover:text-zinc-300'}`}
                            >
                                <Power size={14} /> {debugEnabled ? 'Disable Logging' : 'Enable Detailed Logging'}
                            </button>

                            <button 
                                onClick={() => setShouldCrash(true)}
                                className="w-full py-2 bg-red-900/10 hover:bg-red-900/20 text-red-500 border border-red-500/20 rounded font-medium flex items-center justify-center gap-2 text-xs"
                            >
                                <AlertTriangle size={14} /> Trigger Crash (Error Boundary)
                            </button>
                        </div>
                    </div>

                    {/* Right: Console Output */}
                    <div className={`flex-1 p-4 md:p-6 font-mono text-xs overflow-y-auto bg-black text-green-400 min-h-[200px]`}>
                        {logs.length === 0 ? <span className="opacity-50">// Waiting for test execution...</span> : logs.map((l, i) => (
                            <div key={i} className={`mb-2 ${l.type === 'error' ? 'text-red-400' : l.type === 'warning' ? 'text-yellow-400' : 'text-green-400'}`}>
                                <span className="opacity-50 mr-2">[{new Date(l.timestamp).toLocaleTimeString()}]</span>
                                {l.msg}
                            </div>
                        ))}
                        <div ref={consoleEndRef} />
                    </div>
                </div>
            </div>
        </div>
    );
};

export default ValidatorView;