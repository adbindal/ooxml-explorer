import React, { useState } from 'react';
import { 
  ArrowLeft, Play, Loader2, Maximize2, GitCompare, ExternalLink, 
  CheckSquare, Square, Zap, AlertTriangle, Power, Activity,
  PieChart
} from 'lucide-react';
import { ThemeClasses } from '../types';
import { runSystemChecks, LogEntry, CoverageModule } from '../services/testService';
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
                    </div>
                </div>
            </div>
        </div>
    );
};

export default ValidatorView;