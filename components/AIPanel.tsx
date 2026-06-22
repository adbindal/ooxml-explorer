import React, { useState } from 'react';
import { Sparkles, X, Bot, Loader2, Microscope, FileDiff, Plus, Check, Settings, Key, Save, LogOut } from 'lucide-react';
import { 
  analyzeFile, 
  analyzeDiff, 
  DiffFileContext, 
  EditorFileContext, 
  getApiKey, 
  setApiKey, 
  clearApiKey,
  AIAnalysis,
  AIDiffAnalysis
} from '../services/geminiService';
import { ThemeClasses } from '../types';

interface AIPanelProps {
  onClose: () => void;
  context: {
    mode: 'editor' | 'diff';
    fileName?: string;
    
    // Editor specific
    content?: string;
    
    // Diff specific
    diffOriginal?: string;
    diffModified?: string;
    
    // Multi-file context
    relatedFiles?: string[];
    onLoadContext?: (paths: string[]) => Promise<EditorFileContext[] | DiffFileContext[]>; // Returns EditorFileContext[] or DiffFileContext[]
  };
  themeClasses: ThemeClasses;
}

const AIPanel: React.FC<AIPanelProps> = ({ onClose, context, themeClasses }) => {
  const [loading, setLoading] = useState(false);
  const [editorResponse, setEditorResponse] = useState<AIAnalysis | null>(null);
  const [diffResponse, setDiffResponse] = useState<AIDiffAnalysis | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  
  const [activeAction, setActiveAction] = useState<string | null>(null);
  const [selectedRelated, setSelectedRelated] = useState<Set<string>>(new Set());
  const [isExpanded, setIsExpanded] = useState(false);
  
  // API Key Configuration State
  const [needsApiKey, setNeedsApiKey] = useState(() => !getApiKey());
  const [apiKeyInput, setApiKeyInput] = useState('');
  const [showSettings, setShowSettings] = useState(false);

  const handleSaveKey = () => {
      if (!apiKeyInput.trim()) return;
      setApiKey(apiKeyInput.trim());
      setNeedsApiKey(false);
      setShowSettings(false);
  };

  const handleClearKey = () => {
      clearApiKey();
      setNeedsApiKey(true);
      setShowSettings(false);
      setApiKeyInput('');
  };
  
  const handleEditorAction = async (mode: 'explain' | 'technical') => {
    if (!context.fileName || !context.content) return;
    setLoading(true);
    setActiveAction(mode);
    setEditorResponse(null);
    setDiffResponse(null);
    setErrorMessage(null);
    
    try {
        // 1. Prepare Primary Context
        const primaryContext: EditorFileContext = {
            fileName: context.fileName,
            content: context.content
        };

        let finalContext = [primaryContext];

        // 2. Load Additional Context if selected
        if (context.onLoadContext && selectedRelated.size > 0) {
            const extraContext = await context.onLoadContext(Array.from(selectedRelated)) as EditorFileContext[];
            finalContext = [...finalContext, ...extraContext];
        }

        const result = await analyzeFile(finalContext, mode);
        setEditorResponse(result);
    } catch (e: unknown) {
        const err = e as Error;
        if (err.message === 'API_KEY_MISSING') {
            setNeedsApiKey(true);
        } else {
            setErrorMessage(err.message || "Error contacting Gemini. Please check your connection.");
        }
    } finally {
        setLoading(false);
    }
  };

  const handleDiffAction = async (mode: 'summary' | 'technical') => {
    if (!context.fileName || !context.diffOriginal || !context.diffModified) return;
    setLoading(true);
    setActiveAction(mode);
    setEditorResponse(null);
    setDiffResponse(null);
    setErrorMessage(null);
    
    try {
        // 1. Prepare Primary Context
        const primaryContext: DiffFileContext = {
            fileName: context.fileName,
            original: context.diffOriginal,
            modified: context.diffModified
        };

        let finalContext = [primaryContext];

        // 2. Load Additional Context if selected
        if (context.onLoadContext && selectedRelated.size > 0) {
            const extraContext = await context.onLoadContext(Array.from(selectedRelated)) as DiffFileContext[];
            finalContext = [...finalContext, ...extraContext];
        }

        const result = await analyzeDiff(finalContext, mode);
        setDiffResponse(result);
    } catch (e: unknown) {
        const err = e as Error;
        if (err.message === 'API_KEY_MISSING') {
            setNeedsApiKey(true);
        } else {
            console.error(err);
            setErrorMessage(err.message || "Error contacting Gemini.");
        }
    } finally {
        setLoading(false);
    }
  };

  const toggleRelated = (path: string) => {
      const next = new Set(selectedRelated);
      if (next.has(path)) next.delete(path);
      else next.add(path);
      setSelectedRelated(next);
  };

  // --- RENDER DASHBOARDS ---

  const renderEditorDashboard = (data: AIAnalysis) => {
      const hasIssues = data.criticalIssues.length > 0;
      return (
          <div className="space-y-5 animate-in fade-in duration-300 text-xs text-left">
              {/* Summary Card */}
              <div className={`p-4 rounded-xl border ${themeClasses.bg} ${themeClasses.border} shadow-sm space-y-2`}>
                  <div className="flex items-center gap-2 text-blue-500 font-bold text-[10px] uppercase tracking-wider">
                      <Bot size={13} />
                      <span>File Purpose</span>
                  </div>
                  <p className={`text-xs leading-relaxed ${themeClasses.fgMuted}`}>{data.summary}</p>
              </div>

              {/* Critical Issues Panel */}
              <div className="space-y-2.5">
                  <div className="text-[10px] font-bold uppercase tracking-wider">
                      <span className={hasIssues ? "text-amber-500" : "text-green-500"}>
                          {hasIssues ? `⚠️ Compliance Warnings (${data.criticalIssues.length})` : "✅ Compliance Check"}
                      </span>
                  </div>
                  {!hasIssues ? (
                      <div className={`p-3.5 rounded-xl border border-green-500/10 bg-green-500/5 text-green-500/90 leading-relaxed`}>
                          This file complies with standard ECMA-376 specifications. No critical issues or malformations were detected.
                      </div>
                  ) : (
                      <div className="space-y-2.5">
                          {data.criticalIssues.map((issue, idx) => (
                              <div key={idx} className={`p-3.5 rounded-xl border border-amber-500/20 bg-amber-500/5 space-y-2`}>
                                  <div className="font-bold text-amber-500">Issue: {issue.issue}</div>
                                  <div className={themeClasses.fgMuted}>
                                      <span className="font-bold text-[9px] uppercase text-amber-600/70 block mb-0.5">Impact:</span> 
                                      {issue.impact}
                                  </div>
                                  <div className={themeClasses.fgMuted}>
                                      <span className="font-bold text-[9px] uppercase text-green-600/70 block mb-0.5">Remediation:</span> 
                                      {issue.remediation}
                                  </div>
                              </div>
                          ))}
                      </div>
                  )}
              </div>

              {/* Key Elements List */}
              <div className="space-y-2.5">
                  <div className={`text-[10px] font-bold uppercase tracking-wider ${themeClasses.fgMuted}`}>
                      <span>Key XML Tags</span>
                  </div>
                  <div className={`border rounded-xl divide-y ${themeClasses.border} ${themeClasses.bg} overflow-hidden shadow-sm`}>
                      {data.keyElements.map((elem, idx) => (
                          <div key={idx} className="p-3 flex flex-col gap-1.5">
                              <code className="px-2 py-0.5 rounded bg-blue-500/10 text-[#4A89DC] font-mono text-[10px] font-bold shrink-0 self-start border border-blue-500/10">&lt;{elem.tag}&gt;</code>
                              <span className={`${themeClasses.fgMuted} leading-relaxed`}>{elem.purpose}</span>
                          </div>
                      ))}
                  </div>
              </div>
          </div>
      );
  };

  const renderDiffDashboard = (data: AIDiffAnalysis) => {
      return (
          <div className="space-y-5 animate-in fade-in duration-300 text-xs text-left">
              {/* Summary Card */}
              <div className={`p-4 rounded-xl border ${themeClasses.bg} ${themeClasses.border} shadow-sm space-y-2`}>
                  <div className="flex items-center gap-2 text-blue-500 font-bold text-[10px] uppercase tracking-wider">
                      <FileDiff size={13} />
                      <span>Functional Impact</span>
                  </div>
                  <p className={`text-xs leading-relaxed ${themeClasses.fgMuted}`}>{data.summary}</p>
              </div>

              {/* Changes List Panel */}
              <div className="space-y-2.5">
                  <div className={`text-[10px] font-bold uppercase tracking-wider ${themeClasses.fgMuted}`}>
                      <span>Modifications Breakdown</span>
                  </div>
                  <div className="space-y-2.5">
                      {data.changesList.map((change, idx) => {
                          const isAdded = change.changeType === 'added';
                          const isDeleted = change.changeType === 'deleted';
                          const badgeColor = isAdded 
                              ? 'bg-green-500/10 border-green-500/20 text-green-500' 
                              : isDeleted 
                                  ? 'bg-red-500/10 border-red-500/20 text-red-500'
                                  : 'bg-amber-500/10 border-amber-500/20 text-amber-500';
                                  
                          return (
                              <div key={idx} className={`p-3.5 rounded-xl border ${themeClasses.bg} ${themeClasses.border} shadow-sm space-y-2.5`}>
                                  <div className="flex items-center justify-between gap-2">
                                      <code className="px-2 py-0.5 rounded bg-blue-500/5 text-[#4A89DC] font-mono text-[10px] font-bold border border-blue-500/5 truncate max-w-[70%]">&lt;{change.element}&gt;</code>
                                      <span className={`px-1.5 py-0.5 rounded-full text-[8px] font-bold uppercase border tracking-wider ${badgeColor}`}>
                                          {change.changeType}
                                      </span>
                                  </div>
                                  <div className="space-y-2">
                                      <div className={themeClasses.fgMuted}>
                                          <span className="font-bold text-[9px] uppercase block opacity-60 mb-0.5">Description:</span> 
                                          {change.description}
                                      </div>
                                      <div className={themeClasses.fgMuted}>
                                          <span className="font-bold text-[9px] uppercase block opacity-60 mb-0.5">Visual Impact:</span> 
                                          {change.visualImpact}
                                      </div>
                                  </div>
                              </div>
                          );
                      })}
                  </div>
              </div>
          </div>
      );
  };

  // --- RENDER CONFIGURATION ---

  const renderConfiguration = () => (
      <div className="flex-1 flex flex-col p-6 gap-6 items-center justify-center text-center animate-in fade-in">
          <div className="w-16 h-16 rounded-2xl bg-blue-500/10 flex items-center justify-center text-blue-500 mb-2">
              <Key size={32} />
          </div>
          <div className="space-y-2">
              <h3 className={`text-lg font-bold ${themeClasses.fg}`}>Setup Gemini AI</h3>
              <p className={`text-xs ${themeClasses.fgMuted} max-w-[240px] mx-auto leading-relaxed`}>
                  To use AI features, you need a Google Gemini API key. Your key is stored securely in your browser's local storage.
              </p>
          </div>
          
          <div className="w-full space-y-3">
              <input 
                  type="password" 
                  placeholder="Paste API Key here..."
                  className={`w-full border rounded p-3 text-sm focus:outline-none transition-colors ${themeClasses.input}`}
                  value={apiKeyInput}
                  onChange={(e) => setApiKeyInput(e.target.value)}
              />
              <button 
                  onClick={handleSaveKey}
                  disabled={!apiKeyInput.trim()}
                  className="w-full bg-blue-600 hover:bg-blue-500 disabled:opacity-50 disabled:cursor-not-allowed text-white font-bold py-2.5 rounded flex items-center justify-center gap-2 transition-all shadow-md hover:shadow-lg"
              >
                  <Save size={16} /> Save Key
              </button>
          </div>
          
          <a href="https://aistudio.google.com/app/apikey" target="_blank" rel="noreferrer" className="text-xs text-blue-400 hover:text-blue-300 underline">
              Get an API Key from Google AI Studio
          </a>
      </div>
  );

  return (
    <div className={`w-full h-full flex flex-col shadow-2xl z-20 ${themeClasses.bgPanel} ${themeClasses.border}`}>
      <div className={`h-12 border-b flex items-center justify-between px-4 shrink-0 ${themeClasses.bgSec} ${themeClasses.border}`}>
        <div className="flex items-center gap-2 text-blue-500 font-medium">
            <Sparkles size={16} />
            <span>Gemini Assistant</span>
        </div>
        <div className="flex items-center gap-2">
            {!needsApiKey && (
                <button 
                    onClick={() => setShowSettings(!showSettings)} 
                    className={`p-1.5 rounded transition-colors ${showSettings ? 'text-blue-500 bg-blue-500/10' : `${themeClasses.fgMuted} ${themeClasses.hover}`}`}
                    title="Settings"
                >
                    <Settings size={14} />
                </button>
            )}
            <button onClick={onClose} className={`${themeClasses.fgMuted} hover:${themeClasses.fg}`}>
                <X size={16} />
            </button>
        </div>
      </div>

      {needsApiKey ? (
          renderConfiguration()
      ) : showSettings ? (
          <div className="flex-1 p-6 flex flex-col gap-4 animate-in fade-in slide-in-from-right-4">
              <h3 className={`font-bold flex items-center gap-2 ${themeClasses.fg}`}><Settings size={16}/> Settings</h3>
              
              <div className={`p-4 rounded border space-y-3 ${themeClasses.bg} ${themeClasses.border}`}>
                  <div className="flex items-center justify-between">
                      <span className={`text-xs ${themeClasses.fgMuted}`}>API Key Status</span>
                      <span className="text-xs text-green-500 flex items-center gap-1"><Check size={12}/> Configured</span>
                  </div>
                  <div className={`text-[10px] ${themeClasses.fgMuted} font-mono`}>
                      Stored locally in browser
                  </div>
              </div>

              <button 
                  onClick={handleClearKey}
                  className="w-full border border-red-900/30 bg-red-900/10 hover:bg-red-900/20 text-red-500 py-2 rounded text-xs font-bold flex items-center justify-center gap-2 transition-colors"
              >
                  <LogOut size={14} /> Remove API Key
              </button>
              
              <div className={`mt-auto border-t pt-4 ${themeClasses.border}`}>
                 <button onClick={() => setShowSettings(false)} className={`w-full py-2 rounded text-xs font-medium ${themeClasses.bg} ${themeClasses.hover} ${themeClasses.fg}`}>
                     Back to Assistant
                 </button>
              </div>
          </div>
      ) : (
        <div className="p-4 flex flex-col gap-4 flex-1 overflow-hidden">
            {/* Context Header */}
            <div className={`p-3 rounded text-xs border shadow-sm ${themeClasses.bg} ${themeClasses.border}`}>
                <span className={`font-semibold block mb-1 ${themeClasses.fg}`}>Active File:</span>
                <div className={`truncate ${themeClasses.fgMuted}`}>{context.fileName || 'No file selected'}</div>
                
                {/* Context Selector for BOTH Editor and Diff Mode */}
                {context.relatedFiles && context.relatedFiles.length > 0 && (
                    <div className={`mt-3 pt-3 border-t ${themeClasses.border}`}>
                        <button 
                            onClick={() => setIsExpanded(!isExpanded)}
                            className={`flex items-center gap-2 ${themeClasses.fgMuted} hover:${themeClasses.fg} transition-colors w-full`}
                        >
                            <Plus size={12} className={isExpanded ? "rotate-45 transition-transform" : "transition-transform"} />
                            <span className="font-semibold">Add Context ({selectedRelated.size})</span>
                        </button>
                        
                        {isExpanded && (
                            <div className="mt-2 max-h-32 overflow-y-auto space-y-1 pr-1 scrollbar-thin">
                                {context.relatedFiles.filter(f => f !== context.fileName).map(path => (
                                    <div 
                                        key={path} 
                                        onClick={() => toggleRelated(path)}
                                        className={`
                                            flex items-center gap-2 p-1.5 rounded cursor-pointer border transition-colors
                                            ${selectedRelated.has(path) 
                                                ? 'bg-blue-500/20 border-blue-500/50 text-blue-500' 
                                                : `${themeClasses.bgPanel} ${themeClasses.border} ${themeClasses.fgMuted} hover:${themeClasses.hover}`
                                            }
                                        `}
                                    >
                                        <div className={`w-3 h-3 rounded-sm border flex items-center justify-center ${selectedRelated.has(path) ? 'bg-blue-500 border-blue-500' : themeClasses.border}`}>
                                            {selectedRelated.has(path) && <Check size={10} className="text-white" />}
                                        </div>
                                        <span className="truncate flex-1" title={path}>{path}</span>
                                    </div>
                                ))}
                            </div>
                        )}
                    </div>
                )}
            </div>

            {/* Action Buttons */}
            <div className="grid grid-cols-1 gap-2 shrink-0">
                {context.mode === 'editor' ? (
                    <>
                        <button 
                            onClick={() => handleEditorAction('explain')}
                            disabled={loading || !context.fileName}
                            className={`
                                group flex items-center gap-3 p-3 rounded-lg border text-left transition-all shadow-sm
                                ${activeAction === 'explain' 
                                    ? 'bg-blue-500/10 border-blue-500/50' 
                                    : `${themeClasses.bg} ${themeClasses.border} ${themeClasses.hover}`
                                }
                                disabled:opacity-50 disabled:cursor-not-allowed
                            `}
                        >
                            <Bot size={18} className="text-blue-500 shrink-0" />
                            <div className="flex-1">
                                <span className={`block text-xs font-bold ${themeClasses.fg}`}>Explain Purpose</span>
                                <span className={`block text-[10px] ${themeClasses.fgMuted} mt-0.5`}>High-level overview of this file's role.</span>
                            </div>
                        </button>

                        <button 
                            onClick={() => handleEditorAction('technical')}
                            disabled={loading || !context.fileName}
                            className={`
                                group flex items-center gap-3 p-3 rounded-lg border text-left transition-all shadow-sm
                                ${activeAction === 'technical' 
                                    ? 'bg-slate-500/10 border-slate-500/50' 
                                    : `${themeClasses.bg} ${themeClasses.border} ${themeClasses.hover}`
                                }
                                disabled:opacity-50 disabled:cursor-not-allowed
                            `}
                        >
                            <Microscope size={18} className="text-slate-500 shrink-0" />
                            <div className="flex-1">
                                <span className={`block text-xs font-bold ${themeClasses.fg}`}>Technical Analysis</span>
                                <span className={`block text-[10px] ${themeClasses.fgMuted} mt-0.5`}>Validate structure and inspect attributes.</span>
                            </div>
                        </button>
                    </>
                ) : (
                    <>
                        <button 
                            onClick={() => handleDiffAction('summary')}
                            disabled={loading || !context.fileName}
                            className={`
                                group flex items-center gap-3 p-3 rounded-lg border text-left transition-all shadow-sm
                                ${activeAction === 'summary' 
                                    ? 'bg-blue-500/10 border-blue-500/50' 
                                    : `${themeClasses.bg} ${themeClasses.border} ${themeClasses.hover}`
                                }
                                disabled:opacity-50 disabled:cursor-not-allowed
                            `}
                        >
                            <FileDiff size={18} className="text-blue-500 shrink-0" />
                            <div className="flex-1">
                                <span className={`block text-xs font-bold ${themeClasses.fg}`}>Change Summary</span>
                                <span className={`block text-[10px] ${themeClasses.fgMuted} mt-0.5`}>Plain English impact of the changes.</span>
                            </div>
                        </button>

                        <button 
                            onClick={() => handleDiffAction('technical')}
                            disabled={loading || !context.fileName}
                            className={`
                                group flex items-center gap-3 p-3 rounded-lg border text-left transition-all shadow-sm
                                ${activeAction === 'technical' 
                                    ? 'bg-slate-500/10 border-slate-500/50' 
                                    : `${themeClasses.bg} ${themeClasses.border} ${themeClasses.hover}`
                                }
                                disabled:opacity-50 disabled:cursor-not-allowed
                            `}
                        >
                            <Microscope size={18} className="text-slate-500 shrink-0" />
                            <div className="flex-1">
                                <span className={`block text-xs font-bold ${themeClasses.fg}`}>Technical Analysis</span>
                                <span className={`block text-[10px] ${themeClasses.fgMuted} mt-0.5`}>Deep dive into XML delta mechanics.</span>
                            </div>
                        </button>
                    </>
                )}
            </div>

            {/* Output Area - Structured Dashboards */}
            <div className={`flex-1 rounded-lg border p-4 overflow-y-auto relative min-h-0 shadow-inner ${themeClasses.input}`}>
                {loading ? (
                    <div className={`absolute inset-0 flex flex-col items-center justify-center gap-3 backdrop-blur-sm z-10 ${themeClasses.bg}/80`}>
                        <Loader2 className="animate-spin text-blue-500" size={24} />
                        <span className={`text-xs ${themeClasses.fgMuted} animate-pulse`}>Consulting Gemini...</span>
                    </div>
                ) : errorMessage ? (
                    <div className="h-full flex flex-col items-center justify-center text-center p-4 gap-3 text-red-500">
                        <span className="text-sm font-semibold">Analysis Failed</span>
                        <p className={`text-xs ${themeClasses.fgMuted} max-w-[240px]`}>{errorMessage}</p>
                    </div>
                ) : editorResponse ? (
                    renderEditorDashboard(editorResponse)
                ) : diffResponse ? (
                    renderDiffDashboard(diffResponse)
                ) : (
                    <div className={`h-full flex flex-col items-center justify-center text-center gap-2 ${themeClasses.fgMuted}`}>
                        <Sparkles size={32} className="opacity-20" />
                        <p className="text-xs">Select an action above to analyze.</p>
                    </div>
                )}
            </div>
        </div>
      )}
    </div>
  );
};

export default AIPanel;