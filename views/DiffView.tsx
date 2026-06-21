import React, { useState, useEffect, useRef } from 'react';
import { DiffEditor, DiffOnMount } from '@monaco-editor/react';
import { ThemeClasses } from '../types';
import FileTree from '../components/FileTree';
import AIPanel from '../components/AIPanel';
import ErrorBoundary from '../components/ErrorBoundary';
import { DiffFileContext } from '../services/geminiService';
import { formatXml, isXmlFile, isImageFile } from '../utils/xmlUtils';
import { getModifiedPaths } from '../utils/treeUtils';
import { shouldAutoRunDiff } from '../utils/diffUtils';
import Logo from '../components/Logo';
import { useAppStore } from '../store/appStore';
import { defineMonacoThemes } from '../utils/theme';
import { 
  ArrowLeft, PanelLeftClose, PanelLeftOpen, Sparkles, 
  ArrowRight, FileDiff, RefreshCw, Upload, GitCompare,
  ArrowUp, ArrowDown, Rows, Columns, Sun, Moon, Eye, EyeOff, Search,
  ArrowLeftRight, X
} from 'lucide-react';

interface DiffViewProps {
  themeClasses: ThemeClasses;
}

const DiffView: React.FC<DiffViewProps> = ({ themeClasses }) => {
  const { 
      diff: diffState, 
      ui, 
      theme,
      mode, // Import mode from store
      setMode, 
      toggleSidebar, 
      toggleAiPanel, 
      toggleTheme,
      setDiffFiles,
      runDiffComparison,
      updateDiffState,
      setDiffViewMode
  } = useAppStore();

  const { sidebarOpen, showAi, diffViewMode } = ui;
  const { originalFile, modifiedFile, originalZip, modifiedZip, tree, activePath, loading } = diffState;

  const [showUnchanged, setShowUnchanged] = useState(false); 
  const [searchQuery, setSearchQuery] = useState('');
  const [diffContent, setDiffContent] = useState<{ original: string | null; modified: string | null; isImage: boolean }>({ original: null, modified: null, isImage: false });
  const [imageUrls, setImageUrls] = useState<{ original: string | null; modified: string | null }>({ original: null, modified: null });
  
  // Navigation State
  interface DiffLineChange {
      readonly originalStartLineNumber: number;
      readonly originalEndLineNumber: number;
      readonly modifiedStartLineNumber: number;
      readonly modifiedEndLineNumber: number;
  }
  const [diffChanges, setDiffChanges] = useState<DiffLineChange[]>([]);
  const [currentDiffIndex, setCurrentDiffIndex] = useState(-1);

  // Drag and Drop States & Refs
  const [dragCounter, setDragCounter] = useState(0);
  const [dragFileCount, setDragFileCount] = useState(0);
  const [boxDrag, setBoxDrag] = useState<{ original: boolean; modified: boolean }>({ original: false, modified: false });

  const originalInputRef = useRef<HTMLInputElement>(null);
  const modifiedInputRef = useRef<HTMLInputElement>(null);

  const isMounted = useRef(false);
  const diffEditorRef = useRef<Parameters<DiffOnMount>[0] | null>(null);
  const diffListenerRef = useRef<{ dispose(): void } | null>(null);
  const monacoRef = useRef<Parameters<DiffOnMount>[1] | null>(null);

  const modColorClass = theme === 'dark' ? 'text-[#A5B4FC]' : 'text-[#1F3F70]';
  const modBorderClass = theme === 'dark' ? 'border-[#A5B4FC]' : 'border-[#1F3F70]';
  const modBgClass = theme === 'dark' ? 'bg-[#A5B4FC]/10' : 'bg-[#1F3F70]/10';
  const modHoverBorderClass = theme === 'dark' ? 'hover:border-[#A5B4FC]/50' : 'hover:border-[#1F3F70]/50';

  useEffect(() => {
    isMounted.current = true;
    console.log("[DiffView] Mounted. Mode:", mode);
    if (window.innerWidth < 768) {
        toggleSidebar(false);
        setDiffViewMode('inline');
    }

    // Consolidated Auto-Run Logic
    const shouldRun = shouldAutoRunDiff(mode, !!originalFile, !!modifiedFile, !!tree);
    
    if (shouldRun) {
         runDiffComparison().catch(e => alert(e.message));
    }

    return () => {
        isMounted.current = false;
        if (diffListenerRef.current?.dispose) diffListenerRef.current.dispose();
        if (diffEditorRef.current) {
             try {
                 const original = diffEditorRef.current.getOriginalEditor();
                 const modified = diffEditorRef.current.getModifiedEditor();
                 original?.getModel()?.dispose();
                 modified?.getModel()?.dispose();
             } catch {
                 // Ignore cleanup error
             }
        }
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []); // Run only on mount

  // Clean up any remaining image object URLs on unmount to prevent memory leaks
  useEffect(() => {
    return () => {
      setImageUrls(prev => {
        if (prev.original) URL.revokeObjectURL(prev.original);
        if (prev.modified) URL.revokeObjectURL(prev.modified);
        return { original: null, modified: null };
      });
    };
  }, []);

  // Update Monaco theme dynamically when app theme changes
  useEffect(() => {
      if (monacoRef.current) {
          monacoRef.current.editor.setTheme(theme === 'dark' ? 'ooxml-dark' : 'ooxml-light');
      }
  }, [theme]);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
        const isCmdOrCtrl = e.metaKey || e.ctrlKey;
        if (isCmdOrCtrl && e.key.toLowerCase() === 'b') {
            e.preventDefault();
            toggleSidebar();
        }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const handleCompare = async () => {
      runDiffComparison().catch(e => alert(e.message));
  };

  const handleSwap = () => {
      setDiffFiles(modifiedFile, originalFile);
  };

  const handleFetchContext = async (paths: string[]): Promise<DiffFileContext[]> => {
      if (!originalZip || !modifiedZip) return [];
      const results = await Promise.all(paths.map(async (path) => {
          if (!isMounted.current) return null;
          if (isImageFile(path)) return null; 

          const textA = originalZip.file(path) ? await originalZip.file(path)!.async('string') : null;
          const textB = modifiedZip.file(path) ? await modifiedZip.file(path)!.async('string') : null;
          
          const fmtA = textA && isXmlFile(path) ? formatXml(textA) : textA;
          const fmtB = textB && isXmlFile(path) ? formatXml(textB) : textB;

          return { fileName: path, original: fmtA, modified: fmtB };
      }));
      return results.filter(Boolean) as DiffFileContext[];
  };

  useEffect(() => {
    let active = true;
    const loadContent = async () => {
        if (!activePath || !originalZip || !modifiedZip) return;

        if (imageUrls.original) URL.revokeObjectURL(imageUrls.original);
        if (imageUrls.modified) URL.revokeObjectURL(imageUrls.modified);

        if (isImageFile(activePath)) {
             const blobA = originalZip.file(activePath) ? await originalZip.file(activePath)!.async('blob') : null;
             const blobB = modifiedZip.file(activePath) ? await modifiedZip.file(activePath)!.async('blob') : null;
             if (!active || !isMounted.current) return;
             setImageUrls({
                 original: blobA ? URL.createObjectURL(blobA) : null,
                 modified: blobB ? URL.createObjectURL(blobB) : null
             });
             setDiffContent({ original: null, modified: null, isImage: true });
             setCurrentDiffIndex(-1);
             setDiffChanges([]);
        } else {
             const textA = originalZip.file(activePath) ? await originalZip.file(activePath)!.async('string') : '';
             const textB = modifiedZip.file(activePath) ? await modifiedZip.file(activePath)!.async('string') : '';
             if (!active || !isMounted.current) return;
             const fmtA = isXmlFile(activePath) ? formatXml(textA) : textA;
             const fmtB = isXmlFile(activePath) ? formatXml(textB) : textB;
             setDiffContent({ original: fmtA, modified: fmtB, isImage: false });
             setImageUrls({ original: null, modified: null });
             setCurrentDiffIndex(-1);
             setDiffChanges([]);
        }
    };
    loadContent();
    return () => { active = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activePath]);

  // Ensure word wrap is applied to both sides of the diff editor
  useEffect(() => {
    if (diffEditorRef.current) {
      try {
        const original = diffEditorRef.current.getOriginalEditor();
        const modified = diffEditorRef.current.getModifiedEditor();
        original.updateOptions({ wordWrap: 'on' });
        modified.updateOptions({ wordWrap: 'on' });
      } catch (e) {
        console.error("Failed to update diff editor options:", e);
      }
    }
  }, [activePath, diffViewMode, diffContent]);

  const handleFileChange = (type: 'original' | 'modified', file: File | null) => {
    if (type === 'original') setDiffFiles(file, modifiedFile);
    else setDiffFiles(originalFile, file);
  };
  
  // Navigation Logic
  const navigateDiff = (direction: 'next' | 'prev') => {
    if (!diffChanges.length || !diffEditorRef.current) return;
    
    let nextIndex = direction === 'next' ? currentDiffIndex + 1 : currentDiffIndex - 1;
    // Wrap around
    if (nextIndex >= diffChanges.length) nextIndex = 0;
    if (nextIndex < 0) nextIndex = diffChanges.length - 1;
    
    setCurrentDiffIndex(nextIndex);
    
    const change = diffChanges[nextIndex];
    if (change && diffEditorRef.current) {
        const editor = diffEditorRef.current.getModifiedEditor();
        // Determine safe target line. Insertions/Mods have valid modified line numbers. 
        // Deletions might technically be 0, so we clamp to 1 or rely on original mapping context.
        const line = Math.max(1, change.modifiedStartLineNumber > 0 ? change.modifiedStartLineNumber : change.originalStartLineNumber);
        
        // Center and focus
        editor.revealLineInCenter(line);
        editor.setPosition({ lineNumber: line, column: 1 });
        editor.focus();
    }
  };

  const handleDiffMount: DiffOnMount = (editor, monaco) => {
      if (!isMounted.current) return;
      diffEditorRef.current = editor;
      monacoRef.current = monaco;

      // Define shared themes
      defineMonacoThemes(monaco);
      monaco.editor.setTheme(theme === 'dark' ? 'ooxml-dark' : 'ooxml-light');
      
      // Cleanup previous listener
      if (diffListenerRef.current?.dispose) diffListenerRef.current.dispose();

      const updateDiffs = () => {
          // getLineChanges() returns null if diff not computed yet
          const changes = editor.getLineChanges() || [];
          setDiffChanges(changes);
      };
      
      // Listen for diff updates (async computation)
      diffListenerRef.current = editor.onDidUpdateDiff(updateDiffs);
      
      // Initial trigger attempt
      updateDiffs();

      if (editor) {
          try {
            // Initial option application to both sub-editors
            editor.getOriginalEditor().updateOptions({ wordWrap: 'on' });
            editor.getModifiedEditor().updateOptions({ wordWrap: 'on' });

            // Enforce word wrap on both sides whenever a model is updated/loaded (e.g. changing active files)
            editor.getOriginalEditor().onDidChangeModel(() => {
                try {
                    editor.getOriginalEditor().updateOptions({ wordWrap: 'on' });
                } catch (e) { console.error(e); }
            });
            editor.getModifiedEditor().onDidChangeModel(() => {
                try {
                    editor.getModifiedEditor().updateOptions({ wordWrap: 'on' });
                } catch (e) { console.error(e); }
            });
          } catch (e) { console.error(e); }
      }
  };

  // Drag and Drop Handlers
  const handleGlobalDragEnter = (e: React.DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
      setDragCounter(prev => prev + 1);
      if (e.dataTransfer.items) {
          setDragFileCount(e.dataTransfer.items.length);
      }
  };

  const handleGlobalDragLeave = (e: React.DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
      setDragCounter(prev => prev - 1);
  };

  const handleGlobalDragOver = (e: React.DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
  };

  const handleGlobalDrop = (e: React.DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
      setDragCounter(0);
      setDragFileCount(0);
      
      const droppedFiles = e.dataTransfer.files;
      if (droppedFiles && droppedFiles.length >= 2) {
          setDiffFiles(droppedFiles[0], droppedFiles[1]);
      }
  };

  const handleBoxDragEnter = (type: 'original' | 'modified', e: React.DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
      setBoxDrag(prev => ({ ...prev, [type]: true }));
  };

  const handleBoxDragLeave = (type: 'original' | 'modified', e: React.DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
      setBoxDrag(prev => ({ ...prev, [type]: false }));
  };

  const handleBoxDragOver = (e: React.DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
  };

  const handleBoxDrop = (type: 'original' | 'modified', e: React.DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
      setBoxDrag(prev => ({ ...prev, [type]: false }));
      
      const droppedFiles = e.dataTransfer.files;
      if (droppedFiles && droppedFiles.length > 0) {
          handleFileChange(type, droppedFiles[0]);
      }
  };

  const renderUploadBox = (type: 'original' | 'modified', file: File | null) => {
      const isOriginal = type === 'original';
      const iconColor = isOriginal ? 'text-[#4A89DC]' : modColorClass;
      
      const isDragging = boxDrag[type];
      const borderColor = isDragging
        ? (isOriginal ? 'border-[#4A89DC] bg-[#4A89DC]/10 scale-[1.01]' : `${modBorderClass} bg-[#A5B4FC]/10 scale-[1.01]`)
        : (file 
            ? (isOriginal ? 'border-[#4A89DC]' : modBorderClass) 
            : `${themeClasses.border} ${isOriginal ? 'hover:border-[#4A89DC]/50' : modHoverBorderClass}`);
            
      const bgColor = file ? (isOriginal ? 'bg-[#4A89DC]/10' : modBgClass) : '';
      const inputRef = isOriginal ? originalInputRef : modifiedInputRef;

      return (
        <div className="space-y-4 w-full">
            <label className={`block text-sm font-medium uppercase tracking-wider ${themeClasses.fgMuted}`}>
                {type} File {file && <span className={`ml-2 text-xs normal-case ${iconColor}`}>Ready</span>}
            </label>
            <div 
                onClick={() => inputRef.current?.click()}
                onDragEnter={(e) => handleBoxDragEnter(type, e)}
                onDragLeave={(e) => handleBoxDragLeave(type, e)}
                onDragOver={handleBoxDragOver}
                onDrop={(e) => handleBoxDrop(type, e)}
                className={`
                    min-h-[8rem] md:min-h-[10rem] p-4 border-2 border-dashed rounded-xl flex flex-col items-center justify-center relative transition-all duration-200 cursor-pointer
                    ${borderColor} ${bgColor}
                `}
            >
                <input 
                    ref={inputRef}
                    type="file" 
                    className="hidden" 
                    accept=".docx,.xlsx,.pptx"
                    onChange={(e) => handleFileChange(type, e.target.files?.[0] || null)} 
                />
                {file ? (
                    <div className="text-center pointer-events-none w-full">
                        <FileDiff className={`mx-auto mb-3 ${iconColor}`} size={32} />
                        <p className={`text-sm font-bold ${themeClasses.fg} px-4 break-all`}>{file.name}</p>
                        <p className={`text-xs ${themeClasses.fgMuted} mt-1`}>{(file.size / 1024).toFixed(1)} KB</p>
                    </div>
                ) : (
                    <div className={`text-center pointer-events-none ${themeClasses.fgMuted}`}>
                        <Upload className={`mx-auto mb-3 opacity-50`} size={24} />
                        <p className="text-sm font-medium">Click or Drag to upload {type}</p>
                    </div>
                )}
            </div>
        </div>
      );
  };

  if (!tree) {
    const showGlobalOverlay = dragCounter > 0 && dragFileCount >= 2;

    return (
        <div 
            onDragEnter={handleGlobalDragEnter}
            onDragLeave={handleGlobalDragLeave}
            onDragOver={handleGlobalDragOver}
            onDrop={handleGlobalDrop}
            className={`min-h-screen w-full flex flex-col items-center justify-center p-4 md:p-6 transition-colors duration-300 relative ${themeClasses.bg} ${themeClasses.fg}`}
        >
             {/* Global Drag & Drop Overlay */}
             {showGlobalOverlay && (
                <div className="absolute inset-0 z-50 bg-[#4A89DC]/20 backdrop-blur-md m-4 rounded-3xl flex flex-col items-center justify-center pointer-events-none overflow-hidden">
                    <div className="absolute inset-0 border-4 border-[#4A89DC] border-dashed rounded-3xl opacity-50"></div>
                    <div className={`
                        relative z-10 p-12 rounded-3xl shadow-2xl flex flex-col items-center gap-6 text-center
                        ${theme === 'dark' ? 'bg-[#0B1221] border border-[#1F3F70]' : 'bg-white border border-blue-100'}
                    `}>
                        <GitCompare size={80} className="text-[#4A89DC] drop-shadow-lg animate-pulse" />
                        <div className="space-y-2">
                            <span className="block text-2xl font-bold text-[#4A89DC]">Drop both files to Compare</span>
                            <p className={`text-sm ${theme === 'dark' ? 'text-blue-200' : 'text-blue-600'} opacity-80`}>
                                Release to automatically load both documents and run the comparison!
                            </p>
                        </div>
                    </div>
                </div>
             )}

             <div className={`max-w-4xl w-full border rounded-2xl p-4 md:p-8 shadow-2xl ${themeClasses.card}`}>
                <div className="flex items-center justify-between mb-4 md:mb-8">
                    <div className="flex items-center gap-4">
                        <button onClick={() => setMode('landing')} className={themeClasses.hoverText}><ArrowLeft /></button>
                        <h2 className="text-lg md:text-2xl font-bold flex items-center gap-2">
                            <Logo size={24} theme={theme} />
                            Compare Documents
                        </h2>
                    </div>
                    <button onClick={toggleTheme} className={`p-2 rounded-full ${themeClasses.hover}`}>
                        {theme === 'dark' ? <Sun size={20} /> : <Moon size={20} />}
                    </button>
                </div>
                
                <div className="flex flex-col md:flex-row items-center gap-4 md:gap-8 mb-4 md:mb-8 relative">
                    <div className="flex-1 w-full group">{renderUploadBox('original', originalFile)}</div>
                    <div className="flex md:flex-col items-center justify-center pt-2 md:pt-8">
                        <button onClick={handleSwap} className={`p-3 rounded-full border shadow-lg ${themeClasses.bgPanel} ${themeClasses.border} hover:scale-110`}>
                            <ArrowLeftRight size={20} className={themeClasses.fgMuted} />
                        </button>
                    </div>
                    <div className="flex-1 w-full group">{renderUploadBox('modified', modifiedFile)}</div>
                </div>

                <div className="flex justify-end pt-4 border-t border-dashed border-[#4A89DC]/20">
                    <button onClick={handleCompare} disabled={!originalFile || !modifiedFile || loading}
                        className="w-full md:w-auto bg-[#4A89DC] hover:bg-[#3b75c0] disabled:opacity-50 text-white px-8 py-3 rounded-lg font-bold flex items-center justify-center gap-2"
                    >
                        {loading ? <RefreshCw className="animate-spin" /> : <GitCompare />} Run Comparison
                    </button>
                </div>
             </div>
        </div>
    );
  }

  const modifiedPaths = tree ? getModifiedPaths(tree) : [];

  return (
    <div className={`h-screen w-full flex flex-col transition-colors duration-300 ${themeClasses.bg} ${themeClasses.fg} overflow-hidden`}>
       <header className={`h-12 border-b flex items-center justify-between px-2 md:px-4 shrink-0 ${themeClasses.bgSec} ${themeClasses.border}`}>
          <div className="flex items-center gap-2 md:gap-4 overflow-hidden">
            <button onClick={() => setMode('landing')} className={`transition-colors ${themeClasses.hoverText} shrink-0`}><ArrowLeft size={18} /></button>
            <button onClick={() => toggleSidebar()} className={`${themeClasses.icon} hover:${themeClasses.fg} p-1 rounded transition-colors shrink-0`}>
                {sidebarOpen ? <PanelLeftClose size={18} /> : <PanelLeftOpen size={18} />}
            </button>
            <div className="flex items-center gap-2 font-semibold text-sm ml-2 truncate">
                <span className="bg-[#4A89DC] text-white px-2 py-0.5 rounded text-xs shrink-0">DIFF</span>
                <span className="text-[#4A89DC] truncate max-w-[80px] md:max-w-xs">{originalFile?.name}</span>
                <ArrowRight size={14} className={`${themeClasses.fgMuted} shrink-0`} />
                <span className={`${modColorClass} opacity-90 truncate max-w-[80px] md:max-w-xs`}>{modifiedFile?.name}</span>
            </div>
          </div>
          <div className="flex items-center gap-2 md:gap-3 shrink-0">
             <button onClick={() => toggleAiPanel()} 
                disabled={!activePath}
                className={`flex items-center gap-2 px-3 py-1.5 rounded text-xs font-medium border ${themeClasses.border} ${showAi ? 'bg-[#4A89DC]/20 text-[#4A89DC]' : `${themeClasses.bgPanel} ${themeClasses.hover}`}`} 
             >
                <Sparkles size={14} /> <span className="hidden sm:inline">Explain Changes</span>
             </button>
             <button onClick={toggleTheme} className={`p-2 rounded-full ${themeClasses.hover} ${themeClasses.icon} ${themeClasses.hoverText}`}>{theme === 'dark' ? <Sun size={18} /> : <Moon size={18} />}</button>
          </div>
       </header>

       <div className="flex-1 flex overflow-hidden relative">
          {sidebarOpen && (
             <aside className={`
                absolute md:relative z-30 h-full w-64 md:w-64 border-r flex flex-col shrink-0 shadow-2xl md:shadow-none ${themeClasses.bgSidebar} ${themeClasses.border}
             `}>
                <div className={`p-3 border-b space-y-3 ${themeClasses.border} ${themeClasses.bgSec} relative`}>
                    <div className="flex items-center justify-between">
                       <span className={`text-xs font-bold uppercase tracking-wider ${themeClasses.fgMuted}`}>Files</span>
                       <button onClick={() => setShowUnchanged(!showUnchanged)} className={`text-xs flex items-center gap-1.5 px-2 py-1 rounded border transition-colors ${themeClasses.border} ${themeClasses.hover} ${showUnchanged ? themeClasses.fg : themeClasses.fgMuted}`}>
                          {showUnchanged ? <Eye size={12}/> : <EyeOff size={12}/>} {showUnchanged ? 'All' : 'Diffs'}
                       </button>
                    </div>
                    <div className="relative">
                        <Search className={`absolute left-2 top-2 ${themeClasses.fgMuted}`} size={14} />
                        <input type="text" placeholder="Filter files..." className={`w-full text-sm rounded pl-8 pr-2 py-1.5 border focus:outline-none ${themeClasses.input}`} 
                            value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} 
                        />
                    </div>
                     <button className="md:hidden absolute -right-0 -top-0 p-3 text-[#1F3F70]" onClick={() => toggleSidebar(false)}>
                        <X size={14} />
                    </button>
                </div>
                <div className="flex-1 overflow-y-auto py-2">
                    <FileTree 
                        node={tree} 
                        onSelect={(node) => { updateDiffState({ activePath: node.path }); if(window.innerWidth < 768) toggleSidebar(false); }} 
                        activePath={activePath}
                        isDiffMode={true} 
                        showUnchanged={showUnchanged}
                        searchQuery={searchQuery}
                    />
                </div>
             </aside>
          )}

          {sidebarOpen && ( <div className="md:hidden absolute inset-0 bg-black/50 z-20" onClick={() => toggleSidebar(false)}></div> )}

          <main className={`flex-1 flex flex-col overflow-hidden relative min-w-0 ${themeClasses.bg}`}>
             <div className={`h-10 border-b flex items-center justify-between px-4 shrink-0 ${themeClasses.bgSec} ${themeClasses.border}`}>
                <div className="flex items-center gap-3 overflow-hidden">
                    <span className={`text-xs font-semibold ${themeClasses.fg} truncate`}>{activePath || "Select a file"}</span>
                </div>
                {activePath && !diffContent.isImage && (
                   <div className="flex items-center gap-3">
                      <div className="flex items-center gap-1">
                          <span className={`text-[10px] uppercase font-bold tracking-wider mr-1 ${themeClasses.fgMuted} hidden sm:inline`}>
                              {diffChanges.length > 0 ? `${currentDiffIndex + 1} / ${diffChanges.length}` : 'No'} Changes
                          </span>
                          <div className={`flex items-center rounded-md border ${themeClasses.border} bg-opacity-50 overflow-hidden`}>
                              <button onClick={() => navigateDiff('prev')} disabled={diffChanges.length === 0} className={`p-1.5 hover:bg-[#4A89DC]/10 transition-colors ${themeClasses.fg} disabled:opacity-30`}><ArrowUp size={14} /></button>
                              <div className={`w-px h-4 ${themeClasses.border}`}></div>
                              <button onClick={() => navigateDiff('next')} disabled={diffChanges.length === 0} className={`p-1.5 hover:bg-[#4A89DC]/10 transition-colors ${themeClasses.fg} disabled:opacity-30`}><ArrowDown size={14} /></button>
                          </div>
                      </div>
                      <div className={`w-px h-4 ${themeClasses.border}`}></div>
                      <div className={`flex items-center rounded-md border ${themeClasses.border} bg-opacity-50 overflow-hidden p-0.5 gap-0.5`}>
                          <button onClick={() => setDiffViewMode('inline')} className={`p-1 rounded-sm transition-colors ${diffViewMode === 'inline' ? 'bg-[#4A89DC]/20 text-[#4A89DC]' : themeClasses.fgMuted} hover:${themeClasses.fg}`}><Rows size={14} /></button>
                          <button onClick={() => setDiffViewMode('split')} className={`p-1 rounded-sm transition-colors ${diffViewMode === 'split' ? 'bg-[#4A89DC]/20 text-[#4A89DC]' : themeClasses.fgMuted} hover:${themeClasses.fg}`}><Columns size={14} /></button>
                      </div>
                   </div>
                )}
             </div>

             <div className="flex-1 relative">
                {!activePath ? (
                    <div className={`absolute inset-0 flex flex-col items-center justify-center ${themeClasses.fgMuted} p-4 text-center`}>
                        <FileDiff size={48} className="mb-4 opacity-20" />
                        <p>Select a file to compare</p>
                    </div>
                ) : diffContent.isImage ? (
                    <div className="flex-1 flex flex-col md:flex-row items-center justify-center gap-6 md:gap-10 p-4 md:p-8 overflow-y-auto h-full">
                        <div className="text-center w-full md:w-auto">
                            <div className={`mb-2 text-sm font-medium text-[#4A89DC]`}>Original</div>
                            {imageUrls.original ? <img src={imageUrls.original} className={`max-w-full md:max-w-sm max-h-64 md:max-h-96 border rounded mx-auto ${themeClasses.border} ${themeClasses.bgPanel}`} /> : <div className={`w-full md:w-64 h-32 md:h-64 border border-dashed rounded flex items-center justify-center mx-auto ${themeClasses.border} ${themeClasses.fgMuted}`}>Deleted</div>}
                        </div>
                        <div className="text-center w-full md:w-auto">
                            <div className={`mb-2 text-sm font-medium ${modColorClass} opacity-80`}>Modified</div>
                            {imageUrls.modified ? <img src={imageUrls.modified} className={`max-w-full md:max-w-sm max-h-64 md:max-h-96 border rounded mx-auto ${themeClasses.border} ${themeClasses.bgPanel}`} /> : <div className={`w-full md:w-64 h-32 md:h-64 border border-dashed rounded flex items-center justify-center mx-auto ${themeClasses.border} ${themeClasses.fgMuted}`}>Deleted</div>}
                        </div>
                    </div>
                ) : (
                    <ErrorBoundary variant="minimal">
                        <DiffEditor 
                           height="100%"
                           theme={themeClasses.monaco}
                           language="xml"
                           original={diffContent.original || ''}
                           modified={diffContent.modified || ''}
                           onMount={handleDiffMount}
                           options={{ 
                               readOnly: true, 
                               minimap: { enabled: false }, 
                               renderSideBySide: diffViewMode === 'split', 
                               fontSize: 13, 
                               wordWrap: 'on', 
                               scrollBeyondLastLine: false, 
                               padding: { top: 16 },
                               automaticLayout: true
                           }}
                        />
                    </ErrorBoundary>
                )}
             </div>
          </main>

          {showAi && (
            <div className="absolute inset-0 z-40 md:relative md:inset-auto md:w-[450px] md:border-l md:shrink-0 h-full shadow-2xl transition-all">
                <AIPanel 
                    onClose={() => toggleAiPanel(false)} 
                    context={{
                        mode: 'diff',
                        fileName: activePath || undefined,
                        diffOriginal: diffContent.original || undefined,
                        diffModified: diffContent.modified || undefined,
                        relatedFiles: modifiedPaths,
                        onLoadContext: handleFetchContext
                    }}
                    themeClasses={themeClasses}
                />
            </div>
          )}
       </div>
    </div>
  );
};

export default DiffView;