import React, { useState, useEffect, useRef } from 'react';
import Editor, { OnMount } from '@monaco-editor/react';
import { FileNode, ThemeClasses } from '../types';
import FileTree from '../components/FileTree';
import AIPanel from '../components/AIPanel';
import ErrorBoundary from '../components/ErrorBoundary';
import Logo from '../components/Logo';
import { formatXml, minifyXml, isXmlFile, isImageFile, isBinaryFile } from '../utils/xmlUtils';
import { exportModifiedZip } from '../services/zipService';
import { isSaveHotkey, isSaveAllHotkey, isFindHotkey, isSidebarHotkey } from '../utils/hotkeyUtils';
import { EditorFileContext } from '../services/geminiService';
import { useAppStore } from '../store/appStore';
import { defineMonacoThemes } from '../utils/theme';
import { 
  Save, Download, Search, X, PanelLeftClose, PanelLeftOpen, 
  Sparkles, ArrowLeft, Image as ImageIcon, FileCode,
  ListChecks, Sun, Moon
} from 'lucide-react';

interface EditorViewProps {
  themeClasses: ThemeClasses;
}

const EditorView: React.FC<EditorViewProps> = ({ themeClasses }) => {
  const { 
      editor: editorState, 
      ui, 
      theme,
      updateEditorState, 
      setMode, 
      toggleSidebar, 
      toggleAiPanel, 
      toggleTheme 
  } = useAppStore();

  const [filter, setFilter] = useState('');
  const [imageSrc, setImageSrc] = useState<string | null>(null);

  const editorRef = useRef<Parameters<OnMount>[0] | null>(null);
  const monacoRef = useRef<Parameters<OnMount>[1] | null>(null);
  const isMounted = useRef(false);

  useEffect(() => {
    isMounted.current = true;
    console.log("[EditorView] Mounted. File:", editorState.fileName);
    if (window.innerWidth < 768) {
      toggleSidebar(false);
    }
    return () => { 
        isMounted.current = false; 
        if (editorRef.current) {
            try {
                const model = editorRef.current.getModel();
                if (model) model.dispose();
            } catch {
                // Ignore cleanup error
            }
        }
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Update Monaco theme dynamically when app theme changes
  useEffect(() => {
      if (monacoRef.current) {
          monacoRef.current.editor.setTheme(theme === 'dark' ? 'ooxml-dark' : 'ooxml-light');
      }
  }, [theme]);

  const { activePath, pendingChanges, contentCache, modifiedPaths, zip } = editorState;
  const { sidebarOpen, showAi } = ui;
  
  const pendingCount = Object.keys(pendingChanges).length;
  const currentContent = activePath ? (pendingChanges[activePath] ?? contentCache[activePath] ?? '') : '';
  const allPaths = zip ? Object.keys(zip.files).sort() : [];

  const handleFetchContext = async (paths: string[]): Promise<EditorFileContext[]> => {
      if (!zip) return [];
      const results = await Promise.all(paths.map(async (path) => {
          if (!isMounted.current) return null;
          if (isImageFile(path)) return null; 

          if (pendingChanges[path]) return { fileName: path, content: pendingChanges[path] };
          if (contentCache[path]) return { fileName: path, content: contentCache[path] };

          const text = await zip.file(path)?.async('string');
          if (text !== undefined) {
              const fmt = isXmlFile(path) ? formatXml(text) : text;
              return { fileName: path, content: fmt };
          }
          return null;
      }));
      return results.filter(Boolean) as EditorFileContext[];
  };

  useEffect(() => {
    let active = true; 
    if (!activePath || !zip) {
        if (imageSrc !== null) {
            // eslint-disable-next-line react-hooks/set-state-in-effect
            setImageSrc(null);
        }
        return;
    }
    if (isImageFile(activePath)) {
        zip.file(activePath)?.async('blob').then(blob => {
            if (active && isMounted.current) {
                setImageSrc(URL.createObjectURL(blob));
            }
        });
    } else {
        if (imageSrc !== null) {
            setImageSrc(null);
        }
    }
    return () => {
        active = false;
        if (imageSrc) URL.revokeObjectURL(imageSrc);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activePath, zip]);

  const handleSelect = async (node: FileNode) => {
    if (node.isFolder || !zip) return;
    
    if (window.innerWidth < 768) toggleSidebar(false);
    
    if (!editorState.openTabs.includes(node.path)) {
        updateEditorState(prev => ({ ...prev, openTabs: [...prev.openTabs, node.path] }));
    }

    if (!contentCache[node.path]) {
        // If it's a binary file, we don't load content as text and prevent editing
        if (isBinaryFile(node.path)) {
            updateEditorState({ activePath: node.path });
            return;
        }
        
        // Read the file as a string
        const text = await zip.file(node.path)?.async('string');
        if (!isMounted.current) return;
        
        // Format if XML, otherwise store raw text content
        const finalContent = isXmlFile(node.name) ? formatXml(text || '') : (text || '');
        
        updateEditorState(prev => ({
            ...prev,
            activePath: node.path,
            contentCache: { ...prev.contentCache, [node.path]: finalContent }
        }));
    } else {
        updateEditorState({ activePath: node.path });
    }
  };

  const handleEditorChange = (value: string | undefined) => {
    if (!activePath || value === undefined) return;
    updateEditorState(prev => ({
        ...prev,
        pendingChanges: { ...prev.pendingChanges, [activePath]: value }
    }));
  };

  const handleApply = () => {
    if (!activePath || !pendingChanges[activePath] || !zip) return;
    const contentToSave = pendingChanges[activePath] as string;
    const fileContent = isXmlFile(activePath) ? minifyXml(contentToSave) : contentToSave;
    
    zip.file(activePath, fileContent);

    updateEditorState(prev => {
        const newModified = new Set(prev.modifiedPaths);
        newModified.add(activePath);
        const newPending = { ...prev.pendingChanges };
        delete newPending[activePath];
        return { 
            ...prev, 
            modifiedPaths: newModified, 
            pendingChanges: newPending,
            contentCache: { ...prev.contentCache, [activePath]: contentToSave }
        };
    });
  };
  
  const handleApplyAll = () => {
    if (!zip) return;
    updateEditorState(prev => {
        const newModified = new Set(prev.modifiedPaths);
        const newPending = { ...prev.pendingChanges };
        const newCache = { ...prev.contentCache };

        Object.entries(prev.pendingChanges).forEach(([path, val]) => {
            const content = val as string;
            const fileContent = isXmlFile(path) ? minifyXml(content) : content;
            zip.file(path, fileContent);
            newModified.add(path);
            delete newPending[path];
            newCache[path] = content;
        });

        return { ...prev, modifiedPaths: newModified, pendingChanges: newPending, contentCache: newCache };
    });
  };

  const handleExport = async () => {
    if (!zip) return;
    const changesToExport: Record<string, string> = {};
    await exportModifiedZip(zip, changesToExport, editorState.fileName);
  };

  const handleBack = () => {
    if (pendingCount > 0) {
        const confirmLeave = window.confirm(`You have ${pendingCount} unsaved file(s). Are you sure you want to leave and discard your edits?`);
        if (!confirmLeave) return;
    }
    setMode('landing');
  };

  const closeTab = (e: React.MouseEvent, path: string) => {
    e.stopPropagation();
    
    // Warn if closing a file with unsaved/pending changes
    if (pendingChanges[path]) {
        const confirmClose = window.confirm(`"${path.split('/').pop()}" has unsaved changes. Are you sure you want to close it and discard your edits?`);
        if (!confirmClose) return;
    }

    updateEditorState(prev => {
        const newTabs = prev.openTabs.filter(t => t !== path);
        let newActive = prev.activePath;
        if (activePath === path) {
            newActive = newTabs.length > 0 ? newTabs[newTabs.length - 1] : null;
        }
        return { ...prev, openTabs: newTabs, activePath: newActive };
    });
  };
  
  // Shortcuts with extracted Logic
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Save All
      if (isSaveAllHotkey(e)) {
        e.preventDefault();
        handleApplyAll();
        return;
      }
      // Save
      if (isSaveHotkey(e)) {
        e.preventDefault();
        handleApply();
        return;
      }
      // Find (Delegate to Monaco Native Widget)
      if (isFindHotkey(e)) {
          e.preventDefault();
          if (editorRef.current) {
              editorRef.current.focus();
              editorRef.current.trigger('keyboard', 'actions.find', null);
          }
          return;
      }
      // Toggle Sidebar
      if (isSidebarHotkey(e)) {
          e.preventDefault();
          toggleSidebar();
          return;
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activePath, pendingChanges, sidebarOpen]);

  // Warn on browser-level exit (refresh/close) when there are unsaved changes
  useEffect(() => {
    const handleBeforeUnload = (e: BeforeUnloadEvent) => {
      if (pendingCount > 0) {
        e.preventDefault();
        e.returnValue = '';
      }
    };
    window.addEventListener('beforeunload', handleBeforeUnload);
    return () => window.removeEventListener('beforeunload', handleBeforeUnload);
  }, [pendingCount]);

  const handleEditorMount: OnMount = (editor, monaco) => {
      editorRef.current = editor;
      monacoRef.current = monaco;
      
      // Define shared themes
      defineMonacoThemes(monaco);

      // Set initial theme
      monaco.editor.setTheme(theme === 'dark' ? 'ooxml-dark' : 'ooxml-light');
  };

  return (
    <div className={`h-screen w-full flex flex-col transition-colors duration-300 ${themeClasses.bg} ${themeClasses.fg} overflow-hidden`}>
      <header className={`h-12 border-b flex items-center justify-between px-2 md:px-4 shrink-0 ${themeClasses.bgSec} ${themeClasses.border}`}>
        <div className="flex items-center gap-2 md:gap-4 overflow-hidden">
            <button onClick={handleBack} className={`transition-colors ${themeClasses.hoverText} shrink-0`}><ArrowLeft size={18} /></button>
            <button onClick={() => toggleSidebar()} className={`${themeClasses.icon} hover:${themeClasses.fg} mr-1 shrink-0`}>
                {sidebarOpen ? <PanelLeftClose size={18} /> : <PanelLeftOpen size={18} />}
            </button>
            <div className="flex items-center gap-2 font-semibold text-sm truncate">
                <Logo size={20} theme={theme} />
                <span className={`${themeClasses.fg} truncate`}>{editorState.fileName}</span>
            </div>
        </div>
        <div className="flex items-center gap-3 shrink-0 ml-2">
           {pendingCount > 0 && <span className={`text-xs mr-2 ${themeClasses.fgMuted} hidden md:inline`}>{pendingCount} unsaved files</span>}
           
            <div className={`flex items-center rounded border ${themeClasses.border} p-0.5 mr-2`}>
                <button 
                    onClick={handleApply} 
                    disabled={!activePath || pendingChanges[activePath] === undefined}
                    className={`p-1.5 rounded transition-colors ${!activePath || pendingChanges[activePath] === undefined ? themeClasses.fgMuted + ' opacity-50 cursor-not-allowed' : themeClasses.hover + ' text-[#4A89DC]'}`}
                >
                    <Save size={16} />
                </button>
                <div className={`w-px h-4 ${themeClasses.border} mx-0.5`}></div>
                 <button 
                    onClick={handleApplyAll} 
                    disabled={pendingCount === 0}
                    className={`p-1.5 rounded transition-colors flex items-center gap-1 relative ${pendingCount === 0 ? themeClasses.fgMuted + ' opacity-50 cursor-not-allowed' : themeClasses.hover + ' text-[#4A89DC]'}`}
                >
                    <ListChecks size={16} />
                    {pendingCount > 0 && <span className="absolute -top-1 -right-1 w-3 h-3 bg-[#4A89DC] text-[8px] text-white flex items-center justify-center rounded-full">{pendingCount}</span>}
                </button>
            </div>

           <button 
             onClick={handleExport}
             disabled={pendingCount > 0} 
             className={`flex items-center gap-2 text-xs font-medium px-3 py-1.5 rounded transition-colors ${pendingCount > 0 ? `${themeClasses.bgPanel} ${themeClasses.fgMuted} border ${themeClasses.border} cursor-not-allowed opacity-50` : 'bg-[#4A89DC] hover:bg-[#3b75c0] text-white'}`}
            >
             <Download size={14} /> <span className="hidden sm:inline">Export</span>
           </button>

           <div className={`w-px h-4 mx-1 ${themeClasses.border} hidden sm:block`}></div>

           <button 
               onClick={() => toggleAiPanel()} 
               title="AI Assistant"
               className={`p-2 rounded transition-colors ${showAi ? 'bg-[#4A89DC]/20 text-[#4A89DC]' : `${themeClasses.icon} ${themeClasses.hoverText}`}`} 
           >
               <Sparkles size={18} />
           </button>
           <button onClick={toggleTheme} className={`p-2 rounded-full ${themeClasses.hover} ${themeClasses.icon} ${themeClasses.hoverText} hidden sm:flex`}>{theme === 'dark' ? <Sun size={18} /> : <Moon size={18} />}</button>
        </div>
      </header>

      <div className="flex-1 flex overflow-hidden relative">
        {sidebarOpen && editorState.tree && (
            <aside className={`
                absolute md:relative z-30 h-full
                w-64 md:w-64 border-r flex flex-col shrink-0 shadow-2xl md:shadow-none
                ${themeClasses.bgSidebar} ${themeClasses.border}
            `}>
                <div className={`p-3 border-b ${themeClasses.border} relative`}>
                    <Search className={`absolute left-5 top-5 ${themeClasses.fgMuted}`} size={14} />
                    <input 
                        type="text" 
                        placeholder="Filter files..." 
                        value={filter}
                        onChange={(e) => setFilter(e.target.value)}
                        className={`w-full border rounded pl-8 pr-2 py-1.5 text-xs focus:outline-none focus:border-[#4A89DC] ${themeClasses.input}`}
                    />
                    <button className="md:hidden absolute right-2 top-3 text-[#1F3F70]" onClick={() => toggleSidebar(false)}>
                        <X size={14} />
                    </button>
                </div>
                <div className="flex-1 overflow-y-auto py-2">
                    <FileTree 
                        node={editorState.tree} 
                        onSelect={handleSelect} 
                        activePath={activePath} 
                        modifiedPaths={modifiedPaths}
                        pendingChanges={pendingChanges}
                        searchQuery={filter} 
                    />
                </div>
            </aside>
        )}

        {sidebarOpen && (
             <div className="md:hidden absolute inset-0 bg-black/50 z-20" onClick={() => toggleSidebar(false)}></div>
        )}

        <main className={`flex-1 flex flex-col overflow-hidden relative min-w-0 ${themeClasses.bg}`}>
            {editorState.openTabs.length > 0 ? (
                <div className={`flex border-b overflow-x-auto scrollbar-hide h-9 shrink-0 ${themeClasses.bgSec} ${themeClasses.border}`}>
                    {editorState.openTabs.map(path => {
                        const name = path.split('/').pop();
                        const isActive = activePath === path;
                        const isDirty = pendingChanges[path];
                        return (
                            <div 
                                key={path} 
                                onClick={() => handleSelect({ path, name: name!, isFolder: false, children: {} })}
                                className={`
                                    flex items-center gap-2 px-3 min-w-[100px] md:min-w-[120px] max-w-[150px] md:max-w-[200px] text-xs cursor-pointer border-r select-none
                                    ${themeClasses.border}
                                    ${isActive ? themeClasses.activeTab : themeClasses.inactiveTab}
                                `}
                            >
                                <span className="truncate flex-1">{name}</span>
                                {isDirty && <div className="w-1.5 h-1.5 rounded-full bg-[#4A89DC] shrink-0"></div>}
                                <button onClick={(e) => closeTab(e, path)} className={`opacity-100 md:opacity-0 md:hover:opacity-100 group-hover:opacity-100 p-0.5 rounded ${themeClasses.hover}`}>
                                    <X size={12} />
                                </button>
                            </div>
                        )
                    })}
                </div>
            ) : <div className={`h-9 border-b ${themeClasses.bgSec} ${themeClasses.border}`}></div>}

            <div className="flex-1 relative">
                {!activePath ? (
                    <div className={`absolute inset-0 flex flex-col items-center justify-center ${themeClasses.fgMuted} p-4 text-center`}>
                        <FileCode size={48} className="mb-4 opacity-20" />
                        <p>Select a file from the sidebar to edit</p>
                    </div>
                ) : imageSrc ? (
                    <div className={`absolute inset-0 flex flex-col items-center justify-center p-4 md:p-8 ${themeClasses.bgPanel} bg-opacity-50`}>
                        <img src={imageSrc} className={`max-w-full max-h-full object-contain shadow-lg border rounded ${themeClasses.border}`} alt="Preview" />
                        <div className={`mt-4 flex items-center gap-2 text-xs ${themeClasses.fgMuted}`}>
                            <ImageIcon size={14} /> Image Preview (Read Only)
                        </div>
                    </div>
                ) : isBinaryFile(activePath) ? (
                    <div className={`absolute inset-0 flex flex-col items-center justify-center p-4 md:p-8 ${themeClasses.bgPanel} bg-opacity-50 text-center`}>
                        <FileCode size={48} className="mb-4 text-[#EAB308]/80 opacity-80 animate-pulse" />
                        <h3 className="font-bold text-sm mb-2">Binary File Format</h3>
                        <p className={`max-w-sm text-xs leading-relaxed ${themeClasses.fgMuted}`}>
                            This file format contains binary encoded content (e.g. OLE embeddings, font data, or activeX modules). 
                            Visual previews and text editing are disabled to protect package integrity and prevent data corruption.
                        </p>
                    </div>
                ) : (
                    <ErrorBoundary variant="minimal">
                        <Editor 
                            height="100%" 
                            theme={themeClasses.monaco}
                            defaultLanguage="xml"
                            path={activePath} 
                            value={currentContent}
                            onChange={handleEditorChange}
                            onMount={handleEditorMount}
                            options={{
                                minimap: { enabled: false },
                                fontSize: 13,
                                wordWrap: 'on',
                                scrollBeyondLastLine: false,
                                padding: { top: 16 }
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
                        mode: 'editor',
                        fileName: activePath || undefined,
                        content: currentContent,
                        relatedFiles: allPaths,
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

export default EditorView;