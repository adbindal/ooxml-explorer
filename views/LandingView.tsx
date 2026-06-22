import React, { useState } from 'react';
import { GitCompare, FileArchive, Beaker, Sun, Moon, UploadCloud, FilePlus } from 'lucide-react';
import { ThemeClasses } from '../types';
import Logo from '../components/Logo';
import { useAppStore } from '../store/appStore';
import { selectFileWithPicker } from '../utils/filePicker';

interface LandingViewProps {
  themeClasses: ThemeClasses;
}

const LandingView: React.FC<LandingViewProps> = ({ themeClasses }) => {
  const { theme, toggleTheme, setMode, loadEditorFile, setDiffFiles } = useAppStore();
  const [dragCounter, setDragCounter] = useState(0);
  const [dragFileCount, setDragFileCount] = useState(0);

  const handleEditorUpload = (file: File) => {
      loadEditorFile(file).catch(e => alert(e.message));
  };

  const handleFilesDrop = (files: FileList) => {
    console.log(`[Landing] Dropped ${files.length} files`);
    if (files.length === 1) {
        handleEditorUpload(files[0]);
    } else if (files.length >= 2) {
        setDiffFiles(files[0], files[1]);
        setMode('diff-setup'); 
    }
  };

  const handleDragEnter = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragCounter(prev => prev + 1);
    
    // Try to detect count early if possible (browser dependent)
    if (e.dataTransfer.items) {
        setDragFileCount(e.dataTransfer.items.length);
    }
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragCounter(prev => prev - 1);
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    // Continuous update for count as it might not be available in enter
    if (e.dataTransfer.items && e.dataTransfer.items.length !== dragFileCount) {
        setDragFileCount(e.dataTransfer.items.length);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragCounter(0);
    setDragFileCount(0);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        handleFilesDrop(e.dataTransfer.files);
    }
  };

  const isDragging = dragCounter > 0;

  return (
    <div 
        className={`min-h-screen w-full flex flex-col items-center justify-center p-4 md:p-6 relative overflow-hidden transition-colors duration-300 ${themeClasses.bg} ${themeClasses.fg}`}
        onDragEnter={handleDragEnter}
        onDragLeave={handleDragLeave}
        onDragOver={handleDragOver}
        onDrop={handleDrop}
    >
      {/* Drag Overlay */}
      {isDragging && (
        <div className="absolute inset-0 z-50 bg-[#4A89DC]/20 backdrop-blur-md m-4 rounded-3xl flex flex-col items-center justify-center pointer-events-none overflow-hidden">
            {/* Static Border */}
            <div className="absolute inset-0 border-4 border-[#4A89DC] border-dashed rounded-3xl opacity-50"></div>
            
            <div className={`
                relative z-10 p-12 rounded-3xl shadow-2xl flex flex-col items-center gap-6
                ${theme === 'dark' ? 'bg-[#0B1221] border border-[#1F3F70]' : 'bg-white border border-blue-100'}
            `}>
                {dragFileCount > 1 ? (
                    <>
                        <div className="flex items-center gap-4 text-[#4A89DC]">
                            <FileArchive size={56} className="opacity-50 scale-75" />
                            <GitCompare size={80} className="drop-shadow-lg" />
                            <FileArchive size={56} className="opacity-50 scale-75" />
                        </div>
                        <div className="text-center space-y-2">
                            <span className="block text-3xl font-bold text-[#4A89DC]">Comparison Mode</span>
                            <p className={`text-base ${theme === 'dark' ? 'text-blue-200' : 'text-blue-600'} opacity-80`}>
                                Release to compare <span className="font-bold">{dragFileCount}</span> files
                            </p>
                        </div>
                    </>
                ) : (
                    <>
                         <div className="relative text-[#4A89DC]">
                             <UploadCloud size={80} className="drop-shadow-lg" />
                             {dragFileCount === 1 && (
                                <div className="absolute -bottom-2 -right-2 bg-[#4A89DC] text-white p-1 rounded-full border-2 border-white dark:border-[#0B1221] shadow-sm">
                                    <FilePlus size={16} />
                                </div>
                             )}
                         </div>
                         <div className="text-center space-y-2">
                            <span className="block text-3xl font-bold text-[#4A89DC]">Inspect File</span>
                            <p className={`text-base ${theme === 'dark' ? 'text-blue-200' : 'text-blue-600'} opacity-80`}>
                                Release to open Editor
                            </p>
                        </div>
                    </>
                )}
            </div>
            
            {/* Helper Pill */}
            <div className={`
                mt-8 px-6 py-2 rounded-full text-sm font-medium border shadow-lg backdrop-blur-sm transition-colors
                ${theme === 'dark' ? 'bg-[#152238]/90 border-[#1F3F70] text-blue-300' : 'bg-white/90 border-blue-100 text-blue-600'}
            `}>
                {dragFileCount === 0 ? "Drop files to start" : dragFileCount > 1 ? "Diff Viewer" : "OOXML Editor"}
            </div>
        </div>
      )}

      <div className="absolute top-4 right-4 z-20">
        <button onClick={toggleTheme} className={`p-2 rounded-full ${themeClasses.hover} transition-colors`}>
            {theme === 'dark' ? <Sun size={20} /> : <Moon size={20} />}
        </button>
      </div>

      <div className="max-w-2xl w-full text-center space-y-8 z-10">
          <div className="flex flex-col items-center space-y-4">
            <h1 className={`text-5xl font-bold tracking-tight ${themeClasses.fg} flex items-center justify-center gap-5`}>
                <Logo size={64} theme={theme} className="shrink-0" />
                <span>OOXML <span className="text-[#4A89DC]">Explorer</span></span>
            </h1>
            <p className={`text-lg ${themeClasses.fgMuted}`}>Inspect, Edit, and Diff Office Open XML files</p>
            
            <div className={`inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-xs font-medium border ${themeClasses.border} bg-opacity-50 text-[#4A89DC]`}>
                <UploadCloud size={14} />
                <span>Drag 1 file to Edit • 2+ to Diff</span>
            </div>
          </div>
          
          <div className={`flex flex-col gap-6 transition-opacity duration-500 opacity-100`}>
            
            {/* Primary Action Card */}
            <div 
              onClick={async () => {
                  const file = await selectFileWithPicker(['.docx', '.xlsx', '.pptx']);
                  if (file) {
                      handleEditorUpload(file);
                  }
              }}
              className={`group relative border-2 border-dashed rounded-2xl p-10 transition-all cursor-pointer shadow-lg hover:shadow-xl ${themeClasses.card} ${themeClasses.hover} border-[#4A89DC]/30 hover:border-[#4A89DC]`}
            >
              <div className="flex flex-col items-center gap-6 pointer-events-none">
                <div className="w-20 h-20 rounded-full bg-[#4A89DC]/10 flex items-center justify-center text-[#4A89DC] group-hover:scale-110 transition-transform shadow-sm">
                    <FileArchive size={40} />
                </div>
                <div className="space-y-2">
                    <h3 className={`text-2xl font-bold ${themeClasses.fg}`}>Inspect & Edit</h3>
                    <p className={`text-base ${themeClasses.fgMuted}`}>Drag & drop or click to open .docx, .xlsx, .pptx</p>
                </div>
              </div>
            </div>
            <input
                type="file"
                accept=".docx,.xlsx,.pptx"
                style={{ display: 'none' }}
                onChange={(e) => {
                    const file = e.target.files?.[0];
                    if (file) {
                        handleEditorUpload(file);
                    }
                }}
            />

            {/* Secondary Action Card */}
            <div onClick={() => setMode('diff-setup')} className={`group border rounded-xl p-6 transition-all cursor-pointer flex items-center justify-center gap-4 hover:shadow-md ${themeClasses.card} ${themeClasses.hover}`}>
                <div className="w-12 h-12 rounded-full bg-[#4A89DC]/10 flex items-center justify-center text-[#4A89DC] group-hover:scale-110 transition-transform"><GitCompare size={24} /></div>
                <div className="text-left">
                    <h3 className={`font-medium ${themeClasses.fg}`}>Diff Files</h3>
                    <p className={`text-sm ${themeClasses.fgMuted}`}>Compare two separate archives</p>
                </div>
            </div>
            
          </div>
        </div>
        
        <div className="absolute bottom-4 left-0 right-0 text-center z-10">
            <button onClick={() => setMode('validator')} className={`text-sm flex items-center justify-center gap-2 mx-auto opacity-70 hover:opacity-100 transition-opacity ${themeClasses.fgMuted} hover:${themeClasses.fg}`}>
                <Beaker size={14} /> Go to Validator & Test Runner
            </button>
        </div>
    </div>
  );
};

export default LandingView;