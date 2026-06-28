import React, { useState } from 'react';
import { 
  FileCode, FileImage, File, ChevronRight, ChevronDown 
} from 'lucide-react';
import { FileNode, DiffNode } from '../types';
import { isXmlFile, isImageFile } from '../utils/xmlUtils';
import { isNodeVisible } from '../utils/treeUtils';

interface FileTreeProps {
  node: FileNode | DiffNode;
  level?: number;
  onSelect: (node: FileNode | DiffNode) => void;
  activePath: string | null;
  modifiedPaths?: Set<string>; // For Editor Mode
  isDiffMode?: boolean; // Changes rendering slightly
  pendingChanges?: Record<string, string>;
  showUnchanged?: boolean; // For Diff filtering
  searchQuery?: string;
}

const FileTree: React.FC<FileTreeProps> = ({ 
  node, 
  level = 0, 
  onSelect, 
  activePath, 
  modifiedPaths, 
  isDiffMode = false,
  pendingChanges,
  showUnchanged = true,
  searchQuery = ''
}) => {
  const diffNode = isDiffMode ? (node as DiffNode) : null;
  const status = diffNode?.status || 'unchanged';
  const hasChange = diffNode?.hasChange;

  // Use centralized logic for visibility
  const isVisible = isNodeVisible(node, searchQuery, isDiffMode, showUnchanged);

  const shouldStartOpen = true; 
  const [isOpen, setIsOpen] = useState(shouldStartOpen);

  // Adjust state when search or diff parameters change (No-Effect Pattern)
  const [prevSearchQuery, setPrevSearchQuery] = useState(searchQuery);
  const [prevDiffMode, setPrevDiffMode] = useState(isDiffMode);
  
  if (searchQuery !== prevSearchQuery || isDiffMode !== prevDiffMode) {
      setPrevSearchQuery(searchQuery);
      setPrevDiffMode(isDiffMode);
      if (searchQuery || (isDiffMode && !showUnchanged && hasChange)) {
          setIsOpen(true);
      }
  }

  if (!isVisible) return null;

  const hasPending = pendingChanges && pendingChanges[node.path];
  const isModified = modifiedPaths && modifiedPaths.has(node.path);
  const isActive = activePath === node.path;

  const getIcon = () => {
    if (node.isFolder) {
      return isOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />;
    }
    if (isXmlFile(node.name)) return <FileCode size={14} className={isDiffMode ? "opacity-70" : "text-[#4A89DC]"} />;
    if (isImageFile(node.name)) return <FileImage size={14} className={isDiffMode ? "opacity-70" : "text-[#1F3F70] opacity-70"} />;
    return <File size={14} className="opacity-70" />;
  };

  const getStatusClasses = () => {
    if (isDiffMode) {
      if (status === 'added') return 'text-green-500';
      if (status === 'deleted') return 'text-red-500 line-through decoration-red-500/50 opacity-70';
      if (status === 'modified') return 'text-amber-500';
      if (node.isFolder && hasChange) return 'text-inherit font-medium';
      return 'opacity-60'; 
    } else {
        if (isActive) return 'text-inherit font-medium';
        if (hasPending) return 'text-[#4A89DC] font-medium'; 
        return 'opacity-60';
    }
  };

  const handleClick = (e: React.MouseEvent) => {
    e.stopPropagation();
    if (node.isFolder) {
      setIsOpen(!isOpen);
    } else {
      onSelect(node);
    }
  };

  return (
    <div className="select-none">
      <div 
        className={`
          flex items-center gap-2 py-1 px-2 cursor-pointer text-sm transition-colors
          ${isActive ? 'bg-[#4A89DC]/10 border-l-2 border-[#4A89DC]' : 'hover:bg-[#4A89DC]/10 border-l-2 border-transparent'}
          ${getStatusClasses()}
        `}
        style={{ paddingLeft: `${level * 12 + 8}px` }}
        onClick={handleClick}
      >
        <span className="opacity-70 flex-shrink-0">{getIcon()}</span>
        <span className="truncate">{node.name}</span>
        
        {/* Indicators */}
        <div className="ml-auto flex items-center gap-1">
            {!isDiffMode && isModified && !hasPending && (
                <span className="text-[10px] font-bold opacity-60">M</span>
            )}
            {!isDiffMode && hasPending && (
                <div className="w-1.5 h-1.5 rounded-full bg-[#4A89DC]" title="Unsaved changes"></div>
            )}
            
            {/* Diff Indicators */}
            {isDiffMode && status !== 'unchanged' && !node.isFolder && (
                <span className={`
                    text-[10px] font-bold px-1.5 py-0.5 rounded ml-2 min-w-[20px] text-center
                    ${status === 'added' ? 'bg-green-500/10 text-green-600' : ''}
                    ${status === 'deleted' ? 'bg-red-500/10 text-red-600 no-underline' : ''}
                    ${status === 'modified' ? 'bg-amber-500/10 text-amber-600' : ''}
                `}>
                    {status === 'added' ? 'A' : status === 'deleted' ? 'D' : 'M'}
                </span>
            )}
        </div>
      </div>
      
      {node.isFolder && isOpen && node.children && (
        <div>
          {(Object.entries(node.children) as [string, FileNode][])
            // Sort folders first
            .sort(([, a], [, b]) => (a.isFolder === b.isFolder ? 0 : a.isFolder ? -1 : 1))
            .map(([key, child]) => (
            <FileTree 
              key={key} 
              node={child} 
              level={level + 1} 
              onSelect={onSelect} 
              activePath={activePath} 
              modifiedPaths={modifiedPaths}
              pendingChanges={pendingChanges}
              isDiffMode={isDiffMode}
              showUnchanged={showUnchanged}
              searchQuery={searchQuery}
            />
          ))}
        </div>
      )}
    </div>
  );
};

export default FileTree;