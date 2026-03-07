export type AppMode = 'landing' | 'editor' | 'diff-setup' | 'diff-view' | 'validator';

export interface FileNode {
  name: string;
  path: string; // Full path
  isFolder: boolean;
  children: Record<string, FileNode>;
  zipEntry?: any; // Avoiding JSZip import to prevent runtime module loading issues in types
  status?: 'added' | 'deleted' | 'modified' | 'unchanged'; // For diff mode
}

export interface DiffNode extends FileNode {
  hasChange: boolean; // For folder highlighting
}

export interface TreeStructure {
  root: FileNode;
  flat: Record<string, any>;
}

export interface EditorState {
  zip: any | null; // JSZip instance
  tree: FileNode | null;
  fileName: string;
  activePath: string | null;
  openTabs: string[];
  pendingChanges: Record<string, string>; // path -> new content
  modifiedPaths: Set<string>;
  contentCache: Record<string, string>; // path -> original content
}

export interface DiffState {
  originalFile: File | null;
  modifiedFile: File | null;
  originalZip: any | null;
  modifiedZip: any | null;
  tree: DiffNode | null;
  activePath: string | null;
}

export interface ThemeClasses {
  bg: string;
  bgSec: string;
  bgSidebar: string;
  bgPanel: string;
  fg: string;
  fgMuted: string;
  border: string;
  hover: string;
  hoverText: string;
  activeTree: string;
  activeTab: string;
  inactiveTab: string;
  card: string;
  input: string;
  icon: string;
  monaco: string;
}