import type JSZip from 'jszip';
import type { JSZipObject } from 'jszip';

export type AppMode = 'landing' | 'editor' | 'diff-setup' | 'diff-view' | 'validator';

export interface FileNode {
  name: string;
  path: string; // Full path
  isFolder: boolean;
  children: Record<string, FileNode>;
  zipEntry?: JSZipObject; // Strictly typed ZIP entry
  status?: 'added' | 'deleted' | 'modified' | 'unchanged'; // For diff mode
}

export interface DiffNode extends FileNode {
  hasChange: boolean; // For folder highlighting
}

export interface TreeStructure {
  root: FileNode;
  flat: Record<string, JSZipObject>; // Strictly typed flat structure
}

export interface EditorState {
  zip: JSZip | null; // Strictly typed JSZip instance
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
  originalZip: JSZip | null;
  modifiedZip: JSZip | null;
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