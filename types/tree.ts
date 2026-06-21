import type { JSZipObject } from 'jszip';

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
