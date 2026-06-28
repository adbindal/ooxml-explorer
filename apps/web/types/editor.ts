import type JSZip from 'jszip';
import type { FileNode } from './tree';

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
