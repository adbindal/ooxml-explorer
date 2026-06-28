import type JSZip from 'jszip';
import type { DiffNode } from './tree';

export interface DiffState {
  originalFile: File | null;
  modifiedFile: File | null;
  originalZip: JSZip | null;
  modifiedZip: JSZip | null;
  tree: DiffNode | null;
  activePath: string | null;
}
