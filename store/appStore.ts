import { create, StateCreator } from 'zustand';
import { AppMode, EditorState, DiffState } from '../types';
import { loadZipFile, generateDiffTree } from '../services/zipService';

interface UiState {
  sidebarOpen: boolean;
  showAi: boolean;
  diffViewMode: 'split' | 'inline';
}

export interface AppStore {
  // --- Global State ---
  theme: 'dark' | 'light';
  mode: AppMode;
  ui: UiState;
  
  // --- Editor State ---
  editor: EditorState;
  
  // --- Diff State ---
  diff: DiffState & {
    loading: boolean;
  };

  // --- Actions ---
  toggleTheme: () => void;
  setMode: (mode: AppMode) => void;
  
  // UI Actions
  toggleSidebar: (force?: boolean) => void;
  toggleAiPanel: (force?: boolean) => void;
  setDiffViewMode: (mode: 'split' | 'inline') => void;

  // Editor Actions
  loadEditorFile: (file: File) => Promise<void>;
  updateEditorState: (update: Partial<EditorState> | ((prev: EditorState) => Partial<EditorState>)) => void;
  resetEditor: () => void;

  // Diff Actions
  setDiffFiles: (original: File | null, modified: File | null) => void;
  runDiffComparison: () => Promise<void>;
  updateDiffState: (update: Partial<DiffState>) => void;
}

const initialEditorState: EditorState = {
  zip: null,
  tree: null,
  fileName: '',
  activePath: null,
  openTabs: [],
  pendingChanges: {},
  modifiedPaths: new Set(),
  contentCache: {}
};

const initialDiffState: DiffState = {
  originalFile: null,
  modifiedFile: null,
  originalZip: null,
  modifiedZip: null,
  tree: null,
  activePath: null
};

// Export the creator for testing purposes
export const appStoreCreator: StateCreator<AppStore> = (set, get) => ({
  // --- Initial State ---
  theme: 'dark',
  mode: 'landing',
  ui: {
    sidebarOpen: true,
    showAi: false,
    diffViewMode: 'split'
  },
  editor: initialEditorState,
  diff: { ...initialDiffState, loading: false },

  // --- Implementation ---

  toggleTheme: () => set(state => {
    const newTheme = state.theme === 'dark' ? 'light' : 'dark';
    // Update DOM immediately for no-flash
    if (newTheme === 'light') document.documentElement.classList.add('light');
    else document.documentElement.classList.remove('light');
    return { theme: newTheme };
  }),

  setMode: (mode) => set(state => {
    // When returning to landing, reset complex states to ensure fresh start
    if (mode === 'landing') {
        return { 
            mode, 
            editor: initialEditorState,
            diff: { ...initialDiffState, loading: false }
        };
    }
    return { mode };
  }),

  toggleSidebar: (force) => set(state => ({ 
    ui: { ...state.ui, sidebarOpen: force ?? !state.ui.sidebarOpen } 
  })),

  toggleAiPanel: (force) => set(state => ({ 
    ui: { ...state.ui, showAi: force ?? !state.ui.showAi } 
  })),

  setDiffViewMode: (viewMode) => set(state => ({
    ui: { ...state.ui, diffViewMode: viewMode }
  })),

  // Editor Logic
  loadEditorFile: async (file) => {
    console.log(`[Store] Loading Editor File: ${file.name}`);
    try {
      const { zip, tree, flat } = await loadZipFile(file);
      set({
        mode: 'editor',
        editor: {
          ...initialEditorState,
          zip,
          tree,
          fileName: file.name
        },
        // Reset UI for fresh start, but keep theme
        ui: { ...get().ui, sidebarOpen: true, showAi: false }
      });
    } catch (e) {
      console.error("[Store] Load Failed:", e);
      throw e;
    }
  },

  updateEditorState: (update) => set(state => ({
    editor: typeof update === 'function' ? { ...state.editor, ...update(state.editor) } : { ...state.editor, ...update }
  })),

  resetEditor: () => set({ editor: initialEditorState }),

  // Diff Logic
  setDiffFiles: (original, modified) => set(state => ({
    diff: { ...state.diff, originalFile: original, modifiedFile: modified }
  })),

  runDiffComparison: async () => {
    const { diff } = get();
    if (!diff.originalFile || !diff.modifiedFile) return;

    console.log(`[Store] Running Comparison...`);
    set(state => ({ diff: { ...state.diff, loading: true } }));

    try {
      const [zipA, zipB] = await Promise.all([
          loadZipFile(diff.originalFile),
          loadZipFile(diff.modifiedFile)
      ]);
      
      const tree = generateDiffTree(zipA.flat, zipB.flat);
      
      set(state => ({
        diff: {
          ...state.diff,
          originalZip: zipA.zip,
          modifiedZip: zipB.zip,
          tree,
          loading: false
        }
      }));
    } catch (e) {
      console.error("[Store] Comparison Failed:", e);
      set(state => ({ diff: { ...state.diff, loading: false } }));
      throw e;
    }
  },

  updateDiffState: (update) => set(state => ({
    diff: { ...state.diff, ...update }
  }))
});

export const useAppStore = create<AppStore>(appStoreCreator);