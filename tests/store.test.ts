import { describe, it, expect, beforeEach, vi } from '../services/browserTestRunner';
import { appStoreCreator } from '../store/appStore';
import { create } from 'zustand';
import JSZip from 'jszip';

// Create a local, isolated store for testing so we don't affect the live UI
const useTestStore = create(appStoreCreator);

describe('App Store', () => {
  const initialState = useTestStore.getState();

  beforeEach(() => {
    useTestStore.setState(initialState, true);
    vi.clearAllMocks();
  });

  describe('Global UI Slice', () => {
    it('toggles theme between light and dark modes', () => {
      const { toggleTheme } = useTestStore.getState();
      expect(useTestStore.getState().theme).toBe('dark');
      toggleTheme();
      expect(useTestStore.getState().theme).toBe('light');
      toggleTheme();
      expect(useTestStore.getState().theme).toBe('dark');
    });

    it('toggles sidebar visibility state', () => {
      const { toggleSidebar } = useTestStore.getState();
      expect(useTestStore.getState().ui.sidebarOpen).toBe(true);
      toggleSidebar();
      expect(useTestStore.getState().ui.sidebarOpen).toBe(false);
      toggleSidebar(true);
      expect(useTestStore.getState().ui.sidebarOpen).toBe(true);
    });

    it('toggles AI panel visibility state', () => {
      const { toggleAiPanel } = useTestStore.getState();
      expect(useTestStore.getState().ui.showAi).toBe(false);
      toggleAiPanel();
      expect(useTestStore.getState().ui.showAi).toBe(true);
    });

    it('resets editor and diff state when switching mode to "landing"', () => {
      useTestStore.getState().updateEditorState({ fileName: 'dirty.docx' });
      const f1 = new File([''], 'a.docx');
      useTestStore.getState().setDiffFiles(f1, null);
      
      expect(useTestStore.getState().editor.fileName).toBe('dirty.docx');
      expect(useTestStore.getState().diff.originalFile).toBe(f1);

      useTestStore.getState().setMode('landing');
      
      expect(useTestStore.getState().mode).toBe('landing');
      expect(useTestStore.getState().editor.fileName).toBe('');
      expect(useTestStore.getState().diff.originalFile).toBeNull();
    });
    
    it('setDiffViewMode updates view mode', () => {
      const { setDiffViewMode } = useTestStore.getState();
      expect(useTestStore.getState().ui.diffViewMode).toBe('split');
      setDiffViewMode('inline');
      expect(useTestStore.getState().ui.diffViewMode).toBe('inline');
    });

    it('updateDiffState updates arbitrary diff state', () => {
        useTestStore.getState().updateDiffState({ activePath: 'test.xml' });
        expect(useTestStore.getState().diff.activePath).toBe('test.xml');
    });

    it('loadEditorFile handles errors gracefully', async () => {
        // Mock console.error to suppress output
        const spy = vi.spyOn(console, 'error').mockImplementation(() => {});
        
        const badFile = new File(['not a zip'], 'bad.docx');
        // We expect loadZipFile to fail (it's not mocked here but JSZip will fail)
        
        await expect(useTestStore.getState().loadEditorFile(badFile)).rejects.toThrow();
        
        spy.mockRestore();
    });

    });

  describe('Editor Slice', () => {
    it('loadEditorFile correctly populates editor state from zip', async () => {
      const zip = new JSZip();
      zip.file('[Content_Types].xml', '<root/>');
      const content = await zip.generateAsync({ type: 'blob' });
      const mockFile = new File([content], 'test.docx');

      await useTestStore.getState().loadEditorFile(mockFile);

      const state = useTestStore.getState();
      expect(state.mode).toBe('editor');
      expect(state.editor.fileName).toBe('test.docx');
      expect(state.editor.zip).toBeDefined();
    });

    it('updateEditorState correctly merges partial state updates', () => {
      useTestStore.getState().updateEditorState({ activePath: 'word/document.xml' });
      expect(useTestStore.getState().editor.activePath).toBe('word/document.xml');

      useTestStore.getState().updateEditorState(prev => ({ 
          openTabs: [...prev.openTabs, 'new.xml'] 
      }));
      expect(useTestStore.getState().editor.openTabs).toContain('new.xml');
    });

    it('resetEditor clears editor state', () => {
      useTestStore.getState().updateEditorState({ fileName: 'dirty.docx' });
      useTestStore.getState().resetEditor();
      expect(useTestStore.getState().editor.fileName).toBe('');
    });
  });

  describe('Diff Slice', () => {
    it('setDiffFiles updates file references', () => {
      const f1 = new File([''], 'a.docx');
      const f2 = new File([''], 'b.docx');
      useTestStore.getState().setDiffFiles(f1, f2);
      
      const state = useTestStore.getState().diff;
      expect(state.originalFile).toBe(f1);
      expect(state.modifiedFile).toBe(f2);
    });

    it('runDiffComparison generates tree when given valid files', async () => {
      const z1 = new JSZip(); z1.file('[Content_Types].xml', '');
      const z2 = new JSZip(); z2.file('[Content_Types].xml', '');
      const b1 = await z1.generateAsync({type:'blob'});
      const b2 = await z2.generateAsync({type:'blob'});
      const f1 = new File([b1], 'a.docx');
      const f2 = new File([b2], 'b.docx');

      useTestStore.getState().setDiffFiles(f1, f2);
      
      const promise = useTestStore.getState().runDiffComparison();
      expect(useTestStore.getState().diff.loading).toBe(true);
      
      await promise;
      
      const state = useTestStore.getState().diff;
      expect(state.loading).toBe(false);
      expect(state.tree).toBeDefined();
    });

    it('catches errors during diff comparison and resets loading state', async () => {
      // Temporarily silence console.error for expected failure
      const originalError = console.error;
      console.error = () => {};

      const f1 = new File(['invalid'], 'a.docx');
      const f2 = new File(['invalid'], 'b.docx');
      useTestStore.getState().setDiffFiles(f1, f2);
      
      try {
        await useTestStore.getState().runDiffComparison();
      } catch (e) {
         // Expected
      } finally {
         console.error = originalError; // Restore even if assertion fails
      }
      
      expect(useTestStore.getState().diff.loading).toBe(false);
    });
  });
});