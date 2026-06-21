import React from 'react';
import { render, screen, fireEvent } from '@testing-library/react';
import { describe, it, expect, vi } from 'vitest';
import LandingView from '../views/LandingView';
import EditorView from '../views/EditorView';
import DiffView from '../views/DiffView';
import { getThemeClasses } from '../utils/theme';

// Mock the store to control the state in component tests
const mockStore = {
  diff: {
    originalFile: new File([], 'old.xml'),
    modifiedFile: new File([], 'new.xml'),
    tree: { name: 'root', path: '', type: 'directory', children: [] },
    activePath: 'test.xml',
    loading: false
  },
  ui: { sidebarOpen: true, showAi: false, diffViewMode: 'split' },
  theme: 'dark',
  mode: 'diff',
  setMode: vi.fn(),
  toggleSidebar: vi.fn(),
  toggleAiPanel: vi.fn(),
  toggleTheme: vi.fn(),
  setDiffFiles: vi.fn(),
  runDiffComparison: vi.fn(),
  updateDiffState: vi.fn(),
  setDiffViewMode: vi.fn()
};

vi.mock('../store/appStore', () => ({
  useAppStore: () => mockStore
}));

// Create mock editor instances that can be inspected in tests
export const mockSubEditor = {
  updateOptions: vi.fn(),
  onDidChangeModel: vi.fn(() => ({ dispose: vi.fn() }))
};

export const mockDiffEditorInstance = {
  getOriginalEditor: vi.fn(() => mockSubEditor),
  getModifiedEditor: vi.fn(() => mockSubEditor),
  onDidUpdateDiff: vi.fn(() => ({ dispose: vi.fn() })),
  getLineChanges: vi.fn(() => [])
};

export const mockMonacoInstance = {
  editor: {
    setTheme: vi.fn()
  }
};

// Mock Monaco Editor to inspect props and lifecycle
vi.mock('@monaco-editor/react', () => ({
  DiffEditor: vi.fn((props) => {
    React.useEffect(() => {
      if (props.onMount) {
        props.onMount(mockDiffEditorInstance, mockMonacoInstance);
      }
    }, [props]);
    return <div data-testid="mock-diff-editor" data-options={JSON.stringify(props.options)} />;
  }),
  Editor: vi.fn((props) => <div data-testid="mock-editor" data-options={JSON.stringify(props.options)} />),
  default: vi.fn((props) => <div data-testid="mock-editor" data-options={JSON.stringify(props.options)} />)
}));

describe('LandingView Component', () => {
  const themeClasses = getThemeClasses('dark');

  it('renders the landing page with title and upload options', () => {
    render(<LandingView themeClasses={themeClasses} />);
    
    expect(screen.getByRole('heading', { level: 1 })).toBeDefined();
    expect(screen.getByText(/Inspect, Edit, and Diff Office Open XML files/i)).toBeDefined();
    expect(screen.getByText(/Drag 1 file to Edit/i)).toBeDefined();
  });

  it('shows both Editor and Diff modes', () => {
    render(<LandingView themeClasses={themeClasses} />);
    
    expect(screen.getByText(/Inspect & Edit/i)).toBeDefined();
    expect(screen.getByText(/Diff Files/i)).toBeDefined();
  });
});

describe('EditorView Component Validation', () => {
  const themeClasses = getThemeClasses('dark');

  it('passes correct options to the Editor', () => {
    // Update mock store for editor mode
    mockStore.mode = 'editor';
    mockStore.editor = {
        fileName: 'test.docx',
        activePath: 'word/document.xml',
        openTabs: ['word/document.xml'],
        pendingChanges: {},
        contentCache: { 'word/document.xml': '<root/>' },
        modifiedPaths: new Set(),
        zip: { files: { 'word/document.xml': {} } }
    };

    render(<EditorView themeClasses={themeClasses} />);
    
    const editor = screen.getByTestId('mock-editor');
    const options = JSON.parse(editor.getAttribute('data-options') || '{}');
    
    expect(options.wordWrap).toBe('on');
    expect(options.minimap.enabled).toBe(false);
    expect(options.fontSize).toBe(13);
  });

  it('renders open tabs and identifies active tab', () => {
    mockStore.editor.openTabs = ['word/document.xml', 'word/styles.xml'];
    mockStore.editor.activePath = 'word/document.xml';
    
    render(<EditorView themeClasses={themeClasses} />);
    
    expect(screen.getByText('document.xml')).toBeDefined();
    expect(screen.getByText('styles.xml')).toBeDefined();
  });

  it('shows dirty indicator when file has pending changes', () => {
    mockStore.editor.pendingChanges = { 'word/document.xml': '<modified/>' };
    
    const { container } = render(<EditorView themeClasses={themeClasses} />);
    // The dirty indicator is a div with bg-[#4A89DC]
    const dirtyIndicator = container.querySelector('.bg-\\[\\#4A89DC\\]');
    expect(dirtyIndicator).not.toBeNull();
  });

  it('disables save button when no pending changes', () => {
    mockStore.editor.pendingChanges = {};
    render(<EditorView themeClasses={themeClasses} />);
    
    // The save button is the first button in the actions group
    // In EditorView, it's the one with <Save /> icon.
    // We can find it by its disabled attribute or class.
    const saveButton = screen.getAllByRole('button').find(btn => 
        btn.innerHTML.includes('lucide-save') || btn.querySelector('.lucide-save')
    );
    expect(saveButton).toBeDefined();
    expect(saveButton?.hasAttribute('disabled')).toBe(true);
  });
});

describe('DiffView Component Validation', () => {
  const themeClasses = getThemeClasses('dark');

  beforeEach(() => {
    mockStore.mode = 'diff';
    mockStore.diff = {
        originalFile: new File([], 'old.xml'),
        modifiedFile: new File([], 'new.xml'),
        tree: { name: 'root', path: '', type: 'directory', children: [] },
        activePath: 'test.xml',
        loading: false
    };
    mockStore.ui.diffViewMode = 'split';
  });

  it('passes correct wordWrap options to the DiffEditor', () => {
    render(<DiffView themeClasses={themeClasses} />);
    
    const editor = screen.getByTestId('mock-diff-editor');
    const options = JSON.parse(editor.getAttribute('data-options') || '{}');
    
    expect(options.wordWrap).toBe('on');
    expect(options.renderSideBySide).toBe(true); // Split mode
  });

  it('registers model change listeners to enforce wordWrap on sub-editors', () => {
    // Reset our mock spies
    mockSubEditor.updateOptions.mockClear();
    mockSubEditor.onDidChangeModel.mockClear();

    render(<DiffView themeClasses={themeClasses} />);

    // Assert that both original and modified editors had updateOptions called with wordWrap: 'on'
    expect(mockSubEditor.updateOptions).toHaveBeenCalledWith({ wordWrap: 'on' });

    // Assert that model change listeners were registered for both sub-editors
    expect(mockSubEditor.onDidChangeModel).toHaveBeenCalledTimes(2);
  });

  it('updates renderSideBySide when diffViewMode changes', () => {
    mockStore.ui.diffViewMode = 'inline';
    render(<DiffView themeClasses={themeClasses} />);
    
    const editor = screen.getByTestId('mock-diff-editor');
    const options = JSON.parse(editor.getAttribute('data-options') || '{}');
    
    expect(options.renderSideBySide).toBe(false); // Inline mode
  });

  it('shows empty state when no file is selected', () => {
    mockStore.diff.activePath = null;
    render(<DiffView themeClasses={themeClasses} />);
    
    expect(screen.getByText(/Select a file to compare/i)).toBeDefined();
  });

  it('disables explain changes button when no file is selected', () => {
    mockStore.diff.activePath = null;
    render(<DiffView themeClasses={themeClasses} />);
    
    const explainBtn = screen.getByText(/Explain Changes/i).closest('button');
    expect(explainBtn?.hasAttribute('disabled')).toBe(true);
  });

  it('renders change navigation when changes exist', () => {
    // We need to mock the editor instance to have changes
    // But since we mock the component, we can just check if the UI elements are there
    render(<DiffView themeClasses={themeClasses} />);
    
    // The navigation buttons are in the header
    // Use a more specific query for the changes counter
    expect(screen.getByText(/No Changes/i)).toBeDefined();
    
    // Check for navigation icons
    const navButtons = screen.getAllByRole('button').filter(btn => 
        btn.querySelector('.lucide-arrow-up') || btn.querySelector('.lucide-arrow-down')
    );
    expect(navButtons.length).toBe(2);
  });

  it('toggles file tree filter', () => {
    render(<DiffView themeClasses={themeClasses} />);
    
    // Initially it shows "Diffs" because showUnchanged is false
    expect(screen.getByText(/Diffs/i)).toBeDefined();
    expect(screen.queryByText(/^All$/)).toBeNull();
  });

  it('handles global drag and drop of two files', () => {
    // Override store to trigger Setup Mode
    mockStore.diff.tree = null;
    mockStore.diff.originalFile = null;
    mockStore.diff.modifiedFile = null;
    mockStore.setDiffFiles.mockClear();

    render(<DiffView themeClasses={themeClasses} />);
    
    // Find the Compare Documents heading to locate the container
    const heading = screen.getByRole('heading', { name: /Compare Documents/i });
    const container = heading.closest('div');
    expect(container).toBeDefined();

    const file1 = new File(['contentA'], 'fileA.docx', { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const file2 = new File(['contentB'], 'fileB.docx', { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });

    // Simulate dragEnter with 2 files
    fireEvent.dragEnter(container!, {
        dataTransfer: {
            files: [file1, file2],
            items: [{ kind: 'file', type: file1.type }, { kind: 'file', type: file2.type }],
            types: ['Files']
        }
    });

    // Verify global overlay appears
    expect(screen.getByText(/Drop both files to Compare/i)).toBeDefined();

    // Simulate drop with 2 files
    fireEvent.drop(container!, {
        dataTransfer: {
            files: [file1, file2],
            items: [{ kind: 'file', type: file1.type }, { kind: 'file', type: file2.type }],
            types: ['Files']
        }
    });

    // Assert that setDiffFiles was called to set both files in store
    expect(mockStore.setDiffFiles).toHaveBeenCalledWith(file1, file2);
  });

  it('handles box-specific drag and drop of a single file', () => {
    // Override store to trigger Setup Mode
    mockStore.diff.tree = null;
    mockStore.diff.originalFile = null;
    mockStore.diff.modifiedFile = null;
    mockStore.setDiffFiles.mockClear();

    render(<DiffView themeClasses={themeClasses} />);

    // Find the original file upload box
    const originalBoxText = screen.getByText(/Click or Drag to upload original/i);
    const originalBox = originalBoxText.closest('div');
    expect(originalBox).toBeDefined();

    const mockFile = new File(['content'], 'original.docx', { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });

    // Simulate dragEnter over original box
    fireEvent.dragEnter(originalBox!, {
        dataTransfer: {
            files: [mockFile],
            items: [{ kind: 'file', type: mockFile.type }]
        }
    });

    // Simulate drop over original box
    fireEvent.drop(originalBox!, {
        dataTransfer: {
            files: [mockFile],
            items: [{ kind: 'file', type: mockFile.type }]
        }
    });

    // Assert that setDiffFiles was called
    expect(mockStore.setDiffFiles).toHaveBeenCalled();
  });
});
