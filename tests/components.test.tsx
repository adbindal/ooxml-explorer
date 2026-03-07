import React from 'react';
import { render, screen } from '@testing-library/react';
import { describe, it, expect, vi } from 'vitest';
import LandingView from '../views/LandingView';
import { getThemeClasses } from '../utils/theme';

// Mock the store to control the state in component tests
vi.mock('../store/appStore', () => ({
  useAppStore: () => ({
    loadEditorFile: vi.fn(),
    setMode: vi.fn(),
    setDiffFiles: vi.fn()
  })
}));

describe('LandingView Component', () => {
  const themeClasses = getThemeClasses('dark');

  it('renders the landing page with title and upload options', () => {
    render(<LandingView themeClasses={themeClasses} />);
    
    // Use more specific queries to avoid multiple matches (like SVG titles)
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
