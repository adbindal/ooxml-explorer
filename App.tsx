import React, { useEffect } from 'react';
import LandingView from './views/LandingView';
import EditorView from './views/EditorView';
import DiffView from './views/DiffView';
import ValidatorView from './views/ValidatorView';
import ErrorBoundary from './components/ErrorBoundary';
import { useAppStore } from './store/appStore';
import { useThemeClasses } from './utils/theme';

function AppContent() {
  const { mode, theme } = useAppStore();
  const themeClasses = useThemeClasses(theme);

  useEffect(() => {
    console.log(`[App] Mode switched to: ${mode}`);
  }, [mode]);

  return (
    <div className={`font-sans ${themeClasses.bg} ${themeClasses.fg} min-h-screen`}>
      {mode === 'landing' && <LandingView themeClasses={themeClasses} />}
      {mode === 'editor' && <EditorView themeClasses={themeClasses} />}
      {(mode === 'diff-setup' || mode === 'diff-view') && <DiffView themeClasses={themeClasses} />}
      {mode === 'validator' && <ValidatorView themeClasses={themeClasses} />}
    </div>
  );
}

export default function App() {
  return (
    <ErrorBoundary>
      <AppContent />
    </ErrorBoundary>
  );
}