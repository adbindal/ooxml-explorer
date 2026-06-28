import React, { Component, ErrorInfo } from 'react';
import { AlertTriangle, RefreshCw, ShieldAlert, Terminal, Copy, Check, Power, Activity } from 'lucide-react';
import { getLogString, setDebugMode, getDebugMode } from '../services/debugService';

interface Props {
  children?: React.ReactNode;
  variant?: 'full' | 'minimal';
}

interface State {
  hasError: boolean;
  error: Error | null;
  autoRecovered: boolean;
  showLogs: boolean;
  copied: boolean;
  debugMode: boolean;
}

class ErrorBoundary extends Component<Props, State> {
  public state: State = {
      hasError: false,
      error: null,
      autoRecovered: false,
      showLogs: false,
      copied: false,
      debugMode: getDebugMode()
  };

  // Explicitly declare inherited React.Component properties to satisfy strict type-checking
  public props!: Props & { children?: React.ReactNode };
  public setState!: (
    state: Partial<State> | ((prevState: Readonly<State>, props: Readonly<Props>) => Partial<State> | null),
    callback?: () => void
  ) => void;

  static getDerivedStateFromError(error: Error): Partial<State> {
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    const isScriptError = error.message && error.message.toLowerCase().includes('script error');

    // Attempt auto-recovery once for script errors
    if (isScriptError && !this.state.autoRecovered) {
        console.warn("Auto-recovering from Script Error...");
        setTimeout(() => {
            try {
                if (this.state.hasError) {
                    this.setState({ hasError: false, error: null, autoRecovered: true });
                }
            } catch {
                // Ignore unmount errors
            }
        }, 50);
        return;
    }

    if (!isScriptError) {
        console.error("Uncaught error caught by Boundary:", error, errorInfo);
    }
  }

  handleRetry = () => {
    this.setState({ hasError: false, error: null, autoRecovered: false, showLogs: false });
  };

  handleReload = () => {
    try {
        window.location.reload();
    } catch (e) {
        console.error("Reload failed, attempting path reassignment", e);
        window.location.href = window.location.pathname + window.location.search + window.location.hash;
    }
  };

  handleCopyLogs = async () => {
      const logs = getLogString();
      try {
          await navigator.clipboard.writeText(logs);
          this.setState({ copied: true });
          setTimeout(() => this.setState({ copied: false }), 2000);
      } catch (e) {
          console.error("Failed to copy logs", e);
      }
  };

  toggleDebugMode = () => {
    const newState = !this.state.debugMode;
    setDebugMode(newState);
    this.setState({ debugMode: newState });
    if (newState) {
        if (confirm("Logging enabled. Reload page to capture startup logs?")) {
            window.location.reload();
        }
    }
  };

  render() {
    if (this.state.hasError) {
      const errorMsg = this.state.error?.message || '';
      const isScriptError = errorMsg.toLowerCase().includes('script error');
      const { variant = 'full' } = this.props;

      // Dynamic Theme Detection
      const isLight = typeof document !== 'undefined' && document.documentElement.classList.contains('light');

      const colors = isLight ? {
          bg: 'bg-[#F8FAFC]',
          card: 'bg-white',
          text: 'text-[#1F3F70]',
          textMuted: 'text-slate-500',
          border: 'border-slate-200',
          shadow: 'shadow-xl shadow-blue-900/5',
          btnPrimary: 'bg-[#4A89DC] hover:bg-[#3b75c0] text-white',
          btnSecondary: 'bg-white border border-slate-300 hover:bg-slate-50 text-slate-700',
          code: 'bg-slate-50 border-slate-200 text-red-600',
          iconScript: 'text-amber-500',
          iconError: 'text-red-500',
          debugContainer: 'bg-[#0F172A] border-slate-800',
          debugText: 'text-slate-300',
      } : {
          bg: 'bg-[#0B1221]',
          card: 'bg-[#152238]',
          text: 'text-[#E2E8F0]',
          textMuted: 'text-[#94A3B8]',
          border: 'border-[#1F3F70]',
          shadow: 'shadow-2xl shadow-black/50',
          btnPrimary: 'bg-[#4A89DC] hover:bg-[#3b75c0] text-white',
          btnSecondary: 'bg-[#1F3F70]/30 hover:bg-[#1F3F70]/50 text-[#E2E8F0]',
          code: 'bg-[#0B1221] border-[#1F3F70] text-red-400',
          iconScript: 'text-amber-500',
          iconError: 'text-red-500',
          debugContainer: 'bg-black border-[#1F3F70]',
          debugText: 'text-green-400',
      };

      // Minimal UI for nested components
      if (variant === 'minimal') {
          return (
            <div className={`h-full w-full flex flex-col items-center justify-center p-4 rounded text-center border ${colors.card} ${colors.border} ${colors.textMuted}`}>
                <AlertTriangle className={`mb-2 ${isScriptError ? colors.iconScript : colors.iconError}`} size={24} />
                <p className="text-xs font-medium mb-2">Component Error</p>
                <button 
                    type="button"
                    onClick={this.handleRetry}
                    className={`text-xs px-3 py-1.5 rounded flex items-center gap-1 transition-colors ${colors.btnSecondary}`}
                >
                    <RefreshCw size={12} /> Reload
                </button>
            </div>
          );
      }
      
      // Full Screen UI
      return (
        <div className={`h-screen w-full flex flex-col items-center justify-center p-6 overflow-y-auto transition-colors ${colors.bg} ${colors.text}`}>
          <div className={`max-w-xl w-full border rounded-xl p-6 ${colors.card} ${colors.border} ${colors.shadow}`}>
            <div className={`flex items-center gap-3 mb-4 ${isScriptError ? colors.iconScript : colors.iconError}`}>
              {isScriptError ? <ShieldAlert size={32} /> : <AlertTriangle size={32} />}
              <h2 className="text-xl font-bold">{isScriptError ? "External Script Issue" : "Something went wrong"}</h2>
            </div>
            
            <p className={`text-sm mb-6 leading-relaxed ${colors.textMuted}`}>
              {isScriptError 
                ? "An external resource (like a browser extension or CDN script) triggered a security error. This is usually harmless."
                : "The application encountered an unexpected error. Please try reloading."
              }
            </p>

            {/* Error Message Box */}
            {this.state.error && !isScriptError && (
              <div className={`p-3 rounded border mb-6 overflow-auto max-h-48 ${colors.code}`}>
                 <code className="text-xs font-mono break-words whitespace-pre-wrap">{this.state.error.stack || this.state.error.message}</code>
              </div>
            )}
            
            <div className="flex gap-3 mb-6">
                <button
                type="button"
                onClick={this.handleRetry}
                className={`flex-1 py-2 rounded font-medium flex items-center justify-center gap-2 transition-colors ${colors.btnPrimary}`}
                >
                <RefreshCw size={16} /> {isScriptError ? "Resume App" : "Try Again"}
                </button>
                <button
                type="button"
                onClick={this.handleReload}
                className={`flex-1 py-2 rounded font-medium flex items-center justify-center gap-2 transition-colors ${colors.btnSecondary}`}
                >
                <RefreshCw size={16} /> Reload Page
                </button>
            </div>

            {/* Debug Details Section */}
            <div className={`border-t pt-4 ${colors.border}`}>
                <button 
                    type="button"
                    onClick={() => this.setState(s => ({ showLogs: !s.showLogs }))}
                    className={`flex items-center gap-2 text-xs transition-colors mb-3 w-full ${colors.textMuted} hover:${colors.text}`}
                >
                    <Terminal size={14} />
                    {this.state.showLogs ? "Hide Debug Details" : "Show Debug Details"}
                </button>

                {this.state.showLogs && (
                    <div className="animate-in fade-in slide-in-from-top-2 space-y-3">
                        {/* Control Bar */}
                        <div className={`flex justify-between items-center p-2 rounded border ${colors.debugContainer} ${colors.border}`}>
                             <div className="flex items-center gap-2">
                                <span className={`flex items-center gap-1 text-[10px] font-bold px-1.5 py-0.5 rounded ${this.state.debugMode ? 'bg-red-500/20 text-red-500' : 'bg-gray-500/20 text-gray-500'}`}>
                                    <Activity size={10} />
                                    {this.state.debugMode ? 'REC' : 'OFF'}
                                </span>
                             </div>
                             <div className="flex gap-2">
                                <button
                                    type="button"
                                    onClick={this.toggleDebugMode}
                                    className={`text-[10px] flex items-center gap-1 px-2 py-1 rounded transition-colors ${this.state.debugMode ? 'bg-red-900/30 text-red-400 hover:bg-red-900/50' : 'bg-gray-700 text-gray-400 hover:bg-gray-600'}`}
                                >
                                    <Power size={10} /> {this.state.debugMode ? 'Disable Log' : 'Enable Log'}
                                </button>
                                <button 
                                    type="button"
                                    onClick={this.handleCopyLogs}
                                    className={`text-[10px] flex items-center gap-1 px-2 py-1 rounded transition-colors ${this.state.copied ? 'bg-green-500/20 text-green-500' : 'bg-gray-700 hover:bg-gray-600 text-gray-300'}`}
                                >
                                    {this.state.copied ? <Check size={10} /> : <Copy size={10} />}
                                    {this.state.copied ? "Copied" : "Copy"}
                                </button>
                             </div>
                        </div>

                        <div className={`border rounded p-3 h-64 overflow-y-auto font-mono text-[10px] whitespace-pre-wrap ${colors.debugContainer} ${colors.debugText} ${colors.border}`}>
                            {getLogString()}
                        </div>
                        <p className={`text-[10px] ${colors.textMuted}`}>
                            * Capture Mode is {this.state.debugMode ? "Active. Logs are being recorded." : "Inactive. Enable above to capture new logs."}
                        </p>
                    </div>
                )}
            </div>
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}

export default ErrorBoundary;