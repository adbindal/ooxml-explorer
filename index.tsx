import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import { initDebugService } from './services/debugService';

// Initialize logging and global error handling
initDebugService();

const rootElement = document.getElementById('root');
if (!rootElement) {
  throw new Error("Could not find root element to mount to");
}

const root = ReactDOM.createRoot(rootElement);
root.render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);