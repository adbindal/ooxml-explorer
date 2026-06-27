import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import { initDebugService } from './services/debugService';
import { initStorageService } from './services/storageService';

// Initialize logging, global error handling, and local IndexedDB database
initDebugService();
initStorageService();

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