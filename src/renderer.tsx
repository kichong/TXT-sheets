import React from 'react';
import ReactDOM from 'react-dom/client';
import './renderer/browser-mock';
import { App } from './renderer/App';
import './renderer/styles.css';

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode><App /></React.StrictMode>,
);
