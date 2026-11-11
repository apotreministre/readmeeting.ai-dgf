import React, { useState, useEffect } from 'react';
import { createRoot } from 'react-dom/client';
import { MsalProvider } from '@azure/msal-react';
import { IPublicClientApplication } from '@azure/msal-browser';
import App from './App.tsx';
import { getInitializedMsalInstance } from './config/msalConfig.ts';
import './index.css';

/**
 * Composant racine assurant l'initialisation asynchrone garantie de MSAL
 * avant le montage de MsalProvider et de l'arbre des composants.
 */
function Root() {
  const [msalInstance, setMsalInstance] = useState<IPublicClientApplication | null>(null);
  const [initError, setInitError] = useState<string | null>(null);

  useEffect(() => {
    let isMounted = true;

    getInitializedMsalInstance()
      .then((instance) => {
        if (isMounted) {
          setMsalInstance(instance);
        }
      })
      .catch((err) => {
        console.error('[ReadMeeting.ai MSAL Bootstrap Error]', err);
        if (isMounted) {
          setInitError(err instanceof Error ? err.message : String(err));
        }
      });

    return () => {
      isMounted = false;
    };
  }, []);

  if (initError) {
    return (
      <div className="min-h-screen bg-slate-950 flex items-center justify-center p-6 text-slate-200">
        <div className="max-w-md bg-rose-950/60 border border-rose-800 rounded-xl p-6 space-y-3 shadow-xl">
          <div className="flex items-center gap-2 text-rose-300 font-bold text-sm">
            <span className="w-2.5 h-2.5 rounded-full bg-rose-500"></span>
            Erreur d'initialisation Microsoft Entra ID
          </div>
          <p className="text-xs text-rose-200 leading-relaxed font-mono">
            {initError}
          </p>
          <button
            onClick={() => window.location.reload()}
            className="px-3 py-1.5 text-xs bg-rose-900 hover:bg-rose-800 text-white rounded transition-colors"
          >
            Réessayer
          </button>
        </div>
      </div>
    );
  }

  if (!msalInstance) {
    return (
      <div className="min-h-screen bg-slate-950 flex items-center justify-center text-slate-300">
        <div className="flex flex-col items-center gap-3">
          <div className="w-8 h-8 border-2 border-emerald-500 border-t-transparent rounded-full animate-spin"></div>
          <p className="text-xs font-mono text-slate-400">
            Initialisation du socle de sécurité Microsoft Entra ID...
          </p>
        </div>
      </div>
    );
  }

  return (
    <MsalProvider instance={msalInstance}>
      <App onInstanceReplaced={(newInstance) => setMsalInstance(newInstance)} />
    </MsalProvider>
  );
}

const rootElement = document.getElementById('root');
if (rootElement) {
  createRoot(rootElement).render(<Root />);
}
