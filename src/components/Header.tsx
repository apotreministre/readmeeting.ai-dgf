import React, { useState, useRef, useEffect } from 'react';
import {
  Shield,
  Sparkles,
  Building2,
  User,
  LogOut,
  CheckCircle2,
  AlertTriangle,
  Key,
  Bell,
  ShieldAlert,
  ExternalLink,
  Copy,
  Check,
  ChevronRight
} from 'lucide-react';
import { DgfUserProfile, MsalRuntimeConfig } from '../types/auth';
import {
  AdminNotificationConfig,
  ADMIN_AUTH_STATUS_CONFIG,
  OFFICIAL_ADMIN_REQUEST_TEMPLATE
} from '../types/adminNotification';

interface HeaderProps {
  activeTab: 'dashboard' | 'auth_diagnostic' | 'reports' | 'architecture';
  setActiveTab: (tab: 'dashboard' | 'auth_diagnostic' | 'reports' | 'architecture') => void;
  user: DgfUserProfile | null;
  config: MsalRuntimeConfig;
  onLoginClick: () => void;
  onLogoutClick: () => void;
  isLoggingIn: boolean;
  adminConfig?: AdminNotificationConfig;
}

export const Header: React.FC<HeaderProps> = ({
  activeTab,
  setActiveTab,
  user,
  config,
  onLoginClick,
  onLogoutClick,
  isLoggingIn,
  adminConfig,
}) => {
  const [isNotifOpen, setIsNotifOpen] = useState<boolean>(false);
  const [copiedInHeader, setCopiedInHeader] = useState<boolean>(false);
  const notifRef = useRef<HTMLDivElement>(null);

  // Fermeture du menu déroulant en cas de clic extérieur
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (notifRef.current && !notifRef.current.contains(event.target as Node)) {
        setIsNotifOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleCopyRequest = async (e: React.MouseEvent) => {
    e.stopPropagation();
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(OFFICIAL_ADMIN_REQUEST_TEMPLATE);
      }
      setCopiedInHeader(true);
      setTimeout(() => setCopiedInHeader(false), 3000);
    } catch (err) {
      console.warn('[Header copy warning]', err);
    }
  };

  const statusMeta = adminConfig
    ? ADMIN_AUTH_STATUS_CONFIG[adminConfig.status] || ADMIN_AUTH_STATUS_CONFIG.pending
    : ADMIN_AUTH_STATUS_CONFIG.pending;

  const hasPendingAction = adminConfig ? adminConfig.status !== 'fully_configured' : true;
  return (
    <header className="border-b border-slate-800 bg-slate-900/95 sticky top-0 z-40 backdrop-blur-md">
      {/* Top Banner : République de Côte d'Ivoire & DGF */}
      <div className="border-b border-slate-800/80 bg-slate-950/80 px-4 py-1.5 text-xs">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3 text-slate-400">
            <span className="flex items-center gap-1.5 font-medium text-amber-400">
              <span className="w-2 h-2 rounded-full bg-amber-400 animate-pulse"></span>
              RÉPUBLIQUE DE CÔTE D'IVOIRE
            </span>
            <span className="text-slate-600">•</span>
            <span className="text-slate-300 font-medium">Ministère des Finances et du Budget</span>
            <span className="text-slate-600">•</span>
            <span className="text-slate-400">Direction Générale des Financements (DGF)</span>
          </div>

          <div className="flex items-center gap-2">
            <span className="text-slate-500 font-mono text-[11px]">Union - Discipline - Travail</span>
          </div>
        </div>
      </div>

      {/* Main Navigation Bar */}
      <div className="max-w-7xl mx-auto px-4 sm:px-6 py-3 flex items-center justify-between">
        {/* Logo & Product Name */}
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-lg bg-gradient-to-br from-emerald-600 to-teal-800 flex items-center justify-center shadow-lg shadow-emerald-950/40 border border-emerald-500/30">
            <Building2 className="w-5 h-5 text-emerald-200" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <span className="text-lg font-bold text-white tracking-tight">ReadMeeting<span className="text-emerald-400">.ai</span></span>
              <span className="px-2 py-0.5 text-[10px] font-semibold bg-emerald-950/80 text-emerald-300 border border-emerald-800/60 rounded">
                DGF
              </span>
            </div>
            <p className="text-xs text-slate-400">Assistant Institutionnel des Réunions M365</p>
          </div>
        </div>

        {/* Tab Navigation */}
        <nav className="hidden md:flex items-center gap-1 bg-slate-950/80 p-1 rounded-lg border border-slate-800">
          <button
            onClick={() => setActiveTab('dashboard')}
            className={`px-3 py-1.5 text-xs font-medium rounded-md transition-all ${
              activeTab === 'dashboard'
                ? 'bg-emerald-600 text-white shadow-sm'
                : 'text-slate-400 hover:text-white hover:bg-slate-900'
            }`}
          >
            Tableau de Bord
          </button>
          <button
            onClick={() => setActiveTab('auth_diagnostic')}
            className={`px-3 py-1.5 text-xs font-medium rounded-md transition-all flex items-center gap-1.5 ${
              activeTab === 'auth_diagnostic'
                ? 'bg-emerald-600 text-white shadow-sm'
                : 'text-slate-400 hover:text-white hover:bg-slate-900'
            }`}
          >
            <Shield className="w-3.5 h-3.5" />
            Diagnostic Entra ID
            {!config.isConfigured && (
              <span className="w-2 h-2 rounded-full bg-amber-400"></span>
            )}
          </button>
          <button
            onClick={() => setActiveTab('reports')}
            className={`px-3 py-1.5 text-xs font-medium rounded-md transition-all ${
              activeTab === 'reports'
                ? 'bg-emerald-600 text-white shadow-sm'
                : 'text-slate-400 hover:text-white hover:bg-slate-900'
            }`}
          >
            Bibliothèque des CR
          </button>
          <button
            onClick={() => setActiveTab('architecture')}
            className={`px-3 py-1.5 text-xs font-medium rounded-md transition-all ${
              activeTab === 'architecture'
                ? 'bg-emerald-600 text-white shadow-sm'
                : 'text-slate-400 hover:text-white hover:bg-slate-900'
            }`}
          >
            Architecture & Azure
          </button>
        </nav>

        {/* User Account & Notification Center */}
        <div className="flex items-center gap-2 sm:gap-3">
          {/* Notification Center Trigger & Dropdown */}
          <div className="relative" ref={notifRef}>
            <button
              onClick={() => setIsNotifOpen(!isNotifOpen)}
              className="relative p-2 text-slate-300 hover:text-white bg-slate-950 hover:bg-slate-800 rounded-lg transition-colors border border-slate-800 shadow-sm"
              title="Centre de notifications institutionnel DGF"
            >
              <Bell className="w-4 h-4" />
              {hasPendingAction && (
                <span className="absolute -top-1 -right-1 w-4 h-4 bg-amber-500 text-slate-950 font-extrabold text-[10px] rounded-full flex items-center justify-center animate-pulse shadow-sm">
                  1
                </span>
              )}
            </button>

            {/* Persistent Notification Dropdown Panel */}
            {isNotifOpen && (
              <div className="absolute right-0 mt-2 w-80 sm:w-96 bg-slate-900 border border-slate-800 rounded-2xl shadow-2xl z-50 p-4 space-y-3.5 backdrop-blur-xl">
                <div className="flex items-center justify-between border-b border-slate-800 pb-2.5">
                  <div className="flex items-center gap-2">
                    <Bell className="w-4 h-4 text-emerald-400" />
                    <span className="text-xs font-bold text-white uppercase tracking-wider">
                      Centre de Notifications DGF
                    </span>
                  </div>
                  <span className={`text-[10px] font-semibold px-2 py-0.5 rounded-full border ${hasPendingAction ? 'bg-amber-950 text-amber-300 border-amber-800' : 'bg-emerald-950 text-emerald-300 border-emerald-800'}`}>
                    {hasPendingAction ? '1 action requise' : 'Système à jour'}
                  </span>
                </div>

                {/* Persistent Item : Autorisations Microsoft 365 */}
                <div className="bg-slate-950 p-3.5 rounded-xl border border-slate-800 space-y-2.5">
                  <div className="flex items-start justify-between gap-2">
                    <div className="flex items-center gap-2">
                      <ShieldAlert className="w-4 h-4 text-amber-400 shrink-0" />
                      <span className="text-xs font-bold text-white leading-tight">
                        Action requise — Autorisations Microsoft 365
                      </span>
                    </div>
                  </div>

                  <div className="flex items-center gap-2">
                    <span className={`px-2 py-0.5 text-[10px] font-semibold rounded border ${statusMeta.badgeClass}`}>
                      {statusMeta.label}
                    </span>
                  </div>

                  <p className="text-[11px] text-slate-300 leading-relaxed">
                    L'intervention de l'administrateur est requise pour ajouter la permission déléguée <code className="text-emerald-400 font-mono">OnlineMeetingTranscript.Read.All</code> et activer l'accès API dans le Teams Admin Center.
                  </p>

                  <div className="grid grid-cols-2 gap-2 pt-1">
                    <a
                      href="https://entra.microsoft.com/"
                      target="_blank"
                      rel="noopener noreferrer"
                      className="px-2.5 py-1.5 text-[11px] font-semibold text-white bg-blue-600 hover:bg-blue-500 rounded-lg flex items-center justify-center gap-1.5 transition-colors shadow-sm"
                    >
                      <ExternalLink className="w-3 h-3" />
                      Entra ID
                    </a>
                    <a
                      href="https://admin.teams.microsoft.com/"
                      target="_blank"
                      rel="noopener noreferrer"
                      className="px-2.5 py-1.5 text-[11px] font-semibold text-white bg-indigo-600 hover:bg-indigo-500 rounded-lg flex items-center justify-center gap-1.5 transition-colors shadow-sm"
                    >
                      <ExternalLink className="w-3 h-3" />
                      Teams Admin
                    </a>
                  </div>

                  <div className="pt-2 border-t border-slate-900 flex items-center justify-between text-[11px]">
                    <button
                      onClick={handleCopyRequest}
                      className="text-slate-400 hover:text-white flex items-center gap-1 transition-colors"
                      title="Copier le message officiel destiné à l'administrateur DGF"
                    >
                      {copiedInHeader ? (
                        <>
                          <Check className="w-3 h-3 text-emerald-400" />
                          <span className="text-emerald-300 font-semibold">Copié !</span>
                        </>
                      ) : (
                        <>
                          <Copy className="w-3 h-3 text-slate-400" />
                          <span>Copier la demande</span>
                        </>
                      )}
                    </button>

                    <button
                      onClick={() => {
                        setActiveTab('dashboard');
                        setIsNotifOpen(false);
                      }}
                      className="text-emerald-400 hover:text-emerald-300 font-semibold flex items-center gap-1 transition-colors"
                    >
                      <span>Voir sur le tableau de bord</span>
                      <ChevronRight className="w-3 h-3" />
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>

          {user ? (
            <div className="flex items-center gap-3 bg-slate-950 border border-slate-800 rounded-lg p-1.5 pr-3">
              <div className="w-8 h-8 rounded-md bg-emerald-900/50 border border-emerald-700/50 flex items-center justify-center text-emerald-300 font-semibold text-xs">
                {user.displayName.slice(0, 2).toUpperCase()}
              </div>
              <div className="text-left hidden sm:block">
                <div className="text-xs font-medium text-slate-200 leading-tight flex items-center gap-1">
                  <span>{user.displayName}</span>
                  <CheckCircle2 className="w-3 h-3 text-emerald-400" />
                </div>
                <div className="text-[11px] text-slate-400 truncate max-w-[160px]">{user.mail || user.userPrincipalName}</div>
              </div>
              <button
                onClick={onLogoutClick}
                title="Se déconnecter de Microsoft Entra ID"
                className="p-1.5 text-slate-400 hover:text-rose-400 hover:bg-slate-900 rounded transition-colors ml-1"
              >
                <LogOut className="w-4 h-4" />
              </button>
            </div>
          ) : (
            <div className="flex items-center gap-2">
              <button
                onClick={onLoginClick}
                disabled={isLoggingIn}
                className="flex items-center gap-2 px-3 py-2 text-xs font-medium rounded-lg bg-blue-600 hover:bg-blue-500 text-white shadow-md shadow-blue-950/50 transition-all border border-blue-400/20 disabled:opacity-50"
              >
                <Key className="w-3.5 h-3.5" />
                {isLoggingIn ? 'Connexion en cours...' : 'Connexion Entra ID'}
              </button>
            </div>
          )}
        </div>
      </div>
    </header>
  );
};
