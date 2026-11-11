import React, { useState } from 'react';
import {
  ShieldAlert,
  ExternalLink,
  Copy,
  Check,
  CheckCircle2,
  Clock,
  Lock,
  ChevronDown,
  ChevronUp,
  Settings,
  HelpCircle,
  FileText,
  AlertTriangle,
  RefreshCw,
  Sparkles,
  Building2,
  Share2
} from 'lucide-react';
import {
  AdminAuthStatus,
  ADMIN_AUTH_STATUS_CONFIG,
  OFFICIAL_ADMIN_REQUEST_TEMPLATE,
  AdminNotificationConfig
} from '../types/adminNotification';
import { DgfUserProfile } from '../types/auth';
import {
  AdminConsentWorkflowState,
  ADMIN_CONSENT_WORKFLOW_META,
} from '../types/adminConsentWorkflow';
import {
  NativeConsentExecutionResult,
  ConsentVerificationResult,
} from '../services/adminConsentService';
import { AdminConsentWorkflowSection } from './AdminConsentWorkflowSection';

interface AdminPermissionNotificationBannerProps {
  user: DgfUserProfile | null;
  adminConfig: AdminNotificationConfig;
  onUpdateAdminConfig: (newConfig: Partial<AdminNotificationConfig>) => void;
  onRunTechnicalCheck?: () => Promise<{ success: boolean; message: string; details?: string }>;
  isCheckingTechnical?: boolean;
  workflowState?: AdminConsentWorkflowState;
  onRequestConsent?: (justification?: string) => Promise<NativeConsentExecutionResult>;
  onVerifyConsent?: () => Promise<ConsentVerificationResult>;
  onUpdateWorkflowState?: (partial: Partial<AdminConsentWorkflowState>) => void;
  isRequestingConsent?: boolean;
  isVerifyingConsent?: boolean;
}

export const AdminPermissionNotificationBanner: React.FC<AdminPermissionNotificationBannerProps> = ({
  user,
  adminConfig,
  onUpdateAdminConfig,
  onRunTechnicalCheck,
  isCheckingTechnical = false,
  workflowState,
  onRequestConsent,
  onVerifyConsent,
  onUpdateWorkflowState,
  isRequestingConsent = false,
  isVerifyingConsent = false,
}) => {
  const [isCollapsed, setIsCollapsed] = useState<boolean>(Boolean(adminConfig.isBannerCollapsed));
  const [copied, setCopied] = useState<boolean>(false);
  const [showStatusModal, setShowStatusModal] = useState<boolean>(false);
  const [selectedNewStatus, setSelectedNewStatus] = useState<AdminAuthStatus>(adminConfig.status);
  const [statusChangeNotes, setStatusChangeNotes] = useState<string>('');
  const [techCheckResult, setTechCheckResult] = useState<{ success: boolean; message: string; details?: string } | null>(null);

  const isConsentGranted =
    workflowState?.status === 'consent_granted' ||
    adminConfig.status === 'entra_configured' ||
    adminConfig.status === 'fully_configured';
  const effectiveStatusKey: AdminAuthStatus =
    isConsentGranted && adminConfig.status === 'pending'
      ? 'entra_configured'
      : adminConfig.status;
  const currentStatusMeta = ADMIN_AUTH_STATUS_CONFIG[effectiveStatusKey] || ADMIN_AUTH_STATUS_CONFIG.pending;

  const handleCopyRequest = async () => {
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(OFFICIAL_ADMIN_REQUEST_TEMPLATE);
      } else {
        // Fallback standard
        const textarea = document.createElement('textarea');
        textarea.value = OFFICIAL_ADMIN_REQUEST_TEMPLATE;
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand('copy');
        document.body.removeChild(textarea);
      }
      setCopied(true);
      setTimeout(() => setCopied(false), 3000);
    } catch (err) {
      console.warn('[Copy to clipboard warning]', err);
    }
  };

  const handleToggleCollapse = () => {
    const nextVal = !isCollapsed;
    setIsCollapsed(nextVal);
    onUpdateAdminConfig({ isBannerCollapsed: nextVal });
  };

  const handleSaveStatus = () => {
    onUpdateAdminConfig({
      status: selectedNewStatus,
      lastUpdated: new Date().toISOString(),
      updatedByUpn: user?.userPrincipalName || user?.mail || 'Administrateur DGF',
      notes: statusChangeNotes.trim() || undefined,
    });
    setShowStatusModal(false);
  };

  const handleExecuteTechnicalCheck = async () => {
    if (!onRunTechnicalCheck) return;
    setTechCheckResult(null);
    try {
      const res = await onRunTechnicalCheck();
      setTechCheckResult(res);
      if (!res.success && adminConfig.status === 'fully_configured') {
        // Si le test échoue alors que le statut était déclaré complet, repositionner en attente de vérification
        onUpdateAdminConfig({
          status: 'verification_pending',
          lastUpdated: new Date().toISOString(),
          notes: `Vérification technique échouée : ${res.message}`,
        });
      }
    } catch (err: any) {
      setTechCheckResult({
        success: false,
        message: 'Erreur lors du test technique',
        details: err.message || String(err),
      });
    }
  };

  return (
    <div className={`bg-gradient-to-r from-slate-900 ${isConsentGranted ? 'via-emerald-950/20 border-emerald-500/30' : 'via-blue-950/20 border-amber-500/30'} to-slate-900 border rounded-xl shadow-lg transition-all overflow-hidden`}>
      {/* 1. Header de la notification */}
      <div className="p-4 sm:p-5 flex flex-col md:flex-row items-start md:items-center justify-between gap-4 border-b border-slate-800/80 bg-slate-950/40">
        <div className="flex items-start gap-3.5 flex-1">
          <div className={`w-10 h-10 rounded-xl ${isConsentGranted ? 'bg-emerald-950/80 border-emerald-600/40 text-emerald-400' : 'bg-amber-950/80 border-amber-600/40 text-amber-400'} border flex items-center justify-center shrink-0 shadow-sm mt-0.5`}>
            {isConsentGranted ? <CheckCircle2 className="w-5 h-5 text-emerald-400" /> : <ShieldAlert className="w-5 h-5 text-amber-400" />}
          </div>
          <div className="space-y-1 flex-1">
            <div className="flex flex-wrap items-center gap-2.5">
              <h3 className="text-sm sm:text-base font-bold text-white tracking-tight flex items-center gap-2">
                {isConsentGranted ? 'Autorisations Microsoft 365 — Consentement Entra ID Accordé' : 'Action requise — Autorisations Microsoft 365'}
              </h3>
              <span className={`px-2.5 py-0.5 text-[10px] font-semibold rounded-md border flex items-center gap-1.5 shadow-sm ${currentStatusMeta.badgeClass}`}>
                <span className="w-1.5 h-1.5 rounded-full bg-current"></span>
                {currentStatusMeta.label}
              </span>
            </div>
            <p className="text-xs text-slate-300 leading-relaxed max-w-4xl">
              {isConsentGranted ? (
                <>Le consentement administrateur pour l'application <strong>ReadMeeting.ai | DGF</strong> est accordé sur le locataire. Les requêtes de récupération des transcriptions Microsoft Teams via Microsoft Graph sont opérationnelles.</>
              ) : (
                <>La configuration de <strong>ReadMeeting.ai | DGF</strong> nécessite l'activation de certaines autorisations Microsoft 365 pour permettre la récupération des transcriptions des réunions Microsoft Teams. Nous sollicitons l'intervention de l'administrateur Microsoft 365 afin de finaliser les configurations suivantes.</>
              )}
            </p>
          </div>
        </div>

        {/* Boutons d'état et de réduction */}
        <div className="flex items-center gap-2 shrink-0 self-end md:self-center">
          <button
            onClick={() => {
              setSelectedNewStatus(adminConfig.status);
              setStatusChangeNotes(adminConfig.notes || '');
              setShowStatusModal(true);
            }}
            className="px-2.5 py-1.5 text-xs font-medium text-slate-300 hover:text-white bg-slate-900 hover:bg-slate-800 border border-slate-800 rounded-lg flex items-center gap-1.5 transition-colors"
            title="Modifier le statut de suivi administratif"
          >
            <Settings className="w-3.5 h-3.5 text-slate-400" />
            <span className="hidden sm:inline">Statut :</span>
            <span className="font-semibold text-amber-300">{currentStatusMeta.label.split(' ')[0]}...</span>
          </button>

          <button
            onClick={handleToggleCollapse}
            className="p-1.5 text-slate-400 hover:text-white hover:bg-slate-800 rounded-lg transition-colors"
            title={isCollapsed ? 'Afficher les détails de la notification' : 'Réduire la notification'}
          >
            {isCollapsed ? <ChevronDown className="w-4 h-4" /> : <ChevronUp className="w-4 h-4" />}
          </button>
        </div>
      </div>

      {/* 2. Corps de la notification */}
      {!isCollapsed && (
        <div className="p-4 sm:p-5 space-y-4">
          {/* Module Natif Microsoft Entra ID : Demande de consentement administrateur */}
          {workflowState && onRequestConsent && onVerifyConsent && onUpdateWorkflowState && (
            <AdminConsentWorkflowSection
              user={user}
              workflowState={workflowState}
              onRequestConsent={onRequestConsent}
              onVerifyConsent={onVerifyConsent}
              onUpdateWorkflowState={onUpdateWorkflowState}
              isRequesting={isRequestingConsent}
              isVerifying={isVerifyingConsent}
            />
          )}

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
            {/* Volet A : Microsoft Entra ID */}
            <div className="bg-slate-950/80 border border-slate-800 rounded-xl p-4 space-y-3 flex flex-col justify-between">
              <div className="space-y-2.5">
                <div className="flex items-center justify-between border-b border-slate-800/80 pb-2">
                  <div className="flex items-center gap-2">
                    <span className="w-6 h-6 rounded-md bg-blue-950 text-blue-400 font-bold text-xs flex items-center justify-center border border-blue-800/60">
                      A
                    </span>
                    <h4 className="text-xs font-bold text-white uppercase tracking-wider">
                      Microsoft Entra ID
                    </h4>
                  </div>
                  <span className="text-[10px] font-mono text-blue-400 bg-blue-950/60 px-2 py-0.5 rounded border border-blue-800/40">
                    Inscription d'App
                  </span>
                </div>

                <div className="space-y-1.5 text-xs text-slate-300">
                  <div className="flex justify-between py-0.5 border-b border-slate-900 text-[11px]">
                    <span className="text-slate-500">Application :</span>
                    <span className="font-semibold text-white font-mono">ReadMeeting.ai | DGF</span>
                  </div>
                  <div className="flex justify-between py-0.5 border-b border-slate-900 text-[11px]">
                    <span className="text-slate-500">API cible :</span>
                    <span className="font-semibold text-slate-300">Microsoft Graph</span>
                  </div>
                  <div className="flex justify-between py-0.5 border-b border-slate-900 text-[11px]">
                    <span className="text-slate-500">Type d'autorisation :</span>
                    <span className="font-semibold text-slate-300">Déléguée</span>
                  </div>

                  {workflowState?.status === 'consent_granted' || adminConfig.status === 'entra_configured' || adminConfig.status === 'fully_configured' ? (
                    <div className="flex flex-col sm:flex-row sm:justify-between py-1.5 bg-emerald-950/50 p-2.5 rounded-lg border border-emerald-800/70 gap-1 shadow-sm">
                      <span className="text-emerald-400 font-semibold text-xs flex items-center gap-1.5">
                        <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400 shrink-0" />
                        Permission active et consentie :
                      </span>
                      <code className="text-emerald-300 font-mono font-bold text-xs">OnlineMeetingTranscript.Read.All</code>
                    </div>
                  ) : (
                    <div className="flex flex-col sm:flex-row sm:justify-between py-1 bg-slate-900/60 p-2 rounded-lg border border-slate-800/60 gap-1">
                      <span className="text-slate-400 font-medium">Permission requise :</span>
                      <code className="text-amber-400 font-mono font-bold text-xs">OnlineMeetingTranscript.Read.All</code>
                    </div>
                  )}
                </div>

                {workflowState?.status === 'consent_granted' || adminConfig.status === 'entra_configured' || adminConfig.status === 'fully_configured' ? (
                  <div className="space-y-1 text-xs text-emerald-300 pt-1 bg-emerald-950/30 p-2.5 rounded-lg border border-emerald-800/50">
                    <span className="font-semibold block text-[11px] text-emerald-300 flex items-center gap-1.5">
                      <Check className="w-3.5 h-3.5 text-emerald-400 shrink-0" />
                      Statut : Consentement administrateur validé et actif
                    </span>
                    <p className="text-[11px] text-slate-300 leading-relaxed">
                      L'approbation du locataire DGF a été accordée pour le périmètre délégué. L'acquisition du jeton Graph silencieux est désormais pleinement autorisée.
                    </p>
                  </div>
                ) : (
                  <div className="space-y-1 text-xs text-slate-400 pt-1">
                    <span className="text-slate-300 font-semibold block text-[11px]">Action attendue :</span>
                    <ul className="list-disc list-inside space-y-0.5 text-[11px] text-slate-400 pl-1 leading-relaxed">
                      <li>Ajouter la permission déléguée à l'inscription de l'application.</li>
                      <li>Cliquer impérativement sur <strong>« Accorder un consentement d'administrateur »</strong>.</li>
                    </ul>
                  </div>
                )}
              </div>

              <div className="pt-3 border-t border-slate-800/80">
                <a
                  href="https://entra.microsoft.com/"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="w-full px-3.5 py-2 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-500 rounded-lg shadow-sm flex items-center justify-center gap-2 transition-all"
                >
                  <ExternalLink className="w-3.5 h-3.5" />
                  {workflowState?.status === 'consent_granted' || adminConfig.status === 'entra_configured' || adminConfig.status === 'fully_configured'
                    ? 'Vérifier dans le portail Entra ID'
                    : 'Consulter les instructions Entra ID'}
                </a>
              </div>
            </div>

            {/* Volet B : Microsoft Teams */}
            <div className="bg-slate-950/80 border border-slate-800 rounded-xl p-4 space-y-3 flex flex-col justify-between">
              <div className="space-y-2.5">
                <div className="flex items-center justify-between border-b border-slate-800/80 pb-2">
                  <div className="flex items-center gap-2">
                    <span className="w-6 h-6 rounded-md bg-indigo-950 text-indigo-400 font-bold text-xs flex items-center justify-center border border-indigo-800/60">
                      B
                    </span>
                    <h4 className="text-xs font-bold text-white uppercase tracking-wider">
                      Microsoft Teams Admin Center
                    </h4>
                  </div>
                  <span className="text-[10px] font-mono text-indigo-400 bg-indigo-950/60 px-2 py-0.5 rounded border border-indigo-800/40">
                    Stratégies de Réunion
                  </span>
                </div>

                <div className="space-y-2 text-xs text-slate-300">
                  <p className="text-[11px] text-slate-400 leading-relaxed">
                    Dans le <strong>Teams Admin Center</strong> (<em>Meetings &gt; Meeting settings &gt; Transcript API access</em>), vérifier les deux paramètres suivants :
                  </p>

                  <div className="space-y-1.5">
                    <div className="p-2 bg-slate-900/60 rounded-lg border border-slate-800/60 space-y-0.5">
                      <div className="flex items-center justify-between">
                        <code className="text-purple-300 font-mono font-bold text-[11px]">EnableGraphTranscriptAccess</code>
                        <span className="text-[10px] font-semibold text-emerald-400">Valeur : $true</span>
                      </div>
                      <p className="text-[10px] text-slate-400">
                        Autorise l'accès de Microsoft Graph aux transcriptions au niveau du locataire.
                      </p>
                    </div>

                    <div className="p-2 bg-slate-900/60 rounded-lg border border-slate-800/60 space-y-0.5">
                      <div className="flex items-center justify-between">
                        <code className="text-purple-300 font-mono font-bold text-[11px]">EnableAttributedTranscripts</code>
                        <span className="text-[10px] font-semibold text-emerald-400">Valeur : $true</span>
                      </div>
                      <p className="text-[10px] text-slate-400">
                        Autorise l'attribution nominale des interventions aux locuteurs convoqués.
                      </p>
                    </div>
                  </div>

                  <p className="text-[10px] text-slate-500 italic pt-0.5">
                    Ces configurations doivent respecter les politiques de gouvernance et de sécurité du Ministère.
                  </p>
                </div>
              </div>

              <div className="pt-3 border-t border-slate-800/80">
                <a
                  href="https://admin.teams.microsoft.com/"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="w-full px-3.5 py-2 text-xs font-semibold text-white bg-indigo-600 hover:bg-indigo-500 rounded-lg shadow-sm flex items-center justify-center gap-2 transition-all"
                >
                  <ExternalLink className="w-3.5 h-3.5" />
                  Accéder à l'administration Teams
                </a>
              </div>
            </div>
          </div>

          {/* 3. Barre d'actions complémentaires (Copier la demande & Vérification technique) */}
          <div className="flex flex-col sm:flex-row items-center justify-between gap-3 pt-3 border-t border-slate-800/80 bg-slate-950/40 p-3 rounded-lg">
            <div className="flex items-center gap-2 w-full sm:w-auto">
              <button
                onClick={handleCopyRequest}
                className="px-3.5 py-2 text-xs font-semibold text-slate-200 bg-slate-800 hover:bg-slate-700 border border-slate-700 rounded-lg flex items-center gap-2 transition-all shadow-sm w-full sm:w-auto justify-center"
                title="Copier le texte formaté pour envoi direct par courriel à l'administrateur DGF"
              >
                {copied ? (
                  <>
                    <Check className="w-3.5 h-3.5 text-emerald-400" />
                    <span className="text-emerald-300 font-bold">Demande copiée dans le presse-papiers !</span>
                  </>
                ) : (
                  <>
                    <Copy className="w-3.5 h-3.5 text-amber-400" />
                    <span>Copier la demande d'autorisation</span>
                  </>
                )}
              </button>

              {onRunTechnicalCheck && (
                <button
                  onClick={handleExecuteTechnicalCheck}
                  disabled={isCheckingTechnical || !user}
                  className="px-3 py-2 text-xs font-medium text-slate-300 hover:text-white bg-slate-900 hover:bg-slate-800 border border-slate-800 rounded-lg flex items-center gap-1.5 transition-colors disabled:opacity-50"
                  title="Effectuer une requête test réelle sur Microsoft Graph"
                >
                  <RefreshCw className={`w-3.5 h-3.5 ${isCheckingTechnical ? 'animate-spin text-blue-400' : 'text-slate-400'}`} />
                  {isCheckingTechnical ? 'Test en cours...' : 'Tester les droits Graph'}
                </button>
              )}
            </div>

            <div className="flex items-center gap-2 text-[11px] text-slate-400 font-mono">
              <Clock className="w-3 h-3 text-slate-500" />
              <span>Dernière mise à jour : {new Date(adminConfig.lastUpdated).toLocaleDateString('fr-FR', { day: 'numeric', month: 'short', year: 'numeric' })}</span>
              {adminConfig.updatedByUpn && (
                <span className="text-slate-500">({adminConfig.updatedByUpn})</span>
              )}
            </div>
          </div>

          {/* Résultat du test technique en direct */}
          {techCheckResult && (
            <div className={`p-3 rounded-lg text-xs border flex items-start gap-2.5 transition-all ${
              techCheckResult.success
                ? 'bg-emerald-950/60 border-emerald-800 text-emerald-200'
                : 'bg-rose-950/60 border-rose-800 text-rose-200'
            }`}>
              {techCheckResult.success ? (
                <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0 mt-0.5" />
              ) : (
                <AlertTriangle className="w-4 h-4 text-rose-400 shrink-0 mt-0.5" />
              )}
              <div className="space-y-0.5">
                <span className="font-bold block text-white">{techCheckResult.message}</span>
                {techCheckResult.details && (
                  <p className="text-[11px] opacity-90 font-mono">{techCheckResult.details}</p>
                )}
              </div>
            </div>
          )}
        </div>
      )}

      {/* 4. Modal de mise à jour explicite du statut */}
      {showStatusModal && (
        <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-lg w-full p-6 space-y-4 shadow-2xl">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <div className="flex items-center gap-2">
                <Settings className="w-4 h-4 text-amber-400" />
                <h3 className="text-sm font-bold text-white">
                  Mise à jour du statut d'autorisation Microsoft 365
                </h3>
              </div>
              <button
                onClick={() => setShowStatusModal(false)}
                className="text-slate-400 hover:text-white p-1 rounded"
              >
                ✕
              </button>
            </div>

            <p className="text-xs text-slate-300 leading-relaxed">
              Sélectionnez l'état d'avancement réel des configurations effectuées par l'administrateur du locataire DGF.
            </p>

            <div className="space-y-2">
              {(Object.keys(ADMIN_AUTH_STATUS_CONFIG) as AdminAuthStatus[]).map((statusKey) => {
                const conf = ADMIN_AUTH_STATUS_CONFIG[statusKey];
                const isSelected = selectedNewStatus === statusKey;
                return (
                  <label
                    key={statusKey}
                    className={`flex items-start gap-3 p-3 rounded-xl border cursor-pointer transition-all ${
                      isSelected
                        ? 'bg-blue-950/50 border-blue-500/80 shadow-sm'
                        : 'bg-slate-950/60 border-slate-800 hover:border-slate-700'
                    }`}
                  >
                    <input
                      type="radio"
                      name="admin_status"
                      value={statusKey}
                      checked={isSelected}
                      onChange={() => setSelectedNewStatus(statusKey)}
                      className="mt-0.5 text-blue-600 focus:ring-blue-500 bg-slate-900 border-slate-700"
                    />
                    <div className="space-y-0.5 flex-1">
                      <div className="flex items-center justify-between">
                        <span className={`text-xs font-bold ${isSelected ? 'text-white' : 'text-slate-200'}`}>
                          {conf.label}
                        </span>
                        {statusKey === 'fully_configured' && (
                          <span className="text-[10px] text-amber-400 bg-amber-950 px-1.5 py-0.5 rounded border border-amber-800">
                            Validation formelle
                          </span>
                        )}
                      </div>
                      <p className="text-[11px] text-slate-400 leading-normal">
                        {conf.description}
                      </p>
                    </div>
                  </label>
                );
              })}
            </div>

            {selectedNewStatus === 'fully_configured' && (
              <div className="p-3 bg-amber-950/40 border border-amber-800/80 rounded-xl text-xs text-amber-200 flex items-start gap-2">
                <AlertTriangle className="w-4 h-4 text-amber-400 shrink-0 mt-0.5" />
                <p className="text-[11px] leading-relaxed">
                  <strong>Avertissement de gouvernance :</strong> Le statut « Configuration complète » ne doit être consigné que si l'administrateur DGF a effectivement accordé le consentement pour <code className="text-white font-mono">OnlineMeetingTranscript.Read.All</code> et activé les stratégies Teams requises.
                </p>
              </div>
            )}

            <div className="space-y-1.5">
              <label className="text-xs font-semibold text-slate-300">
                Notes ou référence du ticket DSI (optionnel) :
              </label>
              <input
                type="text"
                value={statusChangeNotes}
                onChange={(e) => setStatusChangeNotes(e.target.value)}
                placeholder="Ex : Ticket DSI-2026-042 - Validé par l'admin M365 DGF"
                className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-xs text-white placeholder-slate-500 focus:outline-none focus:border-blue-500 font-sans"
              />
            </div>

            <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-800">
              <button
                onClick={() => setShowStatusModal(false)}
                className="px-4 py-2 text-xs font-medium text-slate-400 hover:text-white bg-slate-800 hover:bg-slate-700 rounded-lg transition-colors"
              >
                Annuler
              </button>
              <button
                onClick={handleSaveStatus}
                className="px-4 py-2 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-500 rounded-lg shadow-sm transition-all"
              >
                Enregistrer la déclaration
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
