import React, { useState } from 'react';
import {
  Send,
  ShieldCheck,
  RefreshCw,
  CheckCircle2,
  AlertTriangle,
  Clock,
  ExternalLink,
  ChevronDown,
  ChevronUp,
  Info,
  XCircle,
  HelpCircle,
  Settings,
  MailCheck,
  FileCheck,
  Sparkles,
  Building,
  Key
} from 'lucide-react';
import {
  AdminConsentWorkflowState,
  AdminConsentWorkflowStatus,
  ADMIN_CONSENT_WORKFLOW_META,
  DGF_ENTRA_PREREQUISITES,
} from '../types/adminConsentWorkflow';
import {
  NativeConsentExecutionResult,
  ConsentVerificationResult,
} from '../services/adminConsentService';
import { DgfUserProfile } from '../types/auth';

interface AdminConsentWorkflowSectionProps {
  user: DgfUserProfile | null;
  workflowState: AdminConsentWorkflowState;
  onRequestConsent: (justification?: string) => Promise<NativeConsentExecutionResult>;
  onVerifyConsent: () => Promise<ConsentVerificationResult>;
  onUpdateWorkflowState: (partial: Partial<AdminConsentWorkflowState>) => void;
  isRequesting?: boolean;
  isVerifying?: boolean;
}

export const AdminConsentWorkflowSection: React.FC<AdminConsentWorkflowSectionProps> = ({
  user,
  workflowState,
  onRequestConsent,
  onVerifyConsent,
  onUpdateWorkflowState,
  isRequesting = false,
  isVerifying = false,
}) => {
  const [showPrereqModal, setShowPrereqModal] = useState<boolean>(false);
  const [showManualStateModal, setShowManualStateModal] = useState<boolean>(false);
  const [justificationInput, setJustificationInput] = useState<string>(
    workflowState.justificationNotes ||
      "Demande d'autorisation pour ReadMeeting.ai | DGF : activation de la transcription Microsoft Teams pour la génération automatisée des procès-verbaux de réunions de la Direction Générale des Financements."
  );
  const [actionFeedback, setActionFeedback] = useState<{
    type: 'success' | 'warning' | 'error' | 'info';
    title: string;
    message: string;
    details?: string;
  } | null>(null);

  const currentMeta = ADMIN_CONSENT_WORKFLOW_META[workflowState.status] || ADMIN_CONSENT_WORKFLOW_META.admin_consent_required;

  const handleTriggerRequest = async () => {
    setActionFeedback(null);
    try {
      const res = await onRequestConsent(justificationInput);
      if (res.success) {
        if (res.status === 'consent_granted') {
          setActionFeedback({
            type: 'success',
            title: 'Consentement accordé immédiatement',
            message: res.message,
            details: `Portées accordées : ${res.grantedScopes?.join(', ')}`,
          });
        } else {
          setActionFeedback({
            type: 'info',
            title: 'Demande soumise au workflow Microsoft Entra ID',
            message: res.message,
            details: res.details,
          });
        }
      } else {
        if (res.status === 'workflow_unavailable') {
          setActionFeedback({
            type: 'warning',
            title: 'Workflow de consentement administrateur inactif',
            message: res.message,
            details: res.details,
          });
        } else {
          setActionFeedback({
            type: 'warning',
            title: 'Notification d\'interaction Microsoft',
            message: res.message,
            details: res.details,
          });
        }
      }
    } catch (err: any) {
      setActionFeedback({
        type: 'error',
        title: 'Erreur lors du déclenchement du flux Microsoft',
        message: err.message || 'Une exception non interceptée s\'est produite.',
      });
    }
  };

  const handleTriggerVerification = async () => {
    setActionFeedback(null);
    try {
      const res = await onVerifyConsent();
      if (res.success) {
        setActionFeedback({
          type: 'success',
          title: 'Vérification positive — Droits confirmés',
          message: res.message,
          details: res.details,
        });
      } else {
        setActionFeedback({
          type: res.status === 'decision_pending' ? 'info' : 'warning',
          title: 'Statut du consentement Microsoft Entra ID',
          message: res.message,
          details: res.details,
        });
      }
    } catch (err: any) {
      setActionFeedback({
        type: 'error',
        title: 'Erreur de vérification',
        message: err.message || 'Impossible de contacter Microsoft Entra ID.',
      });
    }
  };

  // Les 4 étapes linéaires principales
  const linearSteps: AdminConsentWorkflowStatus[] = [
    'admin_consent_required',
    'request_submitted',
    'decision_pending',
    'consent_granted',
  ];

  const getStepProgressIndex = (status: AdminConsentWorkflowStatus): number => {
    switch (status) {
      case 'admin_consent_required':
        return 0;
      case 'request_submitted':
        return 1;
      case 'decision_pending':
        return 2;
      case 'consent_granted':
        return 3;
      case 'request_declined':
      case 'workflow_unavailable':
      case 'verification_needed':
        return 1; // Interruption ou statut intermédiaire
    }
  };

  const activeStepIdx = getStepProgressIndex(workflowState.status);

  return (
    <div className="bg-slate-950/90 border border-blue-900/40 rounded-xl p-4 sm:p-5 space-y-4 shadow-md">
      {/* 1. Entête du volet Workflow Natif */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-800/80 pb-3">
        <div className="flex items-center gap-2.5">
          <div className="w-8 h-8 rounded-lg bg-blue-950 border border-blue-600/40 flex items-center justify-center text-blue-400 shrink-0">
            <ShieldCheck className="w-4 h-4" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h4 className="text-xs sm:text-sm font-bold text-white uppercase tracking-wider">
                Workflow Natif Microsoft Entra ID — Demande de Consentement Administrateur
              </h4>
              <span className={`px-2 py-0.5 text-[10px] font-semibold rounded border ${currentMeta.badgeClass}`}>
                {currentMeta.shortLabel}
              </span>
            </div>
            <p className="text-[11px] text-slate-400">
              Mécanisme officiel Microsoft <em>« Admin consent request workflow »</em> avec envoi automatique de courriel Microsoft Security aux réviseurs DGF.
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 self-end sm:self-center">
          <button
            onClick={() => setShowPrereqModal(true)}
            className="px-2.5 py-1 text-[11px] font-medium text-slate-300 hover:text-white bg-slate-900 hover:bg-slate-800 border border-slate-800 rounded-lg flex items-center gap-1.5 transition-colors"
            title="Consulter les prérequis de configuration du locataire Microsoft Entra ID"
          >
            <HelpCircle className="w-3.5 h-3.5 text-blue-400" />
            <span className="hidden md:inline">Prérequis Entra ID</span>
          </button>

          <button
            onClick={() => setShowManualStateModal(true)}
            className="p-1.5 text-slate-400 hover:text-white bg-slate-900 hover:bg-slate-800 border border-slate-800 rounded-lg transition-colors"
            title="Ajuster manuellement l'état du workflow"
          >
            <Settings className="w-3.5 h-3.5 text-slate-400" />
          </button>
        </div>
      </div>

      {/* 2. Stepper visuel des 4 étapes clés */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-2 pt-1">
        {linearSteps.map((stepKey, idx) => {
          const stepMeta = ADMIN_CONSENT_WORKFLOW_META[stepKey];
          const isCurrent = workflowState.status === stepKey;
          const isPassed = activeStepIdx > idx;

          return (
            <div
              key={stepKey}
              className={`p-2.5 rounded-lg border text-left transition-all ${
                isCurrent
                  ? 'bg-blue-950/60 border-blue-500 shadow-sm ring-1 ring-blue-500/30'
                  : isPassed
                  ? 'bg-slate-900/60 border-emerald-900/60'
                  : 'bg-slate-900/30 border-slate-800/60 opacity-60'
              }`}
            >
              <div className="flex items-center justify-between mb-1">
                <span className={`text-[10px] font-mono font-bold ${
                  isCurrent ? 'text-blue-300' : isPassed ? 'text-emerald-400' : 'text-slate-500'
                }`}>
                  Étape {idx + 1}
                </span>
                {isPassed && <CheckCircle2 className="w-3 h-3 text-emerald-400" />}
                {isCurrent && <Clock className="w-3 h-3 text-blue-400 animate-pulse" />}
              </div>
              <p className={`text-[11px] font-semibold leading-tight ${
                isCurrent ? 'text-white' : isPassed ? 'text-slate-200' : 'text-slate-400'
              }`}>
                {stepMeta.shortLabel}
              </p>
            </div>
          );
        })}
      </div>

      {/* 3. Statut spécifique d'alerte si statut non linéaire (refus, désactivé, vérification) */}
      {(workflowState.status === 'workflow_unavailable' ||
        workflowState.status === 'request_declined' ||
        workflowState.status === 'verification_needed') && (
        <div className={`p-3 rounded-lg border text-xs flex items-start gap-2.5 ${
          workflowState.status === 'request_declined'
            ? 'bg-rose-950/60 border-rose-800 text-rose-200'
            : workflowState.status === 'workflow_unavailable'
            ? 'bg-orange-950/60 border-orange-800 text-orange-200'
            : 'bg-cyan-950/60 border-cyan-800 text-cyan-200'
        }`}>
          {workflowState.status === 'request_declined' ? (
            <XCircle className="w-4 h-4 text-rose-400 shrink-0 mt-0.5" />
          ) : workflowState.status === 'workflow_unavailable' ? (
            <AlertTriangle className="w-4 h-4 text-orange-400 shrink-0 mt-0.5" />
          ) : (
            <Info className="w-4 h-4 text-cyan-400 shrink-0 mt-0.5" />
          )}
          <div className="space-y-0.5">
            <span className="font-bold block">{currentMeta.label}</span>
            <p className="text-[11px] opacity-90">{currentMeta.description}</p>
            <p className="text-[11px] font-medium pt-1 underline">{currentMeta.actionAdvice}</p>
          </div>
        </div>
      )}

      {/* 4. Zone d'action principale : Bouton officiel et Bouton de vérification */}
      <div className="bg-slate-900/70 border border-slate-800 rounded-xl p-3.5 space-y-3">
        <div className="space-y-1">
          <label className="text-xs font-semibold text-slate-200 flex items-center justify-between">
            <span>Motif d'accompagnement (transmis dans le formulaire Microsoft) :</span>
            <span className="text-[10px] text-slate-500 font-mono">Délégué DGF</span>
          </label>
          <input
            type="text"
            value={justificationInput}
            onChange={(e) => setJustificationInput(e.target.value)}
            disabled={isRequesting || workflowState.status === 'consent_granted'}
            placeholder="Justification institutionnelle..."
            className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-xs text-slate-200 placeholder-slate-600 focus:outline-none focus:border-blue-500 font-sans disabled:opacity-60"
          />
        </div>

        <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-2.5 pt-1">
          {/* Bouton A : Déclencher la demande officielle */}
          <button
            onClick={handleTriggerRequest}
            disabled={isRequesting || !user || workflowState.status === 'consent_granted'}
            className="px-4 py-2.5 text-xs font-bold text-white bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 rounded-lg shadow-md flex items-center justify-center gap-2 transition-all disabled:opacity-50 disabled:cursor-not-allowed cursor-pointer"
            title="Ouvre la fenêtre officielle Microsoft pour demander l'approbation de l'administrateur"
          >
            {isRequesting ? (
              <>
                <RefreshCw className="w-4 h-4 animate-spin text-white" />
                <span>Ouverture du formulaire Microsoft...</span>
              </>
            ) : (
              <>
                <Send className="w-4 h-4 text-blue-200" />
                <span>Demander les autorisations administrateur</span>
              </>
            )}
          </button>

          {/* Bouton B : Vérifier silencieusement l'état réel */}
          <button
            onClick={handleTriggerVerification}
            disabled={isVerifying || !user}
            className="px-3.5 py-2.5 text-xs font-semibold text-slate-300 hover:text-white bg-slate-800 hover:bg-slate-700 border border-slate-700 rounded-lg flex items-center justify-center gap-2 transition-all disabled:opacity-50 cursor-pointer"
            title="Interroge silencieusement Microsoft Entra ID pour tester si le consentement a été accordé"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${isVerifying ? 'animate-spin text-cyan-400' : 'text-slate-400'}`} />
            <span>{isVerifying ? 'Vérification en cours...' : 'Vérifier l\'état du consentement'}</span>
          </button>
        </div>

        <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-800/80 text-[11px] text-slate-400 font-mono">
          <div className="flex items-center gap-2">
            <span className="w-2 h-2 rounded-full bg-blue-400"></span>
            <span>Portée demandée : <code className="text-emerald-400 font-bold">OnlineMeetingTranscript.Read.All</code></span>
          </div>
          {workflowState.submittedToMicrosoftAt && (
            <div className="flex items-center gap-1.5 text-slate-400">
              <MailCheck className="w-3.5 h-3.5 text-blue-400" />
              <span>Soumis à Microsoft le : {new Date(workflowState.submittedToMicrosoftAt).toLocaleDateString('fr-FR', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' })}</span>
            </div>
          )}
        </div>
      </div>

      {/* 5. Retour en direct de l'action */}
      {actionFeedback && (
        <div className={`p-3 rounded-lg text-xs border flex items-start gap-2.5 transition-all ${
          actionFeedback.type === 'success'
            ? 'bg-emerald-950/60 border-emerald-800 text-emerald-200'
            : actionFeedback.type === 'warning'
            ? 'bg-amber-950/60 border-amber-800 text-amber-200'
            : actionFeedback.type === 'error'
            ? 'bg-rose-950/60 border-rose-800 text-rose-200'
            : 'bg-blue-950/60 border-blue-800 text-blue-200'
        }`}>
          {actionFeedback.type === 'success' ? (
            <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0 mt-0.5" />
          ) : actionFeedback.type === 'warning' ? (
            <AlertTriangle className="w-4 h-4 text-amber-400 shrink-0 mt-0.5" />
          ) : actionFeedback.type === 'error' ? (
            <XCircle className="w-4 h-4 text-rose-400 shrink-0 mt-0.5" />
          ) : (
            <Info className="w-4 h-4 text-blue-400 shrink-0 mt-0.5" />
          )}
          <div className="space-y-0.5">
            <span className="font-bold block text-white">{actionFeedback.title}</span>
            <p className="text-[11px] leading-relaxed">{actionFeedback.message}</p>
            {actionFeedback.details && (
              <p className="text-[10px] opacity-80 font-mono mt-1 pt-1 border-t border-current/20">
                {actionFeedback.details}
              </p>
            )}
          </div>
        </div>
      )}

      {/* 6. Modal des Prérequis Administrateur Microsoft Entra ID */}
      {showPrereqModal && (
        <div className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-2xl w-full p-6 space-y-4 shadow-2xl max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <div className="flex items-center gap-2">
                <Building className="w-5 h-5 text-blue-400" />
                <div>
                  <h3 className="text-sm font-bold text-white">
                    Prérequis Microsoft Entra ID — Workflow de Consentement Administrateur
                  </h3>
                  <p className="text-[11px] text-slate-400">
                    Locataire DGF : <span className="font-mono text-slate-300">{DGF_ENTRA_PREREQUISITES.tenantId}</span>
                  </p>
                </div>
              </div>
              <button
                onClick={() => setShowPrereqModal(false)}
                className="text-slate-400 hover:text-white p-1 rounded"
              >
                ✕
              </button>
            </div>

            <div className="space-y-3 text-xs text-slate-300 leading-relaxed">
              <div className="p-3 bg-blue-950/40 border border-blue-800/60 rounded-xl space-y-1">
                <span className="font-semibold text-white block">Chemin d'accès dans le portail d'administration Microsoft :</span>
                <p className="font-mono text-[11px] text-blue-300 bg-slate-950 p-2 rounded border border-blue-900/60">
                  {DGF_ENTRA_PREREQUISITES.adminConsentSettingsPath}
                </p>
              </div>

              <div className="space-y-2">
                <span className="font-semibold text-white block">Paramètres indispensables à activer sur le locataire DGF :</span>
                <div className="space-y-2">
                  {DGF_ENTRA_PREREQUISITES.requiredTenantSettings.map((item, idx) => (
                    <div key={idx} className="p-2.5 bg-slate-950 rounded-lg border border-slate-800/80 space-y-1">
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-1">
                        <span className="font-bold text-slate-200 text-[11px]">{item.setting}</span>
                        <span className="text-[10px] font-mono px-2 py-0.5 rounded bg-emerald-950 text-emerald-300 border border-emerald-800 shrink-0">
                          {item.recommendedValue}
                        </span>
                      </div>
                      <p className="text-[11px] text-slate-400">{item.description}</p>
                    </div>
                  ))}
                </div>
              </div>

              <div className="p-3 bg-slate-950/80 border border-slate-800 rounded-xl space-y-1.5 text-[11px]">
                <span className="font-bold text-white flex items-center gap-1.5">
                  <Key className="w-3.5 h-3.5 text-amber-400" />
                  Que reçoit l'administrateur après le clic ?
                </span>
                <p className="text-slate-400">
                  Microsoft expédie automatiquement un e-mail officiel <strong>« Microsoft Security Notification »</strong> portant l'objet : <em>« Action requise : Demande de consentement administrateur pour ReadMeeting.ai | DGF »</em>. L'administrateur peut cliquer sur le lien direct pour approuver ou refuser la permission dans le portail Entra.
                </p>
              </div>
            </div>

            <div className="flex items-center justify-between pt-3 border-t border-slate-800">
              <a
                href="https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ConsentPoliciesMenuBlade/~/AdminConsentSettings"
                target="_blank"
                rel="noopener noreferrer"
                className="px-3 py-1.5 text-xs font-semibold text-blue-400 hover:text-blue-300 flex items-center gap-1.5"
              >
                <ExternalLink className="w-3.5 h-3.5" />
                Ouvrir les Paramètres Entra ID
              </a>
              <button
                onClick={() => setShowPrereqModal(false)}
                className="px-4 py-2 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-500 rounded-lg transition-all"
              >
                Compris
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 7. Modal de mise à jour manuelle de l'état */}
      {showManualStateModal && (
        <div className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-lg w-full p-6 space-y-4 shadow-2xl">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <div className="flex items-center gap-2">
                <Settings className="w-4 h-4 text-blue-400" />
                <h3 className="text-sm font-bold text-white">
                  Ajuster l'état du Workflow Entra ID
                </h3>
              </div>
              <button
                onClick={() => setShowManualStateModal(false)}
                className="text-slate-400 hover:text-white p-1 rounded"
              >
                ✕
              </button>
            </div>

            <p className="text-xs text-slate-300 leading-relaxed">
              Sélectionnez l'état correspondant aux échanges réels avec l'administrateur Microsoft 365 de la DGF.
            </p>

            <div className="space-y-2 max-h-72 overflow-y-auto pr-1">
              {(Object.keys(ADMIN_CONSENT_WORKFLOW_META) as AdminConsentWorkflowStatus[]).map((stKey) => {
                const meta = ADMIN_CONSENT_WORKFLOW_META[stKey];
                const isSelected = workflowState.status === stKey;

                return (
                  <label
                    key={stKey}
                    className={`flex items-start gap-3 p-2.5 rounded-xl border cursor-pointer transition-all ${
                      isSelected
                        ? 'bg-blue-950/60 border-blue-500'
                        : 'bg-slate-950/60 border-slate-800 hover:border-slate-700'
                    }`}
                  >
                    <input
                      type="radio"
                      name="workflow_manual_status"
                      value={stKey}
                      checked={isSelected}
                      onChange={() => {
                        onUpdateWorkflowState({ status: stKey });
                      }}
                      className="mt-0.5 text-blue-600 bg-slate-900 border-slate-700"
                    />
                    <div className="space-y-0.5 flex-1">
                      <div className="flex items-center justify-between">
                        <span className={`text-xs font-bold ${isSelected ? 'text-white' : 'text-slate-200'}`}>
                          {meta.label}
                        </span>
                        <span className={`text-[10px] px-1.5 py-0.2 rounded border ${meta.badgeClass}`}>
                          {meta.shortLabel}
                        </span>
                      </div>
                      <p className="text-[11px] text-slate-400 leading-normal">
                        {meta.description}
                      </p>
                    </div>
                  </label>
                );
              })}
            </div>

            <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-800">
              <button
                onClick={() => setShowManualStateModal(false)}
                className="px-4 py-2 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-500 rounded-lg shadow-sm"
              >
                Fermer
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
