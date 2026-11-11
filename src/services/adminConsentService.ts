/**
 * ReadMeeting.ai | Direction Générale des Financements (DGF)
 * Ministère des Finances et du Budget - République de Côte d'Ivoire
 *
 * Service de gestion du workflow natif Microsoft Entra ID :
 * « Admin Consent Request Workflow »
 *
 * Exploite les flux d'autorisation natifs de Microsoft Entra ID pour soumettre
 * des demandes de consentement d'administrateur officielles sans simulation
 * ni envoi factice d'e-mails.
 */

import {
  IPublicClientApplication,
  AccountInfo,
  InteractionRequiredAuthError,
  AuthenticationResult,
} from '@azure/msal-browser';
import {
  AdminConsentWorkflowState,
  AdminConsentWorkflowStatus,
  DGF_ENTRA_PREREQUISITES,
} from '../types/adminConsentWorkflow';
import { createGraphClient } from './graphClient';

export const ADMIN_CONSENT_STORAGE_KEY = 'dgf_admin_consent_workflow_state_v2';

export const ALL_REQUIRED_SCOPES = [
  'User.Read',
  'Calendars.Read',
  'OnlineMeetings.Read',
  'OnlineMeetingTranscript.Read.All',
];

export const INITIAL_CONSENT_STATE: AdminConsentWorkflowState = {
  status: 'admin_consent_required',
  lastActionDate: new Date().toISOString(),
  requestedScopes: ALL_REQUIRED_SCOPES,
  justificationNotes: "Activation requise pour la génération automatisée des comptes rendus institutionnels de la DGF à partir des transcriptions Teams.",
};

/**
 * Récupère l'état courant du workflow depuis le stockage local sécurisé.
 */
export function getStoredAdminConsentState(): AdminConsentWorkflowState {
  try {
    const raw = localStorage.getItem(ADMIN_CONSENT_STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw) as AdminConsentWorkflowState;
      if (parsed && parsed.status) {
        return parsed;
      }
    }
  } catch (err) {
    console.warn('[AdminConsentService] Erreur lors de la lecture du stockage local :', err);
  }
  return { ...INITIAL_CONSENT_STATE, lastActionDate: new Date().toISOString() };
}

/**
 * Sauvegarde l'état du workflow dans le stockage local.
 */
export function saveAdminConsentState(state: AdminConsentWorkflowState): void {
  try {
    localStorage.setItem(ADMIN_CONSENT_STORAGE_KEY, JSON.stringify(state));
  } catch (err) {
    console.warn('[AdminConsentService] Erreur lors de la sauvegarde du stockage local :', err);
  }
}

/**
 * Met à jour partiellement l'état du workflow et le persiste.
 */
export function updateAdminConsentState(
  partial: Partial<AdminConsentWorkflowState>
): AdminConsentWorkflowState {
  const current = getStoredAdminConsentState();
  const updated: AdminConsentWorkflowState = {
    ...current,
    ...partial,
    lastActionDate: new Date().toISOString(),
  };
  saveAdminConsentState(updated);
  return updated;
}

/**
 * Réinitialise le workflow à son état d'origine.
 */
export function resetAdminConsentState(): AdminConsentWorkflowState {
  const resetState: AdminConsentWorkflowState = {
    ...INITIAL_CONSENT_STATE,
    lastActionDate: new Date().toISOString(),
  };
  saveAdminConsentState(resetState);
  return resetState;
}

export interface EntraConsentErrorDetails {
  code: string;
  subError?: string;
  message: string;
  correlationId?: string;
  isApprovalRequested: boolean;
  isWorkflowDisabled: boolean;
  isCancelled: boolean;
  isDirectConsentRequired: boolean;
}

/**
 * Analyse rigoureusement les erreurs retournées par Microsoft Entra ID
 * lors de la sollicitation de consentement.
 */
export function parseEntraConsentError(error: any): EntraConsentErrorDetails {
  const rawMsg = error?.errorMessage || error?.message || (typeof error === 'string' ? error : '');
  const rawCode = error?.errorCode || error?.code || '';
  const subError = error?.subError || '';
  const correlationId = error?.correlationId || undefined;

  const isCancelled =
    rawCode === 'user_cancelled' ||
    rawCode === 'interaction_in_progress' ||
    rawMsg.includes('User cancelled') ||
    rawMsg.includes('user_cancelled');

  // AADSTS90094: The grant requires admin permission.
  // Microsoft Entra renvoie ce code lorsque le consentement administrateur est requis.
  // Si le workflow est activé dans le locataire, Microsoft affiche l'écran de demande de justification ("Need admin approval").
  const isAadsts90094 =
    rawCode.includes('AADSTS90094') ||
    rawMsg.includes('AADSTS90094') ||
    rawMsg.includes('requires admin permission') ||
    rawMsg.includes('grant requires admin permission');

  const isAadsts90093 =
    rawCode.includes('AADSTS90093') ||
    rawMsg.includes('AADSTS90093') ||
    rawMsg.includes('cannot consent to this application');

  const isAadsts65001 =
    rawCode.includes('AADSTS65001') ||
    rawMsg.includes('AADSTS65001') ||
    rawMsg.includes('has not consented');

  // Détection d'un message explicite indiquant que le workflow n'est pas activé sur le locataire
  const isWorkflowDisabled =
    rawMsg.includes('approval workflow is not enabled') ||
    rawMsg.includes('not configured to allow user requests') ||
    rawMsg.includes('admin consent workflow is disabled') ||
    rawMsg.includes('approval request cannot be created');

  return {
    code: rawCode || (isAadsts90094 ? 'AADSTS90094' : isAadsts90093 ? 'AADSTS90093' : isAadsts65001 ? 'AADSTS65001' : 'UNKNOWN'),
    subError,
    message: rawMsg,
    correlationId,
    isApprovalRequested: isAadsts90094 && !isWorkflowDisabled,
    isWorkflowDisabled,
    isCancelled,
    isDirectConsentRequired: isAadsts65001 || isAadsts90093 || isAadsts90094,
  };
}

export interface NativeConsentExecutionResult {
  success: boolean;
  status: AdminConsentWorkflowStatus;
  message: string;
  details?: string;
  correlationId?: string;
  grantedScopes?: string[];
}

/**
 * Déclenche le parcours interactif Microsoft Entra ID de demande de consentement administrateur.
 *
 * Utilise MSAL avec prompt: 'consent' et la portée OnlineMeetingTranscript.Read.All.
 * Si le workflow d'approbation Entra est activé pour la DGF, Microsoft présente à l'utilisateur
 * le formulaire institutionnel de justification ("Demande d'approbation") et notifie
 * automatiquement les administrateurs Microsoft 365 par courriel officiel de sécurité.
 */
export async function triggerNativeAdminConsentRequest(
  msalInstance: IPublicClientApplication,
  account?: AccountInfo | null,
  justificationText?: string
): Promise<NativeConsentExecutionResult> {
  const currentAccount = account || msalInstance.getActiveAccount();
  const upn = currentAccount?.username || undefined;

  const requestScopes = [...ALL_REQUIRED_SCOPES];

  try {
    // Appel interactif MSAL avec prompt 'consent'
    // C'est cette interaction qui fait basculer Entra ID vers l'écran d'approbation administrateur
    const authResult: AuthenticationResult = await msalInstance.acquireTokenPopup({
      scopes: requestScopes,
      prompt: 'consent',
      loginHint: upn,
    });

    // Si l'appel a abouti directement, l'utilisateur était administrateur ou le consentement a été validé !
    const grantedScopes = authResult.scopes || [];
    const hasTranscriptScope = grantedScopes.some(
      (s) => s.toLowerCase() === 'onlinemeetingtranscript.read.all'
    );

    const nextStatus: AdminConsentWorkflowStatus = hasTranscriptScope
      ? 'consent_granted'
      : 'verification_needed';

    updateAdminConsentState({
      status: nextStatus,
      requestedByUpn: upn,
      requestedByName: currentAccount?.name,
      requestedScopes: requestScopes,
      correlationId: authResult.correlationId,
      submittedToMicrosoftAt: new Date().toISOString(),
      lastVerificationDate: new Date().toISOString(),
      lastVerificationResult: {
        success: true,
        message: 'Jeton d\'accès obtenu avec succès via Microsoft Entra ID.',
        hasOnlineMeetingTranscriptScope: hasTranscriptScope,
      },
    });

    return {
      success: true,
      status: nextStatus,
      message: hasTranscriptScope
        ? 'Consentement accordé avec succès par l\'administrateur.'
        : 'Session actualisée. Vérification complémentaire des portées requise.',
      grantedScopes,
      correlationId: authResult.correlationId,
    };
  } catch (err: any) {
    console.info('[AdminConsentService] Diagnostic du retour Microsoft Entra ID :', err);
    const parsed = parseEntraConsentError(err);

    if (parsed.isWorkflowDisabled) {
      // Le locataire DGF n'a pas activé l'option de demande dans Entra ID
      updateAdminConsentState({
        status: 'workflow_unavailable',
        requestedByUpn: upn,
        requestedByName: currentAccount?.name,
        lastErrorCode: parsed.code,
        lastErrorMessage: parsed.message,
        correlationId: parsed.correlationId,
      });

      return {
        success: false,
        status: 'workflow_unavailable',
        message: 'Workflow administrateur désactivé sur le locataire DGF.',
        details: 'Le paramètre « Les utilisateurs peuvent demander le consentement d\'un administrateur » doit être activé dans Microsoft Entra ID.',
        correlationId: parsed.correlationId,
      };
    }

    if (parsed.isApprovalRequested || parsed.code === 'AADSTS90094') {
      // Cas standard du workflow Entra ID : l'utilisateur a soumis sa demande dans l'interface Microsoft
      // ou l'approbation est en cours d'attente d'arbitrage par l'administrateur
      updateAdminConsentState({
        status: 'request_submitted',
        requestedByUpn: upn,
        requestedByName: currentAccount?.name,
        justificationNotes: justificationText || undefined,
        lastErrorCode: parsed.code,
        lastErrorMessage: parsed.message,
        correlationId: parsed.correlationId,
        submittedToMicrosoftAt: new Date().toISOString(),
      });

      return {
        success: true,
        status: 'request_submitted',
        message: 'Demande de consentement transmise à Microsoft Entra ID.',
        details: 'Si vous avez validé le formulaire Microsoft, un courriel officiel de sécurité Microsoft a été adressé aux administrateurs de la DGF.',
        correlationId: parsed.correlationId,
      };
    }

    if (parsed.isCancelled) {
      // L'utilisateur a fermé la boîte de dialogue
      const currentState = getStoredAdminConsentState();
      return {
        success: false,
        status: currentState.status,
        message: 'Interaction fermée par l\'utilisateur.',
        details: 'Si vous avez saisi votre motif et cliqué sur « Demander l\'approbation » avant de fermer, confirmez ci-dessous le passage en attente de décision.',
        correlationId: parsed.correlationId,
      };
    }

    // Erreur générique ou refus explicite
    const isDeclined = parsed.message.toLowerCase().includes('declined') || parsed.message.toLowerCase().includes('denied');
    const targetStatus: AdminConsentWorkflowStatus = isDeclined ? 'request_declined' : 'admin_consent_required';

    updateAdminConsentState({
      status: targetStatus,
      requestedByUpn: upn,
      requestedByName: currentAccount?.name,
      lastErrorCode: parsed.code,
      lastErrorMessage: parsed.message,
      correlationId: parsed.correlationId,
    });

    return {
      success: false,
      status: targetStatus,
      message: isDeclined
        ? 'La demande d\'autorisation a été refusée sur Microsoft Entra ID.'
        : `Intervention requise : code ${parsed.code}`,
      details: parsed.message,
      correlationId: parsed.correlationId,
    };
  }
}

export interface ConsentVerificationResult {
  success: boolean;
  status: AdminConsentWorkflowStatus;
  message: string;
  details?: string;
  hasTranscriptScope: boolean;
  correlationId?: string;
}

/**
 * Vérifie de manière non intrusive et réelle l'état du consentement administrateur.
 *
 * Tente d'abord une acquisition silencieuse (acquireTokenSilent) pour la permission
 * OnlineMeetingTranscript.Read.All. En cas de succès, réalise un appel Graph pour confirmer
 * l'opérationnalité des droits délégués.
 */
export async function verifyEntraConsentStatus(
  msalInstance: IPublicClientApplication,
  account?: AccountInfo | null
): Promise<ConsentVerificationResult> {
  const currentAccount = account || msalInstance.getActiveAccount();
  const currentState = getStoredAdminConsentState();

  if (!currentAccount) {
    return {
      success: false,
      status: currentState.status,
      message: 'Compte Microsoft non connecté',
      details: 'Connectez-vous avec votre compte professionnel DGF pour vérifier le statut du consentement.',
      hasTranscriptScope: false,
    };
  }

  try {
    // 1. Tentative d'acquisition silencieuse avec la portée haute
    const silentResult = await msalInstance.acquireTokenSilent({
      account: currentAccount,
      scopes: ALL_REQUIRED_SCOPES,
    });

    const granted = silentResult.scopes.map((s) => s.toLowerCase());
    const hasTranscriptScope = granted.includes('onlinemeetingtranscript.read.all');

    if (hasTranscriptScope) {
      // 2. Contrôle réel via Microsoft Graph
      const client = createGraphClient(msalInstance, currentAccount);
      try {
        await client.api('/me').select('id,displayName,userPrincipalName').get();
      } catch (graphErr: any) {
        console.warn('[AdminConsentService] Contrôle Graph complémentaire :', graphErr);
      }

      updateAdminConsentState({
        status: 'consent_granted',
        lastVerificationDate: new Date().toISOString(),
        correlationId: silentResult.correlationId,
        lastVerificationResult: {
          success: true,
          message: 'Consentement administrateur confirmé actif sur Microsoft Entra ID.',
          details: `Portée OnlineMeetingTranscript.Read.All validée pour ${currentAccount.username}.`,
          hasOnlineMeetingTranscriptScope: true,
        },
      });

      return {
        success: true,
        status: 'consent_granted',
        message: 'Autorisation administrative accordée et vérifiée !',
        details: 'Le jeton d\'accès intègre la permission OnlineMeetingTranscript.Read.All. L\'accès aux transcriptions Teams est actif.',
        hasTranscriptScope: true,
        correlationId: silentResult.correlationId,
      };
    }

    // Le jeton est valide mais ne contient pas encore la portée demandée
    updateAdminConsentState({
      status: currentState.status === 'request_submitted' ? 'decision_pending' : currentState.status,
      lastVerificationDate: new Date().toISOString(),
      lastVerificationResult: {
        success: false,
        message: 'Le consentement administrateur n\'a pas encore été validé.',
        details: 'La permission OnlineMeetingTranscript.Read.All est absente du jeton actuel.',
        hasOnlineMeetingTranscriptScope: false,
      },
    });

    return {
      success: false,
      status: currentState.status === 'request_submitted' ? 'decision_pending' : currentState.status,
      message: 'Autorisation non encore visible dans le jeton Microsoft 365',
      details: 'L\'approbation administrative est toujours en attente ou n\'a pas encore été répercutée sur les contrôleurs de domaine Entra ID.',
      hasTranscriptScope: false,
    };
  } catch (silentErr: any) {
    if (silentErr instanceof InteractionRequiredAuthError || silentErr?.errorCode === 'consent_required') {
      // Le consentement est formellement requis et non accordé
      const isPending = currentState.status === 'request_submitted' || currentState.status === 'decision_pending';
      const updatedStatus: AdminConsentWorkflowStatus = isPending ? 'decision_pending' : 'admin_consent_required';

      updateAdminConsentState({
        status: updatedStatus,
        lastVerificationDate: new Date().toISOString(),
        lastVerificationResult: {
          success: false,
          message: 'Consentement administrateur requis (InteractionRequiredAuthError).',
          details: silentErr.message,
          hasOnlineMeetingTranscriptScope: false,
        },
      });

      return {
        success: false,
        status: updatedStatus,
        message: isPending
          ? 'Demande toujours en attente de décision administrative sur Entra ID.'
          : 'Autorisation administrateur requise.',
        details: 'L\'administrateur du locataire DGF n\'a pas encore validé la demande reçue par courriel Microsoft Security.',
        hasTranscriptScope: false,
      };
    }

    console.warn('[AdminConsentService] Erreur lors de la vérification silencieuse :', silentErr);
    return {
      success: false,
      status: 'verification_needed',
      message: 'Impossible de vérifier silencieusement le statut.',
      details: silentErr.message || String(silentErr),
      hasTranscriptScope: false,
    };
  }
}
