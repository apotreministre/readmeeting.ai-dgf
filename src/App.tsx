import React, { useState, useEffect, useCallback } from 'react';
import { useMsal } from '@azure/msal-react';
import { IPublicClientApplication } from '@azure/msal-browser';
import { Header } from './components/Header';
import { AuthDiagnosticPanel } from './components/AuthDiagnosticPanel';
import { DashboardView } from './components/DashboardView';
import { ReportsLibraryView } from './components/ReportsLibraryView';
import { ArchitectureView } from './components/ArchitectureView';
import {
  getMsalRuntimeConfig,
  loginRequest,
  resetMsalInstance
} from './config/msalConfig';
import {
  explainEntraIdError,
  fetchGraphUserProfile,
  fetchRealCalendarMeetings,
  createGraphClient,
  createTranscriptGraphClient,
  GraphErrorExplanation
} from './services/graphClient';
import { DgfUserProfile, MsalRuntimeConfig } from './types/auth';
import { MeetingItem } from './types/meeting';
import { checkMeetingTranscriptStatus, explainTranscriptError } from './services/transcriptService';
import { TranscriptCheckResult } from './types/transcript';
import { AdminNotificationConfig } from './types/adminNotification';
import { AdminConsentWorkflowState } from './types/adminConsentWorkflow';
import {
  getStoredAdminConsentState,
  updateAdminConsentState,
  triggerNativeAdminConsentRequest,
  verifyEntraConsentStatus,
  NativeConsentExecutionResult,
  ConsentVerificationResult,
} from './services/adminConsentService';

const ADMIN_NOTIF_STORAGE_KEY = 'dgf_admin_perm_notification_v1';

const getInitialAdminConfig = (): AdminNotificationConfig => {
  try {
    const raw = localStorage.getItem(ADMIN_NOTIF_STORAGE_KEY);
    if (raw) {
      return JSON.parse(raw);
    }
  } catch (e) {
    console.warn('[Storage Error]', e);
  }
  return {
    status: 'pending',
    lastUpdated: new Date().toISOString(),
    isBannerCollapsed: false,
  };
};

interface AppProps {
  onInstanceReplaced?: (newInstance: IPublicClientApplication) => void;
}

export default function App({ onInstanceReplaced }: AppProps) {
  // Récupération de l'instance MSAL déjà initialisée et des comptes réactifs depuis MsalProvider
  const { instance, accounts } = useMsal();

  const [activeTab, setActiveTab] = useState<'dashboard' | 'auth_diagnostic' | 'reports' | 'architecture'>('auth_diagnostic');
  const [msalConfig, setMsalConfig] = useState<MsalRuntimeConfig>(getMsalRuntimeConfig());
  const [user, setUser] = useState<DgfUserProfile | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState<boolean>(false);
  const [isTestingGraph, setIsTestingGraph] = useState<boolean>(false);
  const [lastError, setLastError] = useState<GraphErrorExplanation | null>(null);
  const [graphRawResponse, setGraphRawResponse] = useState<any>(null);

  const [meetings, setMeetings] = useState<MeetingItem[]>([]);
  const [isLoadingMeetings, setIsLoadingMeetings] = useState<boolean>(false);
  const [calendarPeriod, setCalendarPeriod] = useState<'next_7_days' | 'past_30_days' | 'past_90_days'>('next_7_days');

  const handleLoadRealMeetings = useCallback(async (period: 'next_7_days' | 'past_30_days' | 'past_90_days' = calendarPeriod) => {
    setIsLoadingMeetings(true);
    setLastError(null);
    setCalendarPeriod(period);

    try {
      const account = instance.getActiveAccount() || (accounts.length > 0 ? accounts[0] : null);
      if (!account) return;

      const now = new Date();
      let startIso: string;
      let endIso: string;

      if (period === 'past_30_days') {
        startIso = new Date(now.getTime() - 30 * 24 * 3600 * 1000).toISOString();
        endIso = now.toISOString();
      } else if (period === 'past_90_days') {
        startIso = new Date(now.getTime() - 90 * 24 * 3600 * 1000).toISOString();
        endIso = now.toISOString();
      } else {
        // next_7_days
        startIso = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0).toISOString();
        endIso = new Date(now.getTime() + 7 * 24 * 3600 * 1000).toISOString();
      }

      const realMeetings = await fetchRealCalendarMeetings(instance, account, startIso, endIso);
      setMeetings(realMeetings);
    } catch (error: any) {
      console.error('[Fetch Meetings Error]', error);
      setLastError(explainEntraIdError(error));
    } finally {
      setIsLoadingMeetings(false);
    }
  }, [accounts, instance, calendarPeriod]);

  // Synchronisation réactive avec les comptes MSAL initialisés
  useEffect(() => {
    if (accounts && accounts.length > 0) {
      const activeAccount = instance.getActiveAccount() || accounts[0];
      instance.setActiveAccount(activeAccount);
      fetchGraphUserProfile(instance, activeAccount)
        .then((profile) => {
          setUser(profile);
          const now = new Date();
          const start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0).toISOString();
          const end = new Date(now.getTime() + 7 * 24 * 3600 * 1000).toISOString();
          fetchRealCalendarMeetings(instance, activeAccount, start, end)
            .then((m) => setMeetings(m))
            .catch((err) => console.warn('[Auto-fetch meetings error]', err));
        })
        .catch((err) => {
          console.warn('[MSAL Auto-restore warning]', err);
        });
    } else {
      setUser(null);
      setMeetings([]);
    }
  }, [accounts, instance]);

  const handleRefreshConfig = useCallback(async () => {
    try {
      const updatedPca = await resetMsalInstance();
      if (onInstanceReplaced) {
        onInstanceReplaced(updatedPca);
      }
      const updatedConfig = getMsalRuntimeConfig();
      setMsalConfig(updatedConfig);
      setLastError(null);
    } catch (err: any) {
      console.error('[MSAL Reset Error]', err);
      setLastError(explainEntraIdError(err));
    }
  }, [onInstanceReplaced]);

  const handleLogin = async (useRedirect: boolean = false) => {
    setIsLoggingIn(true);
    setLastError(null);
    setGraphRawResponse(null);

    try {
      if (useRedirect) {
        await instance.loginRedirect({
          ...loginRequest,
          prompt: 'select_account',
        });
        return;
      }

      const loginResponse = await instance.loginPopup({
        ...loginRequest,
        prompt: 'select_account',
      });

      if (loginResponse.account) {
        instance.setActiveAccount(loginResponse.account);
        const profile = await fetchGraphUserProfile(instance, loginResponse.account);
        setUser(profile);
        setActiveTab('dashboard');

        const now = new Date();
        const start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0).toISOString();
        const end = new Date(now.getTime() + 7 * 24 * 3600 * 1000).toISOString();
        fetchRealCalendarMeetings(instance, loginResponse.account, start, end)
          .then((m) => setMeetings(m))
          .catch((err) => console.warn('[Auto-fetch meetings error]', err));
      }
    } catch (error: any) {
      console.error('[Entra ID Login Exception]', error);
      const explanation = explainEntraIdError(error);
      setLastError(explanation);
      setActiveTab('auth_diagnostic');
    } finally {
      setIsLoggingIn(false);
    }
  };

  const handleLogout = async () => {
    try {
      const active = instance.getActiveAccount() || (accounts.length > 0 ? accounts[0] : null);
      if (active) {
        await instance.logoutPopup({ account: active });
      }
    } catch (e) {
      console.warn('[MSAL Logout warning]', e);
    } finally {
      setUser(null);
      setMeetings([]);
      setGraphRawResponse(null);
    }
  };

  const handleTestGraphCall = async () => {
    if (!user) return;
    setIsTestingGraph(true);
    setLastError(null);

    try {
      const account = instance.getActiveAccount() || (accounts.length > 0 ? accounts[0] : null);
      if (!account) throw new Error('Aucun compte actif trouvé dans la session Entra ID.');

      const client = createGraphClient(instance, account);
      const res = await client.api('/me').get();
      setGraphRawResponse(res);
    } catch (error: any) {
      console.error('[Graph Test Error]', error);
      setLastError(explainEntraIdError(error));
    } finally {
      setIsTestingGraph(false);
    }
  };

  const [adminConfig, setAdminConfig] = useState<AdminNotificationConfig>(getInitialAdminConfig);
  const [isCheckingTechnical, setIsCheckingTechnical] = useState<boolean>(false);

  const [workflowState, setWorkflowState] = useState<AdminConsentWorkflowState>(getStoredAdminConsentState);
  const [isRequestingConsent, setIsRequestingConsent] = useState<boolean>(false);
  const [isVerifyingConsent, setIsVerifyingConsent] = useState<boolean>(false);

  const handleUpdateAdminConfig = useCallback((partial: Partial<AdminNotificationConfig>) => {
    setAdminConfig((prev) => {
      const next = { ...prev, ...partial };
      try {
        localStorage.setItem(ADMIN_NOTIF_STORAGE_KEY, JSON.stringify(next));
      } catch (e) {
        console.warn('[LocalStorage Error]', e);
      }
      return next;
    });
  }, []);

  const handleTriggerAdminConsent = useCallback(
    async (justification?: string): Promise<NativeConsentExecutionResult> => {
      setIsRequestingConsent(true);
      try {
        const account = instance.getActiveAccount() || (accounts.length > 0 ? accounts[0] : null);
        const res = await triggerNativeAdminConsentRequest(instance, account, justification);
        setWorkflowState(getStoredAdminConsentState());
        if (res.status === 'consent_granted') {
          handleUpdateAdminConfig({
            status: 'entra_configured',
            notes: 'Consentement administrateur accordé avec succès via le flux Microsoft.',
          });
        }
        return res;
      } finally {
        setIsRequestingConsent(false);
      }
    },
    [accounts, instance, handleUpdateAdminConfig]
  );

  const handleVerifyAdminConsent = useCallback(async (): Promise<ConsentVerificationResult> => {
    setIsVerifyingConsent(true);
    try {
      const account = instance.getActiveAccount() || (accounts.length > 0 ? accounts[0] : null);
      const res = await verifyEntraConsentStatus(instance, account);
      setWorkflowState(getStoredAdminConsentState());
      if (res.status === 'consent_granted') {
        handleUpdateAdminConfig({
          status: 'entra_configured',
          notes: 'Consentement administrateur vérifié par jeton silencieux.',
        });
      }
      return res;
    } finally {
      setIsVerifyingConsent(false);
    }
  }, [accounts, instance, handleUpdateAdminConfig]);

  const handleUpdateWorkflowState = useCallback((partial: Partial<AdminConsentWorkflowState>) => {
    const updated = updateAdminConsentState(partial);
    setWorkflowState(updated);
  }, []);

  const handleRunAdminTechnicalCheck = useCallback(async (): Promise<{ success: boolean; message: string; details?: string }> => {
    setIsCheckingTechnical(true);
    try {
      const account = instance.getActiveAccount() || (accounts.length > 0 ? accounts[0] : null);
      if (!account) {
        return {
          success: false,
          message: 'Compte Microsoft non connecté',
          details: 'Connectez-vous avec votre compte professionnel DGF pour tester les autorisations Graph.',
        };
      }

      const client = createGraphClient(instance, account);
      try {
        const testRes = await client.api('/me/onlineMeetings').top(1).get();
        return {
          success: true,
          message: 'Accès Microsoft Graph confirmé',
          details: `L'API onlineMeetings a répondu favorablement (${testRes?.value?.length ?? 0} réunion détectée). L'autorisation globale reste soumise au consentement délégué de transcription.`,
        };
      } catch (err: any) {
        const explanation = explainTranscriptError(err);
        if (explanation.code === 'PERMISSION_MISSING') {
          return {
            success: false,
            message: 'Autorisation Entra ID non accordée (403 Forbidden)',
            details: "Le consentement administrateur pour 'OnlineMeetingTranscript.Read.All' n'est pas encore actif.",
          };
        }
        if (explanation.code === 'GRAPH_ACCESS_DISABLED') {
          return {
            success: false,
            message: 'Accès API désactivé dans Teams Admin (403 Forbidden)',
            details: "Le paramètre EnableGraphTranscriptAccess doit être activé par l'administrateur Teams.",
          };
        }
        return {
          success: false,
          message: explanation.title || 'Vérification en échec',
          details: explanation.message,
        };
      }
    } finally {
      setIsCheckingTechnical(false);
    }
  }, [accounts, instance]);

  const handleCheckTranscript = async (
    meeting: MeetingItem,
    forceRefresh: boolean = false
  ): Promise<TranscriptCheckResult> => {
    const account = instance.getActiveAccount() || (accounts.length > 0 ? accounts[0] : null);
    if (!account) {
      return checkMeetingTranscriptStatus(null, meeting, forceRefresh, user, false);
    }
    const client = createTranscriptGraphClient(instance, account);
    return checkMeetingTranscriptStatus(client, meeting, forceRefresh, user, true);
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 flex flex-col font-sans">
      <Header
        activeTab={activeTab}
        setActiveTab={setActiveTab}
        user={user}
        config={msalConfig}
        onLoginClick={() => handleLogin(false)}
        onLogoutClick={handleLogout}
        isLoggingIn={isLoggingIn}
        adminConfig={adminConfig}
      />

      <main className="flex-1 p-4 sm:p-6 lg:p-8 max-w-7xl mx-auto w-full">
        {activeTab === 'auth_diagnostic' && (
          <AuthDiagnosticPanel
            config={msalConfig}
            user={user}
            onRefreshConfig={handleRefreshConfig}
            onLoginClick={() => handleLogin(false)}
            onLoginRedirectClick={() => handleLogin(true)}
            onTestGraphCall={handleTestGraphCall}
            isLoggingIn={isLoggingIn}
            isTestingGraph={isTestingGraph}
            lastError={lastError}
            graphRawResponse={graphRawResponse}
          />
        )}

        {activeTab === 'dashboard' && (
          <DashboardView
            user={user}
            meetings={meetings}
            isLoadingMeetings={isLoadingMeetings}
            onRefreshMeetings={() => handleLoadRealMeetings(calendarPeriod)}
            calendarPeriod={calendarPeriod}
            onChangeCalendarPeriod={handleLoadRealMeetings}
            onNavigateToAuth={() => setActiveTab('auth_diagnostic')}
            onSelectMeetingForReport={() => setActiveTab('reports')}
            onCheckTranscript={handleCheckTranscript}
            adminConfig={adminConfig}
            onUpdateAdminConfig={handleUpdateAdminConfig}
            onRunTechnicalCheck={handleRunAdminTechnicalCheck}
            isCheckingTechnical={isCheckingTechnical}
            workflowState={workflowState}
            onRequestConsent={handleTriggerAdminConsent}
            onVerifyConsent={handleVerifyAdminConsent}
            onUpdateWorkflowState={handleUpdateWorkflowState}
            isRequestingConsent={isRequestingConsent}
            isVerifyingConsent={isVerifyingConsent}
          />
        )}

        {activeTab === 'reports' && <ReportsLibraryView />}

        {activeTab === 'architecture' && <ArchitectureView />}
      </main>

      {/* Institutional Footer */}
      <footer className="border-t border-slate-900 bg-slate-950 px-6 py-4 text-xs text-slate-500">
        <div className="max-w-7xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <span className="font-semibold text-slate-400">ReadMeeting.ai</span>
            <span>•</span>
            <span>Direction Générale des Financements (DGF)</span>
            <span>•</span>
            <span>République de Côte d'Ivoire</span>
          </div>
          <div className="text-[11px] font-mono text-slate-600">
            Socle d'Intégration Microsoft Entra ID & Microsoft Graph v1.0
          </div>
        </div>
      </footer>
    </div>
  );
}
