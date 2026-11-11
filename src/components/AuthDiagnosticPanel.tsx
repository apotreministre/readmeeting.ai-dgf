import React, { useState } from 'react';
import {
  Shield,
  Key,
  Copy,
  Check,
  AlertTriangle,
  CheckCircle2,
  ExternalLink,
  RefreshCw,
  Info,
  Server,
  Lock,
  Globe,
  Database
} from 'lucide-react';
import { DgfUserProfile, MsalRuntimeConfig } from '../types/auth';
import {
  saveMsalCredentials,
  purgeStoredRedirectUri,
  purgeStoredMsalCredentials,
  PRODUCTION_REDIRECT_URI,
} from '../config/msalConfig';
import { GraphErrorExplanation } from '../services/graphClient';

interface AuthDiagnosticPanelProps {
  config: MsalRuntimeConfig;
  user: DgfUserProfile | null;
  onRefreshConfig: () => void | Promise<void>;
  onLoginClick: () => void;
  onLoginRedirectClick?: () => void;
  onTestGraphCall: () => Promise<void>;
  isLoggingIn: boolean;
  isTestingGraph: boolean;
  lastError: GraphErrorExplanation | null;
  graphRawResponse: any;
}

export const AuthDiagnosticPanel: React.FC<AuthDiagnosticPanelProps> = ({
  config,
  user,
  onRefreshConfig,
  onLoginClick,
  onTestGraphCall,
  isLoggingIn,
  isTestingGraph,
  lastError,
  graphRawResponse,
}) => {
  const [copiedField, setCopiedField] = useState<string | null>(null);
  const [isEditing, setIsEditing] = useState<boolean>(!config.isConfigured);
  const [customClientId, setCustomClientId] = useState<string>(config.clientId);
  const [customTenantId, setCustomTenantId] = useState<string>(config.tenantId);
  const [customRedirectUri, setCustomRedirectUri] = useState<string>(config.redirectUri);
  const [saveSuccess, setSaveSuccess] = useState<boolean>(false);
  const [purgeSuccess, setPurgeSuccess] = useState<boolean>(false);
  const [saveError, setSaveError] = useState<string | null>(null);

  const currentOrigin = typeof window !== 'undefined' ? window.location.origin : '';
  const currentDevRedirectUri = `${currentOrigin}/redirect.html`;

  const copyToClipboard = (text: string, fieldName: string) => {
    navigator.clipboard.writeText(text);
    setCopiedField(fieldName);
    setTimeout(() => setCopiedField(null), 2000);
  };

  const handleSaveCredentials = (e: React.FormEvent) => {
    e.preventDefault();
    setSaveError(null);
    const targetUri = customRedirectUri.trim() || currentDevRedirectUri;
    const saveResult = saveMsalCredentials(customClientId, customTenantId, targetUri);
    if (!saveResult.success) {
      setSaveError(saveResult.error || "L'URI de redirection n'est pas autorisée.");
      return;
    }
    onRefreshConfig();
    setIsEditing(false);
    setSaveSuccess(true);
    setTimeout(() => setSaveSuccess(false), 3000);
  };

  const handlePurgeCache = () => {
    purgeStoredRedirectUri();
    onRefreshConfig();
    setPurgeSuccess(true);
    setTimeout(() => setPurgeSuccess(false), 3000);
  };

  const checks = [
    {
      id: 'origin',
      title: 'Origine Web & Chiffrement HTTPS',
      status: currentOrigin.startsWith('https://') || currentOrigin.includes('localhost') ? 'success' : 'warning',
      description: currentOrigin,
      detail: 'Requis par Microsoft Entra ID pour la sécurisation du flux OAuth 2.0 PKCE.',
    },
    {
      id: 'client_id',
      title: "Identifiant d'Application (Client ID)",
      status: config.clientId && config.clientId.length > 20 ? 'success' : 'error',
      description: config.clientId ? `${config.clientId.slice(0, 8)}...${config.clientId.slice(-4)}` : 'Non renseigné',
      detail: "Généré lors de l'enregistrement de l'application dans le portail Microsoft Entra.",
    },
    {
      id: 'tenant_id',
      title: "Identifiant de l'Annuaire (Tenant ID)",
      status: config.tenantId && config.tenantId.length > 20 ? 'success' : 'error',
      description: config.tenantId ? `${config.tenantId.slice(0, 8)}...${config.tenantId.slice(-4)}` : 'Non renseigné',
      detail: 'Définit le périmètre du locataire DGF (login.microsoftonline.com/{tenantId}).',
    },
    {
      id: 'redirect_uri',
      title: 'URI Redirect Bridge MSAL v5 (Active)',
      status: config.redirectUri.endsWith('/redirect.html') ? 'success' : 'warning',
      description: config.redirectUri,
      detail: 'Doit impérativement cibler redirect.html et être déclarée sous SPA dans Entra ID.',
    },
    {
      id: 'scopes',
      title: 'Périmètre des Permissions Microsoft Graph',
      status: 'success',
      description: config.scopes.join(', '),
      detail: 'User.Read, Calendars.Read, OnlineMeetings.Read.',
    },
  ];

  return (
    <div className="space-y-6 max-w-7xl mx-auto pb-12">
      {/* Top Banner Warning & Policy */}
      <div className="bg-gradient-to-r from-blue-950/60 to-slate-900 border border-blue-800/40 rounded-xl p-5 shadow-lg">
        <div className="flex items-start gap-4">
          <div className="w-10 h-10 rounded-lg bg-blue-900/60 border border-blue-700/50 flex items-center justify-center shrink-0">
            <Lock className="w-5 h-5 text-blue-300" />
          </div>
          <div className="space-y-1">
            <div className="flex items-center gap-2">
              <h2 className="text-base font-semibold text-white">Console de Diagnostic & Validation Microsoft Entra ID</h2>
              <span className="px-2 py-0.5 text-[10px] uppercase font-bold tracking-wider bg-blue-900/80 text-blue-300 border border-blue-700/60 rounded">
                Flux Réel PKCE
              </span>
            </div>
            <p className="text-xs text-slate-300 leading-relaxed">
              Conformément aux directives de la DGF, <strong>aucune simulation n'est effectuée</strong>. L'authentification s'exécute directement auprès de la plateforme d'identité Microsoft. Si vos identifiants réels sont configurés, vous pouvez déclencher la connexion interactive et valider l'acquisition de token en direct.
            </p>
          </div>
        </div>
      </div>

      {/* Grid: 2 columns */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left Column (2 spans): Configuration & Live Test */}
        <div className="lg:col-span-2 space-y-6">
          {/* Diagnostic Checks List */}
          <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-5 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <h3 className="text-sm font-semibold text-slate-200 flex items-center gap-2">
                <Shield className="w-4 h-4 text-emerald-400" />
                Vérification des Prérequis Techniques
              </h3>
              <div className="flex items-center gap-3">
                <button
                  onClick={handlePurgeCache}
                  title="Supprimer les anciennes URI mémorisées dans le navigateur"
                  className="text-xs text-slate-400 hover:text-amber-400 flex items-center gap-1 transition-colors"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                  Purger cache local
                </button>
                <button
                  onClick={onRefreshConfig}
                  className="text-xs text-slate-400 hover:text-emerald-400 flex items-center gap-1 transition-colors"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                  Actualiser
                </button>
              </div>
            </div>

            {purgeSuccess && (
              <div className="p-3 bg-amber-950/70 border border-amber-800/80 rounded-lg text-xs text-amber-300 flex items-center gap-2">
                <CheckCircle2 className="w-4 h-4 text-amber-400 shrink-0" />
                Le cache local des URI a été purgé avec succès. L'URI active utilise désormais la configuration prioritaire.
              </div>
            )}

            <div className="divide-y divide-slate-800/60">
              {checks.map((check) => (
                <div key={check.id} className="py-3 flex items-start justify-between gap-4">
                  <div className="space-y-0.5">
                    <div className="flex items-center gap-2">
                      <span className="text-xs font-medium text-slate-200">{check.title}</span>
                      {check.status === 'success' && (
                        <span className="text-[10px] font-medium text-emerald-400 flex items-center gap-1">
                          <CheckCircle2 className="w-3 h-3" /> Conforme
                        </span>
                      )}
                      {check.status === 'warning' && (
                        <span className="text-[10px] font-medium text-amber-400 flex items-center gap-1">
                          <AlertTriangle className="w-3 h-3" /> À surveiller
                        </span>
                      )}
                      {check.status === 'error' && (
                        <span className="text-[10px] font-medium text-rose-400 flex items-center gap-1">
                          <AlertTriangle className="w-3 h-3" /> Requis
                        </span>
                      )}
                    </div>
                    <p className="text-xs font-mono text-slate-400 break-all">{check.description}</p>
                    <p className="text-[11px] text-slate-500">{check.detail}</p>
                  </div>

                  {check.id === 'redirect_uri' && (
                    <div className="flex flex-col gap-1 items-end shrink-0">
                      <button
                        onClick={() => copyToClipboard(check.description, 'redirect_uri')}
                        className="px-2.5 py-1 text-[11px] font-medium text-slate-300 bg-slate-800 hover:bg-slate-700 rounded border border-slate-700 flex items-center gap-1 transition-all"
                      >
                        {copiedField === 'redirect_uri' ? (
                          <>
                            <Check className="w-3 h-3 text-emerald-400" />
                            Copié
                          </>
                        ) : (
                          <>
                            <Copy className="w-3 h-3" />
                            Copier URI active
                          </>
                        )}
                      </button>
                    </div>
                  )}
                </div>
              ))}
            </div>

            {/* Extra callout for Production vs Dev URI */}
            <div className="mt-3 p-3.5 bg-slate-950/70 rounded-lg border border-slate-800/80 space-y-2">
              <span className="text-[11px] font-semibold text-slate-300 uppercase tracking-wider block">
                URIs à déclarer dans Microsoft Entra ID (Plateforme SPA) :
              </span>
              <div className="space-y-1.5 text-xs font-mono">
                <div className="flex items-center justify-between gap-2 p-1.5 bg-slate-900 rounded border border-slate-800">
                  <div className="truncate">
                    <span className="text-emerald-400 font-semibold mr-1.5">[Production DGF] :</span>
                    <span className="text-slate-300">{PRODUCTION_REDIRECT_URI}</span>
                  </div>
                  <button
                    onClick={() => copyToClipboard(PRODUCTION_REDIRECT_URI, 'prod_uri')}
                    className="px-2 py-0.5 text-[10px] text-slate-300 bg-slate-800 hover:bg-slate-700 rounded border border-slate-700 shrink-0"
                  >
                    {copiedField === 'prod_uri' ? 'Copié !' : 'Copier'}
                  </button>
                </div>
                <div className="flex items-center justify-between gap-2 p-1.5 bg-slate-900 rounded border border-slate-800">
                  <div className="truncate">
                    <span className="text-blue-400 font-semibold mr-1.5">[Prévisualisation Dev] :</span>
                    <span className="text-slate-400">{currentDevRedirectUri}</span>
                  </div>
                  <button
                    onClick={() => copyToClipboard(currentDevRedirectUri, 'dev_uri')}
                    className="px-2 py-0.5 text-[10px] text-slate-300 bg-slate-800 hover:bg-slate-700 rounded border border-slate-700 shrink-0"
                  >
                    {copiedField === 'dev_uri' ? 'Copié !' : 'Copier'}
                  </button>
                </div>
              </div>
            </div>
          </div>

          {/* Credentials Setup / Live Modifier */}
          <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-5 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <div>
                <h3 className="text-sm font-semibold text-slate-200 flex items-center gap-2">
                  <Key className="w-4 h-4 text-blue-400" />
                  Paramètres de Connexion au Tenant DGF
                </h3>
                <p className="text-xs text-slate-400 mt-0.5">
                  Ces valeurs proviennent de votre App Registration dans Microsoft Entra ID.
                </p>
              </div>
              <button
                onClick={() => {
                  setIsEditing(!isEditing);
                  setSaveError(null);
                }}
                className="px-3 py-1 text-xs font-medium text-slate-300 bg-slate-800 hover:bg-slate-700 rounded-md border border-slate-700 transition-colors"
              >
                {isEditing ? 'Fermer' : 'Modifier les identifiants'}
              </button>
            </div>

            {saveSuccess && (
              <div className="p-3 bg-emerald-950/80 border border-emerald-800/80 rounded-lg text-xs text-emerald-300 flex items-center gap-2">
                <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
                Identifiants enregistrés avec succès dans le navigateur pour cette session de test.
              </div>
            )}

            {saveError && (
              <div className="p-3 bg-rose-950/80 border border-rose-800/80 rounded-lg text-xs text-rose-300 flex items-start gap-2">
                <AlertTriangle className="w-4 h-4 text-rose-400 shrink-0 mt-0.5" />
                <div>
                  <span className="font-semibold block">URI de redirection rejetée :</span>
                  {saveError}
                </div>
              </div>
            )}

            {isEditing ? (
              <form onSubmit={handleSaveCredentials} className="space-y-4 pt-1">
                <div>
                  <label className="block text-xs font-medium text-slate-300 mb-1">
                    Application (client) ID <span className="text-rose-400">*</span>
                  </label>
                  <input
                    type="text"
                    required
                    placeholder="ex: 11111111-2222-3333-4444-555555555555"
                    value={customClientId}
                    onChange={(e) => setCustomClientId(e.target.value)}
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-xs font-mono text-slate-100 focus:outline-none focus:border-blue-500"
                  />
                  <span className="text-[11px] text-slate-500">Visible sur la page Vue d'ensemble de l'application Entra ID.</span>
                </div>

                <div>
                  <label className="block text-xs font-medium text-slate-300 mb-1">
                    Directory (tenant) ID <span className="text-rose-400">*</span>
                  </label>
                  <input
                    type="text"
                    required
                    placeholder="ex: aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"
                    value={customTenantId}
                    onChange={(e) => setCustomTenantId(e.target.value)}
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-xs font-mono text-slate-100 focus:outline-none focus:border-blue-500"
                  />
                  <span className="text-[11px] text-slate-500">ID du tenant DGF Microsoft 365.</span>
                </div>

                <div>
                  <label className="block text-xs font-medium text-slate-300 mb-1">
                    URI de Redirection MSAL (Optionnel - défaut automatique)
                  </label>
                  <input
                    type="text"
                    placeholder={`ex: ${PRODUCTION_REDIRECT_URI}`}
                    value={customRedirectUri}
                    onChange={(e) => {
                      setCustomRedirectUri(e.target.value);
                      if (saveError) setSaveError(null);
                    }}
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-xs font-mono text-slate-100 focus:outline-none focus:border-blue-500"
                  />
                  <span className="text-[11px] text-slate-500">
                    Doit être l'URI officielle DGF (<code>{PRODUCTION_REDIRECT_URI}</code>) ou l'origine de dev autorisée (<code>{currentDevRedirectUri}</code>).
                  </span>
                </div>

                <div className="flex items-center gap-2 pt-2">
                  <button
                    type="submit"
                    className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg text-xs font-medium transition-all shadow-md shadow-blue-950/40"
                  >
                    Appliquer et Mettre à jour MSAL
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setIsEditing(false);
                      setSaveError(null);
                    }}
                    className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg text-xs font-medium transition-colors"
                  >
                    Annuler
                  </button>
                </div>
              </form>
            ) : (
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                <div className="p-3 bg-slate-950 rounded-lg border border-slate-800">
                  <span className="text-slate-500 block mb-1">Client ID actif</span>
                  <span className="font-mono text-slate-300 break-all">
                    {config.clientId || 'Non configuré (cliquez sur Modifier)'}
                  </span>
                </div>
                <div className="p-3 bg-slate-950 rounded-lg border border-slate-800">
                  <span className="text-slate-500 block mb-1">Tenant ID actif</span>
                  <span className="font-mono text-slate-300 break-all">
                    {config.tenantId || 'Non configuré (cliquez sur Modifier)'}
                  </span>
                </div>
              </div>
            )}
          </div>

          {/* Error Deciphering Block (if an error occurred) */}
          {lastError && (
            <div className="bg-rose-950/40 border border-rose-800/80 rounded-xl p-5 space-y-3">
              <div className="flex items-start gap-3">
                <AlertTriangle className="w-5 h-5 text-rose-400 shrink-0 mt-0.5" />
                <div className="space-y-1">
                  <div className="flex items-center gap-2">
                    <h4 className="text-xs font-bold text-rose-200">{lastError.title}</h4>
                    <span className="px-1.5 py-0.5 text-[10px] font-mono bg-rose-900/60 text-rose-300 rounded border border-rose-700/50">
                      {lastError.code}
                    </span>
                  </div>
                  <p className="text-xs text-rose-300 leading-relaxed">{lastError.description}</p>
                </div>
              </div>

              <div className="p-3 bg-rose-900/30 rounded-lg border border-rose-800/50 text-xs text-rose-200">
                <span className="font-semibold block text-white mb-0.5">Action corrective requise :</span>
                {lastError.actionRequired}
              </div>
            </div>
          )}

          {/* Raw Microsoft Graph Response Viewer */}
          {graphRawResponse && (
            <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-5 space-y-2">
              <div className="flex items-center justify-between">
                <h4 className="text-xs font-semibold text-slate-300 flex items-center gap-2">
                  <Database className="w-3.5 h-3.5 text-emerald-400" />
                  Données Brutes Renvoyées par Microsoft Graph (/me)
                </h4>
                <span className="text-[10px] text-emerald-400 font-medium">Réponse HTTP 200 OK</span>
              </div>
              <pre className="p-3 bg-slate-950 rounded-lg border border-slate-800 text-[11px] font-mono text-emerald-300 overflow-x-auto max-h-56">
                {JSON.stringify(graphRawResponse, null, 2)}
              </pre>
            </div>
          )}
        </div>

        {/* Right Column (1 span): Live Actions & Admin Guide */}
        <div className="space-y-6">
          {/* Action Card : Connexion & Test */}
          <div className="bg-gradient-to-b from-slate-900 to-slate-950 border border-slate-800 rounded-xl p-5 space-y-4">
            <h3 className="text-sm font-semibold text-white flex items-center gap-2">
              <Key className="w-4 h-4 text-emerald-400" />
              Exécution du Test d'Authentification
            </h3>

            <p className="text-xs text-slate-400 leading-relaxed">
              Cliquez ci-dessous pour ouvrir la mire d'authentification Microsoft officielle (pop-up sécurisée avec échange de code PKCE).
            </p>

            <div className="space-y-2 pt-1">
              <button
                onClick={onLoginClick}
                disabled={isLoggingIn || !config.isConfigured}
                className="w-full flex items-center justify-center gap-2 px-4 py-2.5 text-xs font-semibold rounded-lg bg-blue-600 hover:bg-blue-500 text-white shadow-lg shadow-blue-950/50 transition-all border border-blue-400/20 disabled:opacity-50 disabled:cursor-not-allowed"
              >
                <Key className="w-4 h-4" />
                {isLoggingIn
                  ? 'Négociation du token en cours...'
                  : user
                  ? 'Renouveler le token Entra ID'
                  : 'Se Connecter avec Microsoft Entra ID'}
              </button>

              {user && (
                <button
                  onClick={onTestGraphCall}
                  disabled={isTestingGraph}
                  className="w-full flex items-center justify-center gap-2 px-4 py-2 text-xs font-medium rounded-lg bg-emerald-700 hover:bg-emerald-600 text-white transition-colors disabled:opacity-50"
                >
                  <RefreshCw className={`w-3.5 h-3.5 ${isTestingGraph ? 'animate-spin' : ''}`} />
                  {isTestingGraph ? 'Interrogation Microsoft Graph...' : 'Tester la requête Graph (/me)'}
                </button>
              )}
            </div>

            {!config.isConfigured && (
              <div className="p-3 bg-amber-950/50 border border-amber-800/60 rounded-lg text-xs text-amber-300 flex items-start gap-2">
                <Info className="w-4 h-4 text-amber-400 shrink-0 mt-0.5" />
                <div>
                  <span className="font-semibold block">Configuration requise :</span>
                  Renseignez votre Client ID et Tenant ID dans la section ci-contre pour activer le bouton de connexion.
                </div>
              </div>
            )}
          </div>

          {/* User Profile Card (if connected) */}
          {user && (
            <div className="bg-slate-900/90 border border-emerald-800/60 rounded-xl p-5 space-y-3">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-full bg-emerald-900/80 border border-emerald-600 flex items-center justify-center text-emerald-200 font-bold text-sm">
                  {user.displayName.slice(0, 2).toUpperCase()}
                </div>
                <div>
                  <h4 className="text-xs font-bold text-white">{user.displayName}</h4>
                  <p className="text-[11px] text-slate-400">{user.userPrincipalName}</p>
                </div>
              </div>

              <div className="border-t border-slate-800 pt-3 space-y-2 text-xs">
                <div className="flex justify-between">
                  <span className="text-slate-500">Poste / Rôle :</span>
                  <span className="text-slate-300 font-medium">{user.jobTitle || 'Cadre DGF'}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-500">Direction :</span>
                  <span className="text-slate-300 font-medium">{user.department || 'DGF'}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-500">Localisation :</span>
                  <span className="text-slate-300 font-medium">{user.officeLocation || 'Abidjan'}</span>
                </div>
              </div>
            </div>
          )}

          {/* Quick Setup Checklist for DGF Admin */}
          <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-5 space-y-3">
            <h4 className="text-xs font-semibold text-slate-200 flex items-center gap-2">
              <Info className="w-3.5 h-3.5 text-blue-400" />
              Rappel Procédure Admin M365 (DGF)
            </h4>

            <ol className="text-xs text-slate-400 space-y-2 list-decimal list-inside leading-relaxed">
              <li>
                Aller sur <strong className="text-slate-200">portal.azure.com</strong> &gt; Inscriptions d'applications.
              </li>
              <li>
                Sélectionner la plateforme <strong className="text-emerald-400">SPA</strong> (Single-page application).
              </li>
              <li>
                Enregistrer l'URI de redirection officielle de production (Redirect Bridge) :
                <div className="mt-1 pl-4">
                  <code className="text-[11px] text-emerald-300 bg-slate-950 px-1.5 py-0.5 rounded border border-slate-800 break-all select-all block">
                    {PRODUCTION_REDIRECT_URI}
                  </code>
                </div>
                <div className="mt-1 pl-4 text-[11px] text-slate-500">
                  Et pour vos tests en prévisualisation dev :
                  <code className="text-[10px] text-slate-400 bg-slate-950 px-1 py-0.5 rounded border border-slate-800 break-all select-all block mt-0.5">
                    {currentDevRedirectUri}
                  </code>
                </div>
              </li>
              <li>
                Vérifier les 3 permissions déléguées avec consentement admin accordé :
                <div className="mt-1 pl-4 text-[11px] text-slate-300">
                  <span className="font-mono bg-slate-950 px-1 py-0.5 rounded">User.Read</span>,{' '}
                  <span className="font-mono bg-slate-950 px-1 py-0.5 rounded">Calendars.Read</span>,{' '}
                  <span className="font-mono bg-slate-950 px-1 py-0.5 rounded">OnlineMeetings.Read</span>.
                </div>
              </li>
            </ol>
          </div>
        </div>
      </div>
    </div>
  );
};
