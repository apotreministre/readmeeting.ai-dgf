import { Configuration, LogLevel, PublicClientApplication } from '@azure/msal-browser';
import { MsalRuntimeConfig } from '../types/auth';

const STORAGE_KEY_CLIENT_ID = 'readmeeting_entra_client_id';
const STORAGE_KEY_TENANT_ID = 'readmeeting_entra_tenant_id';
const STORAGE_KEY_REDIRECT_URI = 'readmeeting_entra_redirect_uri';

export const OFFICIAL_PRODUCTION_ORIGIN = 'https://readmeeting-ai-dgf.ai.studio';
export const OFFICIAL_REDIRECT_PATH = '/redirect.html';
export const PRODUCTION_REDIRECT_URI = `${OFFICIAL_PRODUCTION_ORIGIN}${OFFICIAL_REDIRECT_PATH}`;

export interface RedirectUriValidationResult {
  isValid: boolean;
  sanitizedUri?: string;
  error?: string;
}

/**
 * Valide strictement une URI de redirection MSAL.
 * 
 * Règles appliquées :
 * 1. Analyse obligatoire avec `new URL(rawUri)`.
 * 2. Rejet des URI malformées.
 * 3. Rejet strict des protocoles non autorisés (seul HTTPS est admis, ou HTTP sur localhost/127.0.0.1).
 * 4. Exigence du chemin exact '/redirect.html' (url.pathname === '/redirect.html').
 * 5. Rejet des paramètres de requête (?search) et des fragments d'ancrage (#hash).
 * 6. Comparaison exacte de l'origine (url.origin === expectedOrigin) et JAMAIS par startsWith().
 * 7. Origines autorisées :
 *    - L'origine officielle de production DGF : https://readmeeting-ai-dgf.ai.studio
 *    - L'origine active explicitement prévue pour l'environnement d'exécution (activeOrigin ou window.location.origin).
 * 8. Rejet strict des domaines ressemblants (ex: readmeeting-ai-dgf.ai.studio.evil-example.com).
 * 9. Rejet des anciennes URI Cloud Run orphelines.
 */
export function validateMsalRedirectUri(
  rawUri: string | null | undefined,
  activeOrigin?: string
): RedirectUriValidationResult {
  if (!rawUri || typeof rawUri !== 'string' || !rawUri.trim()) {
    return { isValid: false, error: 'URI de redirection non fournie ou vide.' };
  }

  const trimmed = rawUri.trim();

  let parsed: URL;
  try {
    parsed = new URL(trimmed);
  } catch {
    return { isValid: false, error: "Format d'URI invalide (URL absolue requise)." };
  }

  // Protocole obligatoire : HTTPS (ou HTTP sur localhost de dev)
  const isLocalhost = parsed.hostname === 'localhost' || parsed.hostname === '127.0.0.1';
  if (parsed.protocol !== 'https:' && (!isLocalhost || parsed.protocol !== 'http:')) {
    return { isValid: false, error: 'Protocole invalide : HTTPS requis (ou HTTP sur localhost).' };
  }

  // Rejet strict des paramètres de recherche et fragments
  if (parsed.search && parsed.search.length > 0) {
    return { isValid: false, error: 'Paramètres de requête (?...) interdits dans le Redirect Bridge.' };
  }
  if (parsed.hash && parsed.hash.length > 0) {
    return { isValid: false, error: "Fragments (#...) interdits dans l'URI de redirection MSAL." };
  }

  // Rejet strict des identifiants utilisateur intégrés (username:password@)
  if (parsed.username || parsed.password) {
    return {
      isValid: false,
      error: "Les identifiants utilisateur sont interdits dans l'URI de redirection.",
    };
  }

  // Chemin exact : doit être strictement '/redirect.html'
  if (parsed.pathname !== OFFICIAL_REDIRECT_PATH) {
    return {
      isValid: false,
      error: `Chemin non conforme : attendu '${OFFICIAL_REDIRECT_PATH}', reçu '${parsed.pathname}'.`,
    };
  }

  // Origine courante active
  const currentOrigin =
    activeOrigin ||
    (typeof window !== 'undefined' ? window.location.origin : '');

  // Vérification stricte par égalité exacte (===)
  const isProduction = parsed.origin === OFFICIAL_PRODUCTION_ORIGIN;
  const isActiveEnv = Boolean(currentOrigin && parsed.origin === currentOrigin);

  if (!isProduction && !isActiveEnv) {
    return {
      isValid: false,
      error: `Origine non autorisée : '${parsed.origin}'. Seule la production DGF ('${OFFICIAL_PRODUCTION_ORIGIN}') ou l'origine active autorisée ('${currentOrigin}') sont acceptées.`,
    };
  }

  return {
    isValid: true,
    sanitizedUri: `${parsed.origin}${OFFICIAL_REDIRECT_PATH}`,
  };
}

export function getMsalRuntimeConfig(): MsalRuntimeConfig {
  const envClientId = (import.meta.env.VITE_AZURE_CLIENT_ID as string) || '';
  const envTenantId = (import.meta.env.VITE_AZURE_TENANT_ID as string) || '';
  const envRedirectUri = (import.meta.env.VITE_AZURE_REDIRECT_URI as string) || '';

  const storedClientId = typeof window !== 'undefined' ? localStorage.getItem(STORAGE_KEY_CLIENT_ID) : null;
  const storedTenantId = typeof window !== 'undefined' ? localStorage.getItem(STORAGE_KEY_TENANT_ID) : null;
  let storedRedirectUri = typeof window !== 'undefined' ? localStorage.getItem(STORAGE_KEY_REDIRECT_URI) : null;

  // Assainissement strict du localStorage par validation centralisée
  if (typeof window !== 'undefined' && storedRedirectUri) {
    const validation = validateMsalRedirectUri(storedRedirectUri, window.location.origin);
    if (!validation.isValid) {
      console.warn('[MSAL Config] Purge automatique d\'une URI invalide ou non autorisée dans localStorage :', storedRedirectUri, validation.error);
      localStorage.removeItem(STORAGE_KEY_REDIRECT_URI);
      storedRedirectUri = null;
    } else {
      storedRedirectUri = validation.sanitizedUri ?? null;
    }
  }

  const clientId = storedClientId || envClientId;
  const tenantId = storedTenantId || envTenantId;

  // Résolution prioritaire de l'URI de redirection :
  // 1. Variable d'environnement VITE_AZURE_REDIRECT_URI (si valide)
  // 2. Domaine de production stable si l'application s'exécute sur le domaine officiel
  // 3. URI stockée manuellement et validée
  // 4. Origine active du navigateur + /redirect.html
  // 5. Fallback officiel de production
  let redirectUri = '';
  if (envRedirectUri && envRedirectUri.trim().length > 0) {
    const envVal = validateMsalRedirectUri(
      envRedirectUri,
      typeof window !== 'undefined' ? window.location.origin : undefined
    );
    if (envVal.isValid && envVal.sanitizedUri) {
      redirectUri = envVal.sanitizedUri;
    } else {
      console.warn('[MSAL Config] VITE_AZURE_REDIRECT_URI rejetée car non conforme :', envRedirectUri, envVal.error);
    }
  }

  if (!redirectUri && typeof window !== 'undefined' && window.location.hostname === 'readmeeting-ai-dgf.ai.studio') {
    redirectUri = PRODUCTION_REDIRECT_URI;
  }

  if (!redirectUri && storedRedirectUri) {
    redirectUri = storedRedirectUri;
  }

  if (!redirectUri && typeof window !== 'undefined') {
    redirectUri = `${window.location.origin}${OFFICIAL_REDIRECT_PATH}`;
  }

  if (!redirectUri) {
    redirectUri = PRODUCTION_REDIRECT_URI;
  }

  const scopes = [
    'User.Read',
    'Calendars.Read',
    'OnlineMeetings.Read',
  ];

  const isConfigured = Boolean(
    clientId &&
    clientId.trim().length > 10 &&
    !clientId.includes('MY_') &&
    tenantId &&
    tenantId.trim().length > 10
  );

  return {
    clientId: clientId.trim(),
    tenantId: tenantId.trim(),
    redirectUri: redirectUri.trim(),
    scopes,
    isConfigured,
  };
}

export function saveMsalCredentials(
  clientId: string,
  tenantId: string,
  redirectUri?: string
): { success: boolean; error?: string } {
  if (typeof window !== 'undefined') {
    let sanitizedRedirectUri: string | null = null;

    // 1. Validation préalable stricte du redirectUri avant toute écriture
    if (redirectUri && redirectUri.trim()) {
      const validation = validateMsalRedirectUri(redirectUri.trim(), window.location.origin);
      if (!validation.isValid) {
        // En cas d'échec : arrêt immédiat, aucune des trois clés n'est modifiée
        return {
          success: false,
          error: validation.error || "L'URI de redirection spécifiée n'est pas autorisée.",
        };
      }
      sanitizedRedirectUri = validation.sanitizedUri!;
    }

    // 2. Écriture atomique dans localStorage uniquement après succès complet de la validation
    if (clientId) localStorage.setItem(STORAGE_KEY_CLIENT_ID, clientId.trim());
    else localStorage.removeItem(STORAGE_KEY_CLIENT_ID);

    if (tenantId) localStorage.setItem(STORAGE_KEY_TENANT_ID, tenantId.trim());
    else localStorage.removeItem(STORAGE_KEY_TENANT_ID);

    if (sanitizedRedirectUri) {
      localStorage.setItem(STORAGE_KEY_REDIRECT_URI, sanitizedRedirectUri);
    } else {
      localStorage.removeItem(STORAGE_KEY_REDIRECT_URI);
    }
  }
  return { success: true };
}

export function purgeStoredRedirectUri(): void {
  if (typeof window !== 'undefined') {
    localStorage.removeItem(STORAGE_KEY_REDIRECT_URI);
  }
}

export function purgeStoredMsalCredentials(): void {
  if (typeof window !== 'undefined') {
    localStorage.removeItem(STORAGE_KEY_CLIENT_ID);
    localStorage.removeItem(STORAGE_KEY_TENANT_ID);
    localStorage.removeItem(STORAGE_KEY_REDIRECT_URI);
  }
}

export function createMsalConfiguration(): Configuration {
  const config = getMsalRuntimeConfig();
  const authority = config.tenantId
    ? `https://login.microsoftonline.com/${config.tenantId}`
    : 'https://login.microsoftonline.com/common';

  return {
    auth: {
      clientId: config.clientId || '00000000-0000-0000-0000-000000000000',
      authority,
      redirectUri: config.redirectUri,
      postLogoutRedirectUri: config.redirectUri,
    },
    cache: {
      cacheLocation: 'sessionStorage',
    },
    system: {
      loggerOptions: {
        loggerCallback: (level: LogLevel, message: string, containsPii: boolean) => {
          if (containsPii) return;
          if (level === LogLevel.Error) {
            console.error('[MSAL Error]', message);
          } else if (level === LogLevel.Warning) {
            console.warn('[MSAL Warning]', message);
          }
        },
        logLevel: LogLevel.Warning,
      },
    },
  };
}

// Instance unique singleton et verrou d'initialisation asynchrone
let msalInstance: PublicClientApplication | null = null;
let msalInitPromise: Promise<PublicClientApplication> | null = null;

/**
 * Initialise de manière asynchrone et idempotente l'instance unique de PublicClientApplication.
 * Garantit que initialize() n'est appelé et attendu qu'une seule fois, même en cas d'appels simultanés.
 */
export async function getInitializedMsalInstance(): Promise<PublicClientApplication> {
  if (msalInstance) {
    return msalInstance;
  }

  if (msalInitPromise) {
    return msalInitPromise;
  }

  const config = createMsalConfiguration();
  const instance = new PublicClientApplication(config);

  msalInitPromise = (async () => {
    try {
      await instance.initialize();
      msalInstance = instance;
      return instance;
    } catch (error) {
      // En cas d'échec, libérer la promesse pour permettre une réinitialisation propre
      msalInitPromise = null;
      throw error;
    }
  })();

  return msalInitPromise;
}

/**
 * Accès synchrone sécurisé : retourne l'instance si elle a déjà été initialisée,
 * ou déclenche une exception explicite pour empêcher l'erreur uninitialized_public_client_application.
 */
export function getMsalInstance(): PublicClientApplication {
  if (!msalInstance) {
    throw new Error(
      "Tentative d'accès à l'instance MSAL avant son initialisation asynchrone. Utilisez getInitializedMsalInstance() ou le hook useMsal()."
    );
  }
  return msalInstance;
}

/**
 * Réinitialise proprement l'instance MSAL avec une nouvelle configuration et attend son initialisation complète.
 * Évite toute multiplication d'instances en écrasant la référence unique seulement une fois initialisée.
 */
export async function resetMsalInstance(): Promise<PublicClientApplication> {
  const config = createMsalConfiguration();
  const newInstance = new PublicClientApplication(config);

  msalInitPromise = (async () => {
    try {
      await newInstance.initialize();
      msalInstance = newInstance;
      return newInstance;
    } catch (error) {
      msalInitPromise = null;
      throw error;
    }
  })();

  return msalInitPromise;
}

export const loginRequest = {
  scopes: ['User.Read', 'Calendars.Read', 'OnlineMeetings.Read'],
};

export const graphTranscriptRequest = {
  scopes: ['OnlineMeetings.Read', 'OnlineMeetingTranscript.Read.All'],
};
