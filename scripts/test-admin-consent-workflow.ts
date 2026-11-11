/**
 * Suite de tests unitaires et d'intégration :
 * Microsoft Entra ID Native Admin Consent Request Workflow
 * ReadMeeting.ai | Direction Générale des Financements (DGF)
 */

import {
  parseEntraConsentError,
  triggerNativeAdminConsentRequest,
  verifyEntraConsentStatus,
  getStoredAdminConsentState,
  saveAdminConsentState,
  updateAdminConsentState,
  resetAdminConsentState,
  ALL_REQUIRED_SCOPES,
  ADMIN_CONSENT_STORAGE_KEY
} from '../src/services/adminConsentService';
import {
  AdminConsentWorkflowStatus,
  ADMIN_CONSENT_WORKFLOW_META,
  DGF_ENTRA_PREREQUISITES
} from '../src/types/adminConsentWorkflow';

// Polyfill minimal de localStorage pour l'environnement Node/TSX
const memoryStore: Record<string, string> = {};
(global as any).localStorage = {
  getItem: (k: string) => memoryStore[k] || null,
  setItem: (k: string, v: string) => { memoryStore[k] = v; },
  removeItem: (k: string) => { delete memoryStore[k]; },
  clear: () => { Object.keys(memoryStore).forEach(k => delete memoryStore[k]); }
};

let passed = 0;
let failed = 0;

function assert(condition: boolean, testName: string, details?: string) {
  if (condition) {
    console.log(`  \x1b[32m✔\x1b[0m ${testName}`);
    passed++;
  } else {
    console.error(`  \x1b[31m✖\x1b[0m ${testName}`);
    if (details) console.error(`    \x1b[33mDétails: ${details}\x1b[0m`);
    failed++;
  }
}

console.log('\n========================================================================');
console.log('TEST SUITE : Microsoft Entra Native Admin Consent Request Workflow (DGF)');
console.log('========================================================================\n');

// 1. Validation des 7 états formels du workflow
console.log('1. Couverture des 7 états formels du workflow :');
const expectedStates: AdminConsentWorkflowStatus[] = [
  'admin_consent_required',
  'request_submitted',
  'decision_pending',
  'consent_granted',
  'request_declined',
  'workflow_unavailable',
  'verification_needed'
];

expectedStates.forEach(status => {
  const meta = ADMIN_CONSENT_WORKFLOW_META[status];
  assert(
    !!meta && typeof meta.label === 'string' && typeof meta.badgeClass === 'string',
    `État [${status}] configuré avec métadonnées et libellé institutionnel`,
    meta ? meta.label : 'Manquant'
  );
});

// 2. Validation des prérequis du locataire DGF
console.log('\n2. Validation des prérequis Entra ID et métadonnées DGF :');
assert(
  DGF_ENTRA_PREREQUISITES.tenantId === '9ddae6a3-6cb4-40e3-a4cd-305e2fe8c32a',
  'Tenant ID DGF rigoureusement configuré (9ddae6a3-6cb4-40e3-a4cd-305e2fe8c32a)'
);
assert(
  DGF_ENTRA_PREREQUISITES.permissionRequested === 'OnlineMeetingTranscript.Read.All',
  'Permission déléguée cible exacte : OnlineMeetingTranscript.Read.All'
);
assert(
  DGF_ENTRA_PREREQUISITES.requiredTenantSettings.length >= 4,
  `Paramètres de stratégie Entra documentés (${DGF_ENTRA_PREREQUISITES.requiredTenantSettings.length} paramètres)`
);
assert(
  ALL_REQUIRED_SCOPES.includes('OnlineMeetingTranscript.Read.All') &&
  ALL_REQUIRED_SCOPES.includes('User.Read') &&
  ALL_REQUIRED_SCOPES.includes('Calendars.Read') &&
  ALL_REQUIRED_SCOPES.includes('OnlineMeetings.Read'),
  'Portées complètes demandées cohérentes avec les flux Graph'
);

// 3. Analyse des codes erreurs Entra ID
console.log('\n3. Analyse des codes erreurs Entra ID (parseEntraConsentError) :');

const errAadsts90094Standard = {
  errorCode: 'AADSTS90094',
  message: 'The grant requires admin permission. User submitted justification in Entra ID dialog.',
  correlationId: 'test-corr-1'
};
const parsed90094 = parseEntraConsentError(errAadsts90094Standard);
assert(
  parsed90094.code === 'AADSTS90094' && parsed90094.isApprovalRequested === true,
  'AADSTS90094 standard identifié comme demande d\'approbation utilisateur soumise'
);

const errWorkflowDisabled = {
  errorCode: 'AADSTS90094',
  message: 'AADSTS90094: The grant requires admin permission. The approval workflow is not enabled in this directory.',
  correlationId: 'test-corr-2'
};
const parsedDisabled = parseEntraConsentError(errWorkflowDisabled);
assert(
  parsedDisabled.isWorkflowDisabled === true && parsedDisabled.isApprovalRequested === false,
  'Workflow désactivé sur le locataire détecté avec précision (workflow_unavailable)'
);

const errUserCancelled = {
  errorCode: 'user_cancelled',
  message: 'User cancelled the flow by closing the popup window.'
};
const parsedCancelled = parseEntraConsentError(errUserCancelled);
assert(
  parsedCancelled.isCancelled === true,
  'Annulation/fermeture par l\'utilisateur distinguée de l\'échec technique'
);

const errAadsts65001 = {
  errorCode: 'AADSTS65001',
  message: 'The user or administrator has not consented to use the application with ID...'
};
const parsed65001 = parseEntraConsentError(errAadsts65001);
assert(
  parsed65001.isDirectConsentRequired === true,
  'AADSTS65001 identifié comme consentement direct requis'
);

// 4. Persistence et mise à jour de l'état local
console.log('\n4. Persistance de l\'état (AdminConsentWorkflowState) :');
resetAdminConsentState();
let state = getStoredAdminConsentState();
assert(state.status === 'admin_consent_required', 'État initialisé à [admin_consent_required]');

updateAdminConsentState({
  status: 'request_submitted',
  requestedByUpn: 'agent.dgf@finances.gouv.ci',
  correlationId: 'corr-unit-test'
});
state = getStoredAdminConsentState();
assert(
  state.status === 'request_submitted' && state.requestedByUpn === 'agent.dgf@finances.gouv.ci',
  'Mise à jour atomique et persistance vers [request_submitted]'
);

// 5. Test du déclenchement natif interactif avec mock MSAL
console.log('\n5. Déclenchement de triggerNativeAdminConsentRequest() :');

// Mock 1 : Succès avec octroi de token immédiat (ex: compte administrateur)
const mockMsalAdminSuccess: any = {
  getActiveAccount: () => ({ username: 'admin.dgf@finances.gouv.ci', name: 'Admin DGF' }),
  acquireTokenPopup: async (req: any) => ({
    accessToken: 'mock_token',
    scopes: ['User.Read', 'OnlineMeetingTranscript.Read.All'],
    correlationId: 'corr-admin-granted'
  })
};

(async () => {
  const resultAdmin = await triggerNativeAdminConsentRequest(
    mockMsalAdminSuccess,
    mockMsalAdminSuccess.getActiveAccount(),
    'Justification test'
  );
  assert(
    resultAdmin.success === true && resultAdmin.status === 'consent_granted',
    'Administrateur validant le consentement : bascule immédiate en [consent_granted]'
  );

  // Mock 2 : Utilisateur non-admin avec workflow activé (AADSTS90094)
  resetAdminConsentState();
  const mockMsalUserApproval: any = {
    getActiveAccount: () => ({ username: 'collab.dgf@finances.gouv.ci', name: 'Collaborateur DGF' }),
    acquireTokenPopup: async () => {
      const err: any = new Error('AADSTS90094: The grant requires admin permission.');
      err.errorCode = 'AADSTS90094';
      err.correlationId = 'corr-approval-sent';
      throw err;
    }
  };

  const resultUser = await triggerNativeAdminConsentRequest(
    mockMsalUserApproval,
    mockMsalUserApproval.getActiveAccount(),
    'Besoin pour réunion DGF'
  );
  assert(
    resultUser.success === true && resultUser.status === 'request_submitted',
    'Non-admin avec workflow actif : confirmation de demande soumise [request_submitted]'
  );

  // Mock 3 : Locataire avec workflow désactivé
  resetAdminConsentState();
  const mockMsalDisabled: any = {
    getActiveAccount: () => ({ username: 'collab.dgf@finances.gouv.ci' }),
    acquireTokenPopup: async () => {
      const err: any = new Error('AADSTS90094: approval workflow is not enabled in this directory.');
      err.errorCode = 'AADSTS90094';
      throw err;
    }
  };

  const resultDisabled = await triggerNativeAdminConsentRequest(
    mockMsalDisabled,
    mockMsalDisabled.getActiveAccount()
  );
  assert(
    resultDisabled.status === 'workflow_unavailable',
    'Locataire avec workflow inactif : détection précise [workflow_unavailable]'
  );

  // 6. Test de la vérification silencieuse (verifyEntraConsentStatus)
  console.log('\n6. Vérification silencieuse (verifyEntraConsentStatus) :');

  // Cas A : Consentement accordé (token silencieux valide)
  const mockSilentSuccess: any = {
    getActiveAccount: () => ({ username: 'collab.dgf@finances.gouv.ci' }),
    acquireTokenSilent: async () => ({
      accessToken: 'valid_token',
      scopes: ['user.read', 'calendars.read', 'onlinemeetings.read', 'onlinemeetingtranscript.read.all'],
      correlationId: 'corr-silent-1'
    }),
    getClientConfig: () => ({ auth: { clientId: 'test' } })
  };

  const verifyResSuccess = await verifyEntraConsentStatus(
    mockSilentSuccess,
    mockSilentSuccess.getActiveAccount()
  );
  assert(
    verifyResSuccess.success === true && verifyResSuccess.status === 'consent_granted',
    'Vérification silencieuse positive : détection de la portée OnlineMeetingTranscript.Read.All -> [consent_granted]'
  );

  // Cas B : Consentement toujours manquant (InteractionRequired)
  updateAdminConsentState({ status: 'request_submitted' });
  const mockSilentPending: any = {
    getActiveAccount: () => ({ username: 'collab.dgf@finances.gouv.ci' }),
    acquireTokenSilent: async () => {
      const err: any = new Error('InteractionRequiredAuthError: AADSTS65001: The user or administrator has not consented');
      err.errorCode = 'consent_required';
      throw err;
    }
  };

  const verifyResPending = await verifyEntraConsentStatus(
    mockSilentPending,
    mockSilentPending.getActiveAccount()
  );
  assert(
    verifyResPending.success === false && verifyResPending.status === 'decision_pending',
    'Vérification silencieuse en attente : conservation exacte de [decision_pending] sans faux positif'
  );

  console.log('\n------------------------------------------------------------------------');
  console.log(`RÉSULTAT GLOBAL : ${passed} passés, ${failed} échoués`);
  console.log('------------------------------------------------------------------------\n');

  if (failed > 0) {
    process.exit(1);
  }
})();
